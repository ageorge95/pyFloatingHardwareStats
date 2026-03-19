from PySide6.QtWidgets import (QApplication,
                               QWidget,
                               QFrame,
                               QLabel,
                               QGridLayout,
                               QMainWindow)
from PySide6.QtCore import (Qt,
                            QTimer,
                            Signal,
                            QThread)
from PySide6.QtGui import (QMouseEvent,
                           QIcon,
                           QCloseEvent)
from threading import Thread
from time import sleep
from ag95 import red_green_from_range_value
import sys
import psutil
import win32gui
import win32con
import os
import json
import requests

LIBRE_HARDWARE_MONITOR_PORT = 8085

def CPU_usage_updater(data_storage: dict):
    while not os.path.isfile(get_running_path('exit')):
        cpu_percent = psutil.cpu_percent(interval=0.5)
        data_storage['cpu_percent'] = cpu_percent

        # no sleep() needed here as cpu_percent() already takes up 0.5s

def network_speed_updater(data_storage: dict):
    sampling_interval_s = 0.75
    while not os.path.isfile(get_running_path('exit')):
        initial_stats = psutil.net_io_counters()
        initial_bytes_sent = initial_stats.bytes_sent
        initial_bytes_recv = initial_stats.bytes_recv

        sleep(sampling_interval_s)

        final_stats = psutil.net_io_counters()
        final_bytes_sent = final_stats.bytes_sent
        final_bytes_recv = final_stats.bytes_recv

        download_speed_MB = (final_bytes_recv - initial_bytes_recv) / sampling_interval_s / 1024**2
        upload_speed_MB = (final_bytes_sent - initial_bytes_sent) / sampling_interval_s / 1024**2

        # add the new values to the history
        data_storage['download_speed_history_MB'].append(round(download_speed_MB,3))
        data_storage['upload_speed_history_MB'].append(round(upload_speed_MB, 3))

        # trim the history to only contain 1000 entries
        data_storage['download_speed_history_MB'] = data_storage['download_speed_history_MB'][-1000:]
        data_storage['upload_speed_history_MB'] = data_storage['upload_speed_history_MB'][-1000:]

def RAM_stats_updater(data_storage: dict):
    while not os.path.isfile(get_running_path('exit')):
        ram_usage = psutil.virtual_memory().used / (1024 ** 3)  # in GB
        ram_total = psutil.virtual_memory().total / (1024 ** 3)  # in GB
        data_storage |= {'ram_usage': round(ram_usage,1),
                         'ram_total': round(ram_total,1)}

        sleep(0.5)

def libre_hw_mon_updater(data_storage: dict):
    """
    Use LibreHardwareMonitor's web server JSON endpoint.
    Polls the HTTP endpoint at 127.0.0.1:<LIBRE_HARDWARE_MONITOR_PORT>/data.json every 2 seconds.
    """
    while not os.path.isfile(get_running_path('exit')):
        try:
            # Make HTTP request to LibreHardwareMonitor web server
            response = requests.get(f'http://127.0.0.1:{LIBRE_HARDWARE_MONITOR_PORT}/data.json', timeout=5)
            response.raise_for_status()  # Raise exception for bad status codes
            data = response.json()

            # Initialize with defaults
            cpu_temp = 0
            igpu_temp = 0
            igpu_usage = 0
            dgpu_temp = 0
            dgpu_usage = 0
            disk1_activity = 0
            disk1_read_speed = 0
            disk1_write_speed = 0
            disk2_activity = 0
            disk2_read_speed = 0
            disk2_write_speed = 0

            # Helper function to search for sensors in the nested JSON structure
            def find_sensor(data_node, sensor_type, name_filter=None, hardware_id_filter=None):
                """Recursively search for sensors in the JSON tree"""
                results = []

                # Check if this node is a sensor
                if 'Type' in data_node:
                    if data_node['Type'] == sensor_type:
                        if name_filter is None or (name_filter and name_filter in data_node.get('Text', '')):
                            if hardware_id_filter is None or (
                                    hardware_id_filter and hardware_id_filter in data_node.get('HardwareId', '')):
                                results.append(data_node)

                # Recursively search children
                for child in data_node.get('Children', []):
                    results.extend(find_sensor(child, sensor_type, name_filter, hardware_id_filter))

                return results

            # Helper function to find a hardware node by its exact "Text" field
            def find_hardware_node(data_node, text_to_find):
                """Recursively search the JSON tree for a hardware node with a specific Text value."""
                if data_node.get('Text') == text_to_find and 'HardwareId' in data_node:
                    # Check if it's a hardware device and not just a category
                    if data_node.get('ImageURL', '').startswith('images_icon/'):
                        return data_node
                for child in data_node.get('Children', []):
                    found = find_hardware_node(child, text_to_find)
                    if found:
                        return found
                return None

            # Helper function to extract "Total Activity" from a disk node
            def get_disk_activity(disk_node):
                if not disk_node:
                    return 0
                activity_sensors = find_sensor(disk_node, 'Load', 'Total Activity')
                if activity_sensors:
                    value_str = activity_sensors[0].get('Value', '0')
                    try:
                        return float(''.join(c for c in value_str.split()[0] if c.isdigit() or c == '.'))
                    except (ValueError, AttributeError):
                        pass
                return 0

            # Helper function to get disk throughput in MB/s
            def get_disk_throughput(disk_node, rate_name):
                if not disk_node:
                    return 0
                throughput_sensors = find_sensor(disk_node, 'Throughput', rate_name)
                if throughput_sensors:
                    # Use RawValue which is in B/s for consistent conversion
                    value_str = throughput_sensors[0].get('RawValue', '0')
                    try:
                        # Value is like "307710.3 B/s", split and take the numeric part
                        bytes_per_second = float(value_str.split()[0])
                        megabytes_per_second = bytes_per_second / (1024 ** 2)
                        return megabytes_per_second
                    except (ValueError, AttributeError, IndexError):
                        pass
                return 0

            # Find 'disk1' and 'disk2' and get their stats
            disk1_node = find_hardware_node(data, 'disk1')
            disk1_activity = get_disk_activity(disk1_node)
            disk1_read_speed = get_disk_throughput(disk1_node, "Read Rate")
            disk1_write_speed = get_disk_throughput(disk1_node, "Write Rate")

            disk2_node = find_hardware_node(data, 'disk2')
            disk2_activity = get_disk_activity(disk2_node)
            disk2_read_speed = get_disk_throughput(disk2_node, "Read Rate")
            disk2_write_speed = get_disk_throughput(disk2_node, "Write Rate")

            # Find CPU temperature (CPU Package or SoC)
            cpu_temp_sensors = find_sensor(data, 'Temperature', 'CPU Package')
            if not cpu_temp_sensors:
                cpu_temp_sensors = find_sensor(data, 'Temperature', 'SoC')

            if cpu_temp_sensors:
                # Extract numeric value from "Value" field (e.g., "64.0 °C" -> 64.0)
                value_str = cpu_temp_sensors[0].get('Value', '0')
                # Remove non-numeric characters except decimal point
                try:
                    cpu_temp = float(''.join(c for c in value_str.split()[0] if c.isdigit() or c == '.'))
                except (ValueError, AttributeError):
                    cpu_temp = 0

            # Find dedicated GPU (NVIDIA) temperature
            dgpu_temp_sensors = find_sensor(data, 'Temperature', 'GPU Core')
            # Filter for NVIDIA GPU by checking parent HardwareId
            if dgpu_temp_sensors:
                # We need to check if this is actually the NVIDIA GPU
                # usually NVIDIA GPU is under hardware ID "/gpu-nvidia/0"
                nvidia_dgpu_temp = [s for s in dgpu_temp_sensors
                                    if s.get('SensorId', '').startswith('/gpu-nvidia')]
                if nvidia_dgpu_temp:
                    value_str = nvidia_dgpu_temp[0].get('Value', '0')
                    try:
                        dgpu_temp = float(''.join(c for c in value_str.split()[0] if c.isdigit() or c == '.'))
                    except (ValueError, AttributeError):
                        dgpu_temp = 0

            # Find dedicated GPU (NVIDIA) usage
            dgpu_load_sensors = find_sensor(data, 'Load', 'GPU Core')
            if dgpu_load_sensors:
                nvidia_dgpu_load = [s for s in dgpu_load_sensors
                                    if s.get('SensorId', '').startswith('/gpu-nvidia')]
                if nvidia_dgpu_load:
                    value_str = nvidia_dgpu_load[0].get('Value', '0')
                    try:
                        dgpu_usage = float(''.join(c for c in value_str.split()[0] if c.isdigit() or c == '.'))
                    except (ValueError, AttributeError):
                        dgpu_usage = 0

            # Find integrated GPU (Intel) usage - D3D 3D load
            # usually Intel GPU is under "/gpu-intel-integrated/"
            igpu_load_sensors = find_sensor(data, 'Load', 'D3D 3D')
            if igpu_load_sensors:
                intel_igpu_load = [s for s in igpu_load_sensors
                                   if '/gpu-intel-integrated/' in s.get('SensorId', '')]
                if intel_igpu_load:
                    value_str = intel_igpu_load[0].get('Value', '0')
                    try:
                        igpu_usage = float(''.join(c for c in value_str.split()[0] if c.isdigit() or c == '.'))
                    except (ValueError, AttributeError):
                        igpu_usage = 0

            # Integrated GPU temperature - fallback to CPU temperature
            igpu_temp = cpu_temp

            # Update data storage including history for disk speeds
            data_storage['CPU_temp'] = int(cpu_temp)
            data_storage['iGPU_temp'] = int(igpu_temp)
            data_storage['iGPU_usage'] = round(igpu_usage, 2)
            data_storage['dGPU_temp'] = int(dgpu_temp)
            data_storage['dGPU_usage'] = round(dgpu_usage, 2)
            data_storage['disk1_activity'] = round(disk1_activity, 2)
            data_storage['disk1_read_speed'] = round(disk1_read_speed, 2)
            data_storage['disk1_write_speed'] = round(disk1_write_speed, 2)
            data_storage['disk2_activity'] = round(disk2_activity, 2)
            data_storage['disk2_read_speed'] = round(disk2_read_speed, 2)
            data_storage['disk2_write_speed'] = round(disk2_write_speed, 2)

            # Append to history and trim
            for key in ['disk1_read_speed', 'disk1_write_speed', 'disk2_read_speed', 'disk2_write_speed']:
                history_key = f"{key}_history_MBs"
                data_storage[history_key].append(data_storage[key])
                data_storage[history_key] = data_storage[history_key][-1000:]

        except requests.exceptions.RequestException as e:
            # On error, reset current values but keep history
            data_storage.update({'CPU_temp': 0, 'iGPU_temp': 0, 'iGPU_usage': 0, 'dGPU_temp': 0, 'dGPU_usage': 0,
                                 'disk1_activity': 0, 'disk1_read_speed': 0, 'disk1_write_speed': 0,
                                 'disk2_activity': 0, 'disk2_read_speed': 0, 'disk2_write_speed': 0})
            print(f'LibreHardwareMonitor HTTP request failed: {e}')
        except (json.JSONDecodeError, Exception) as e:
            # On error, reset current values but keep history
            data_storage.update({'CPU_temp': 0, 'iGPU_temp': 0, 'iGPU_usage': 0, 'dGPU_temp': 0, 'dGPU_usage': 0,
                                 'disk1_activity': 0, 'disk1_read_speed': 0, 'disk1_write_speed': 0,
                                 'disk2_activity': 0, 'disk2_read_speed': 0, 'disk2_write_speed': 0})
            print(f'LibreHardwareMonitor error: {e}')

        sleep(2.0)  # Keep the same relaxed polling interval

# Worker thread to update stats
class StatsUpdater(QThread):
    stats_updated = Signal(list, list, list)  # Signal to send updated stats to the main thread

    def __init__(self):
        super().__init__()

        # previously emitted rows
        self._last_rows = None
        self._last_colors = None

        self.cpu_usage = {'cpu_percent': 0}
        t_CPU = Thread(target=CPU_usage_updater, args=(self.cpu_usage,))
        t_CPU.start()

        self.network_stats = {'download_speed_history_MB': [0.001],
                              'upload_speed_history_MB': [0.001]}
        t_network = Thread(target=network_speed_updater, args=(self.network_stats,))
        t_network.start()

        self.RAM_stats = {'ram_usage': 0,
                          'ram_total': 0}
        t_RAM = Thread(target=RAM_stats_updater, args=(self.RAM_stats,))
        t_RAM.start()

        self.libre_hw_mon = {'CPU_temp': 0, 'iGPU_temp': 0, 'iGPU_usage': 0, 'dGPU_temp': 0, 'dGPU_usage': 0,
                             'disk1_activity': 0, 'disk1_read_speed': 0, 'disk1_write_speed': 0,
                             'disk2_activity': 0, 'disk2_read_speed': 0, 'disk2_write_speed': 0,
                             'disk1_read_speed_history_MBs': [0.001], 'disk1_write_speed_history_MBs': [0.001],
                             'disk2_read_speed_history_MBs': [0.001], 'disk2_write_speed_history_MBs': [0.001]}
        t_libre_hw_mon = Thread(target=libre_hw_mon_updater, args=(self.libre_hw_mon,))
        t_libre_hw_mon.start()

    def run(self):
        while not os.path.isfile(get_running_path('exit')):
            ram_percent = round((self.RAM_stats['ram_usage'] / self.RAM_stats['ram_total']) * 100, 1) if self.RAM_stats['ram_total'] > 0 else 0

            # Collect all the data points in a 4x4 grid
            rows = [
                [f"CPU[%]: {self.cpu_usage['cpu_percent']}", f"RAM[%]: {ram_percent}", f"iGPU[%]: {self.libre_hw_mon['iGPU_usage']}", f"dGPU[%]: {self.libre_hw_mon['dGPU_usage']}"],
                [f"CPU[C]: {self.libre_hw_mon['CPU_temp']}", f"RAM[GB]: {self.RAM_stats['ram_usage']}", f"iGPU[C]: {self.libre_hw_mon['iGPU_temp']}", f"dGPU[C]: {self.libre_hw_mon['dGPU_temp']}"],
                [f"NET⬆️: {self.network_stats['upload_speed_history_MB'][-1]}", f"D1[%]]: {self.libre_hw_mon['disk1_activity']}", f"D1_R[MB\s]: {self.libre_hw_mon['disk1_read_speed']}", f"D1_W[MB\s]: {self.libre_hw_mon['disk1_write_speed']}"],
                [f"NET⬇️: {self.network_stats['download_speed_history_MB'][-1]}", f"D2[%]: {self.libre_hw_mon['disk2_activity']}", f"D2_R[MB\s]: {self.libre_hw_mon['disk2_read_speed']}", f"D2_W[MB\s]: {self.libre_hw_mon['disk2_write_speed']}"]
            ]

            colors = [
                [red_green_from_range_value(self.cpu_usage['cpu_percent'], 0, 100),
                 red_green_from_range_value(self.RAM_stats['ram_usage'], 0, self.RAM_stats['ram_total']),
                 red_green_from_range_value(self.libre_hw_mon['iGPU_usage'], 0, 100),
                 red_green_from_range_value(self.libre_hw_mon['dGPU_usage'], 0, 100)],
                [red_green_from_range_value(self.libre_hw_mon['CPU_temp'], 40, 90),
                 red_green_from_range_value(self.RAM_stats['ram_usage'], 0, self.RAM_stats['ram_total']),
                 red_green_from_range_value(self.libre_hw_mon['iGPU_temp'], 40, 90),
                 red_green_from_range_value(self.libre_hw_mon['dGPU_temp'], 40, 90)],
                [red_green_from_range_value(self.network_stats['upload_speed_history_MB'][-1], 0, max(self.network_stats['upload_speed_history_MB'])),
                 red_green_from_range_value(self.libre_hw_mon['disk1_activity'], 0, 100),
                 red_green_from_range_value(self.libre_hw_mon['disk1_read_speed'], 0, max(self.libre_hw_mon['disk1_read_speed_history_MBs'])),
                 red_green_from_range_value(self.libre_hw_mon['disk1_write_speed'], 0, max(self.libre_hw_mon['disk1_write_speed_history_MBs']))],
                [red_green_from_range_value(self.network_stats['download_speed_history_MB'][-1], 0, max(self.network_stats['download_speed_history_MB'])),
                 red_green_from_range_value(self.libre_hw_mon['disk2_activity'], 0, 100),
                 red_green_from_range_value(self.libre_hw_mon['disk2_read_speed'], 0, max(self.libre_hw_mon['disk2_read_speed_history_MBs'])),
                 red_green_from_range_value(self.libre_hw_mon['disk2_write_speed'], 0, max(self.libre_hw_mon['disk2_write_speed_history_MBs']))]
            ]

            # Emit formatted data
            if self._last_rows is None:
                changed = [[True] * len(row) for row in rows]
            else:
                changed = [[rows[r][c] != self._last_rows[r][c] or
                            colors[r][c] != self._last_colors[r][c]
                            for c in range(len(rows[r]))] for r in range(len(rows))]

            self._last_rows  = [row[:] for row in rows] # deep copy
            self._last_colors = [row[:] for row in colors]

            self.stats_updated.emit(rows, colors, changed)
            self.msleep(500)

def get_running_path(relative_path):
    if '_internal' in os.listdir():
        return os.path.join('_internal', relative_path)
    else:
        return relative_path

class DraggableWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # these position coordinates will be used to keep the main window exactly where the user drags it
        # see the logic from move_window_to_fixed_position for details
        self.dragged_x_pos = 0
        self.dragged_y_pos = 0

        # first clear the exit flag by removing the exit file if it exists
        # but also try to read the last window position on close
        if os.path.isfile(get_running_path('exit')):
            try:
                with open(get_running_path('exit'), 'r') as file_in_handle:
                    last_position = json.load(file_in_handle)
                    self.dragged_x_pos = last_position['dragged_x_pos']
                    self.dragged_y_pos = last_position['dragged_y_pos']
            except:
                pass

            os.remove(get_running_path('exit'))

        # Set up the window properties
        self.setWindowTitle("pyFloatingHardwareStats v" + open(get_running_path('version.txt')).read())
        self.setGeometry(100, 100, 1, 1)  # Initial position and size
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)  # Always on top, no frame
        self.setStyleSheet("background-color: rgba(255, 255, 255, 220);")  # Light transparent background
        self.setWindowOpacity(0.8)  # 80% opaque
        self.setWindowIcon(QIcon(get_running_path('icon.ico')))

        # Central widget and layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Create the main grid layout for labels
        grid_layout = QGridLayout(central_widget)
        grid_layout.setSpacing(0)
        grid_layout.setContentsMargins(0, 0, 0, 0)

        # Make the window draggable
        self.drag_frame = QFrame(self)
        self.drag_frame.setStyleSheet("background-color: gray;")
        self.drag_frame.setFixedHeight(5)
        # This frame should not be part of the grid_layout, instead it can sit on top
        self.setMenuWidget(self.drag_frame)
        self.drag_frame.mousePressEvent = self.start_drag
        self.drag_frame.mouseMoveEvent = self.do_drag

        # Create a table-like visualization with labels
        # The previous implementation with QTableWidget was not ok as
        # the rows height could not be customized beyond certain limits
        self._cells = {}
        for row_index in range(4):  # 4 rows
            for column_index in range(4):  # 4 columns
                label = QLabel(f"R{row_index}C{column_index}")
                # the cells are stored with the <row_nr>_<col_nr> keys
                self._cells[f'{row_index}_{column_index}'] = label

                label.setStyleSheet("color: black; "
                                    "padding: 0px; "
                                    "margin: 0px; "
                                    "font-family: Arial; "
                                    "font-size: 10px; "
                                    "font-weight: bold;"
                                    )
                # Add label directly to the grid layout
                grid_layout.addWidget(label, row_index, column_index)

        # Timer to keep the window always on top
        self.keep_on_top_timer = QTimer(self)
        self.keep_on_top_timer.timeout.connect(self.ensure_window_above_taskbar)
        self.keep_on_top_timer.start(100)  # Ensure window stays on top every 100ms

        # Timer to move the window to the last user position
        self.move_window_to_fixed_position_timer = QTimer(self)
        self.move_window_to_fixed_position_timer.timeout.connect(self.move_window_to_fixed_position)
        self.move_window_to_fixed_position_timer.start(2000)  # Ensure the window moves every 2s

        # Variables for drag functionality
        self.start_x = 0
        self.start_y = 0

        # Get screen geometry to allow dragging over taskbar
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        self.screen_width = screen_geometry.width()
        self.screen_height = screen_geometry.height()

        # Start the stats updater thread
        self.stats_updater = StatsUpdater()
        self.stats_updater.stats_updated.connect(self.update_table)
        self.stats_updater.start()

    def update_table(self, rows, colors, changed):
        # Data is structured as a 4x4 grid (list of rows)
        for r, (row_data, color_row, changed_row) in enumerate(zip(rows, colors, changed)):
            for c, (text, colour, changed_flag) in enumerate(zip(row_data, color_row, changed_row)):
                if changed_flag:
                    label = self._cells.get(f'{r}_{c}')
                    if label:
                        label.setText(text)
                        r_val, g_val, b_val = colour
                        label.setStyleSheet("color: black; "
                                            "padding: 0px; "
                                            "margin: 0px; "
                                            "font-family: Arial; "
                                            "font-size: 10px; "
                                            "font-weight: bold; "
                                            f"background-color: rgb({r_val}, {g_val}, {b_val});")

    def start_drag(self, event: QMouseEvent):
        # Record the current position of the window and mouse
        self.start_x = event.globalPosition().x()
        self.start_y = event.globalPosition().y()

    def do_drag(self, event: QMouseEvent):
        # Calculate the new position based on the mouse movement
        delta_x = event.globalPosition().x() - self.start_x
        delta_y = event.globalPosition().y() - self.start_y

        # Update the starting position to prevent the window from "jumping"
        self.start_x = event.globalPosition().x()
        self.start_y = event.globalPosition().y()

        # Calculate the new position
        new_x = self.x() + delta_x
        new_y = self.y() + delta_y

        self.dragged_x_pos = new_x
        self.dragged_y_pos = new_y

        # Allow window to move over the taskbar (no screen boundary restriction)
        self.move(new_x, new_y)

    def ensure_window_above_taskbar(self):
        hwnd = self.winId()
        # Ensure the window stays on top of the taskbar
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def move_window_to_fixed_position(self):
         # move the window to its latest user position
         # very handy when the window is reset by windows - for example when reconnection to remote machines via RDP
         if self.dragged_x_pos and self.dragged_y_pos:
            self.move(self.dragged_x_pos,self.dragged_y_pos)

    def closeEvent(self, event: QCloseEvent):
        # Custom logic to run when the window is closed
        # Simply create an exit file which can be seen by all the threads so that they can close gracefully
        # this file will contain the last known position of the main window
        with open(get_running_path('exit'), 'w') as file_out_handle:
            json.dump({'dragged_x_pos': self.dragged_x_pos,
                           'dragged_y_pos': self.dragged_y_pos}, file_out_handle)

# Run the application
app = QApplication(sys.argv)
window = DraggableWindow()
window.show()
sys.exit(app.exec())
