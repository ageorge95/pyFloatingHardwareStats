from PySide6.QtWidgets import (QApplication,
                               QWidget,
                               QVBoxLayout,
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
import sys
import psutil
import win32gui
import win32con
import wmi
import os

def value_to_rgb_to_QTableWidgetItem(value, min_value, max_value):
    # Function that returns a QTableWidgetItem used to color the table cells in a certain manner
    # Ensure value is within the range [min_value, max_value]
    value = max(min(value, max_value), min_value)

    # Calculate the ratio of the value within the range
    ratio = (value - min_value) / (max_value - min_value)

    # Interpolate between green (0, 255, 0) and red (255, 0, 0)
    red = int(ratio * 255)  # Red increases as the value increases
    green = int((1 - ratio) * 255)  # Green decreases as the value increases

    # return the RGB value
    return (red, green, 0)

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
        data_storage |= {'ram_usage': round(ram_usage,2),
                         'ram_total': round(ram_total,2)}

        sleep(0.5)

def libre_hw_mon_updater(data_storage: dict):
    while not os.path.isfile(get_running_path('exit')):
        try:
            # Connect to LibreHardwareMonitor's WMI namespace
            w = wmi.WMI(namespace="root\\LibreHardwareMonitor")
            query = 'SELECT Value FROM Sensor WHERE SensorType="Temperature" AND Name="CPU Package"'
            results = w.query(query)
            data_storage |= {'CPU_temp': results[0].Value}
        except Exception as e:
            data_storage |= {'CPU_temp': 0}
            print('Querying LibreHardwareMonitor failed !')

        sleep(0.25)

# Worker thread to update stats
class StatsUpdater(QThread):
    stats_updated = Signal(list, list)  # Signal to send updated stats to the main thread

    def __init__(self):
        super().__init__()

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

        self.libre_hw_mon = {'CPU_temp': 0}
        t_libre_hw_mon = Thread(target=libre_hw_mon_updater, args=(self.libre_hw_mon,))
        t_libre_hw_mon.start()

    def run(self):
        while not os.path.isfile(get_running_path('exit')):

            # Collect all the data points in a list
            rows = [[f"CPU[%]: {self.cpu_usage['cpu_percent']}",
                     f"CPU[C]: {self.libre_hw_mon['CPU_temp']}"],
                    [f"RAM[GB]: {self.RAM_stats['ram_usage']} / {self.RAM_stats['ram_total']}",
                     f"RAM[%]: {round((self.RAM_stats['ram_usage'] / self.RAM_stats['ram_total']) * 100, 2)}"],
                    [f"Network⬆️ [MB]: {self.network_stats['download_speed_history_MB'][-1]}",
                     f"Network⬇️ [MB]: {self.network_stats['upload_speed_history_MB'][-1]}"]]

            colors = [[value_to_rgb_to_QTableWidgetItem(self.cpu_usage['cpu_percent'], 0, 100),
                       value_to_rgb_to_QTableWidgetItem(self.libre_hw_mon['CPU_temp'], 40, 90)],
                      [value_to_rgb_to_QTableWidgetItem(self.RAM_stats['ram_usage'],0,self.RAM_stats['ram_total']),
                       value_to_rgb_to_QTableWidgetItem(self.RAM_stats['ram_usage'], 0, self.RAM_stats['ram_total'])],
                      [value_to_rgb_to_QTableWidgetItem(self.network_stats['download_speed_history_MB'][-1],
                                                        0,
                                                        max(self.network_stats['download_speed_history_MB'])),
                       value_to_rgb_to_QTableWidgetItem(self.network_stats['upload_speed_history_MB'][-1],
                                                        0,
                                                        max(self.network_stats['upload_speed_history_MB']))]
                      ]

            # Emit formatted data
            self.stats_updated.emit(rows, colors)  # Emit signal with formatted rows and their colors
            self.msleep(500)  # Sleep for 500ms before updating again

def get_running_path(relative_path):
    if '_internal' in os.listdir():
        return os.path.join('_internal', relative_path)
    else:
        return relative_path

class DraggableWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # first clear the exit flag by removing the exit file if it exists
        if os.path.isfile(get_running_path('exit')):
            os.remove(get_running_path('exit'))

        # Set up the window properties
        self.setWindowTitle("System Stats v" + open(get_running_path('version.txt')).read())
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

        # these position coordinates will be used to keep the main window exactly where the user drags it
        # see the logic from move_window_to_fixed_position
        self.dragged_x_pos = 0
        self.dragged_y_pos = 0

        # Create a table-like visualization with labels
        # The previous implementation with QTableWidget was not ok as
        # the rows height could not be customized beyond certain limits
        self._cells = {}
        for column_index in range(3):  # 3 columns
            column_layout = QVBoxLayout()  # Vertical layout for a single column
            column_layout.setSpacing(0)  # Remove spacing between labels
            column_layout.setContentsMargins(0, 0, 0, 0)  # Remove margins around the column

            for row_index in range(2):  # Five rows per column
                label = QLabel(f"Row {row_index}, Col {column_index}: TBD")
                # the cells are stored with the <col_nr>_<row_nr> keys
                self._cells[f'{column_index}_{row_index}'] = label

                label.setStyleSheet("color: black; padding: 0px; margin: 0px;")

                # Add label to column
                column_layout.addWidget(label)

                # Add the column to the main layout
                grid_layout.addLayout(column_layout, row_index, column_index)

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

    def update_table(self, rows, colors):

        for (col_idx, row), (_, color) in zip(enumerate(rows), enumerate(colors)):
            for (row_idx, cell_data), (_, color) in zip(enumerate(row), enumerate(color)):
                self._cells[f'{col_idx}_{row_idx}'].setText(cell_data)
                self._cells[f'{col_idx}_{row_idx}'].setStyleSheet(f"background-color: rgb{color};")

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
        # Simply create an exit file which can be seen by all thethreads so that they can close gracefully
        with open(get_running_path('exit'), 'w') as _:
            pass

# Run the application
app = QApplication(sys.argv)
window = DraggableWindow()
window.show()
sys.exit(app.exec())
