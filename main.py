from PySide6.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QFrame
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
from PySide6.QtCore import Qt, QTimer, Signal, QThread
from PySide6.QtGui import QMouseEvent, QColor, QFont
from threading import Thread
from time import sleep
import sys
import psutil
import win32gui
import win32con
import wmi

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
    while True:
        cpu_percent = psutil.cpu_percent(interval=0.5)
        data_storage['cpu_percent'] = cpu_percent

        # no sleep() needed here as cpu_percent() already takes up 0.5s

def network_speed_updater(data_storage: dict):
    sampling_interval_s = 0.75
    while True:
        initial_stats = psutil.net_io_counters()
        initial_bytes_sent = initial_stats.bytes_sent
        initial_bytes_recv = initial_stats.bytes_recv

        sleep(sampling_interval_s)

        final_stats = psutil.net_io_counters()
        final_bytes_sent = final_stats.bytes_sent
        final_bytes_recv = final_stats.bytes_recv

        download_speed_MB = (final_bytes_recv - initial_bytes_recv) / sampling_interval_s / 1024**2
        upload_speed_MB = (final_bytes_sent - initial_bytes_sent) / sampling_interval_s / 1024**2

        data_storage |= {'download_speed_MB': round(download_speed_MB,3),
                         'upload_speed_MB': round(upload_speed_MB,3)}

        # no sleep() needed here as cpu_percent() already takes up sampling_interval_s

def RAM_stats_updater(data_storage: dict):
    while True:
        ram_usage = psutil.virtual_memory().used / (1024 ** 3)  # in GB
        ram_total = psutil.virtual_memory().total / (1024 ** 3)  # in GB
        data_storage |= {'ram_usage': round(ram_usage,2),
                         'ram_total': round(ram_total,2)}

        sleep(0.5)

def libre_hw_mon_updater(data_storage: dict):
    while True:
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

        self.network_stats = {'download_speed_MB': 0.000,
                              'upload_speed_MB': 0.000}
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
        while True:

            # Collect all the data points in a list
            rows = [[f"CPU[%]: {self.cpu_usage['cpu_percent']}%",
                     f"RAM: {self.RAM_stats['ram_usage']} GB / {self.RAM_stats['ram_total']} GB"],
                    [f"CPU[C]: {self.libre_hw_mon['CPU_temp']}C",
                     f"Network: {self.network_stats['download_speed_MB']}MB Down/ {self.network_stats['upload_speed_MB']}MB Up"]]

            colors = [[value_to_rgb_to_QTableWidgetItem(self.cpu_usage['cpu_percent'],0,100),
                       value_to_rgb_to_QTableWidgetItem(self.RAM_stats['ram_usage'],0,self.RAM_stats['ram_total'])],
                      [value_to_rgb_to_QTableWidgetItem(self.libre_hw_mon['CPU_temp'], 40, 90),
                       (255,255,255)]]

            # Emit formatted data
            self.stats_updated.emit(rows, colors)  # Emit signal with formatted rows and their colors
            self.msleep(500)  # Sleep for 500ms before updating again

class DraggableWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the window properties
        self.setWindowTitle("System Stats")
        self.setGeometry(100, 100, 325, 30)  # Initial position and size
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)  # Always on top, no frame
        self.setStyleSheet("background-color: rgba(255, 255, 255, 220);")  # Light transparent background
        self.setWindowOpacity(0.8)  # 80% opaque

        # Create the layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)  # Remove the outer margins
        layout.setSpacing(0)  # Remove the space between the widgets

        # Make the window draggable
        self.drag_frame = QFrame(self)
        self.drag_frame.setStyleSheet("background-color: gray;")
        self.drag_frame.setFixedHeight(5)
        layout.addWidget(self.drag_frame)
        self.drag_frame.mousePressEvent = self.start_drag
        self.drag_frame.mouseMoveEvent = self.do_drag

        # Create the QTableWidget (Excel-like grid)
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(2)  # Initially setting 2 rows
        self.table_widget.setColumnCount(2)  # 3 columns per row
        font = QFont()
        font.setBold(True)
        self.table_widget.setFont(font) # set the default bold text
        # Remove headers
        self.table_widget.horizontalHeader().setVisible(False)
        self.table_widget.verticalHeader().setVisible(False)
        # Set all columns and rows to be compact
        self.table_widget.setSizeAdjustPolicy(QTableWidget.AdjustToContents)
        self.table_widget.setColumnWidth(0, 90) # adjust the rows width
        self.table_widget.setColumnWidth(1, 230)
        self.table_widget.setRowHeight(0, 1)  # Adjust the height of each row
        self.table_widget.setRowHeight(1, 1)
        # set all items to have the bold font
        layout.addWidget(self.table_widget)

        # Timer to keep the window always on top
        self.keep_on_top_timer = QTimer(self)
        self.keep_on_top_timer.timeout.connect(self.ensure_window_above_taskbar)
        self.keep_on_top_timer.start(100)  # Ensure window stays on top every 100ms

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
        # Clear the table before updating
        self.table_widget.clearContents()

        # Update the table with new data
        row_count = len(rows)
        self.table_widget.setRowCount(row_count)

        for (row_idx, row), (_, color) in zip(enumerate(rows), enumerate(colors)):
            for (col_idx, cell_data), (_, color) in zip(enumerate(row), enumerate(color)):
                item = QTableWidgetItem(cell_data)
                item.setBackground(QColor(*color))
                self.table_widget.setItem(row_idx, col_idx, item)

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

        # Allow window to move over the taskbar (no screen boundary restriction)
        self.move(new_x, new_y)

    def ensure_window_above_taskbar(self):
        hwnd = self.winId()
        # Ensure the window stays on top of the taskbar
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)


# Run the application
app = QApplication(sys.argv)
window = DraggableWindow()
window.show()
sys.exit(app.exec())
