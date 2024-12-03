import sys
import psutil
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QFrame
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem
from PySide6.QtCore import Qt, QTimer, Signal, QThread
from PySide6.QtGui import QMouseEvent
import win32gui
import win32con
from tabulate import tabulate

# Worker thread to update stats
class StatsUpdater(QThread):
    stats_updated = Signal(list)  # Signal to send updated stats to the main thread

    def run(self):
        while True:
            cpu_percent = psutil.cpu_percent(interval=0.5)
            ram_usage = psutil.virtual_memory().used / (1024 ** 3)  # in GB
            ram_total = psutil.virtual_memory().total / (1024 ** 3)  # in GB
            net_io = psutil.net_io_counters()
            sent_GB = f"{net_io.bytes_sent / (1024 ** 3):.2f}"
            received_GB = f"{net_io.bytes_recv / (1024 ** 3):.2f}"

            # Collect all the data points in a list
            rows = [[f"CPU: {cpu_percent}%", f"RAM: {ram_usage:.2f} GB / {ram_total:.2f} GB", 'TBA'],
                    ['TBA', f"Network: {sent_GB}GB Sent/ {received_GB}GB Received", 'TBA']]

            # Emit formatted data
            self.stats_updated.emit(rows)  # Emit signal with formatted rows
            self.msleep(500)  # Sleep for 500ms before updating again

            print('called')


class DraggableWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set up the window properties
        self.setWindowTitle("System Stats")
        self.setGeometry(100, 100, 500, 30)  # Initial position and size
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)  # Always on top, no frame
        self.setStyleSheet("background-color: rgba(255, 255, 255, 220);")  # Light transparent background

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
        self.table_widget.setColumnCount(3)  # 3 columns per row
        # Remove headers
        self.table_widget.horizontalHeader().setVisible(False)
        self.table_widget.verticalHeader().setVisible(False)
        # Set all columns and rows to be compact
        self.table_widget.setSizeAdjustPolicy(QTableWidget.AdjustToContents)
        self.table_widget.setColumnWidth(0, 70)
        self.table_widget.setColumnWidth(1, 220)
        self.table_widget.setColumnWidth(2, 80)
        self.table_widget.setRowHeight(0, 1)  # Adjust the height of each row
        self.table_widget.setRowHeight(1, 1)
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

    def update_table(self, rows):
        # Clear the table before updating
        self.table_widget.clearContents()

        print(rows)

        # Update the table with new data
        row_count = len(rows)
        self.table_widget.setRowCount(row_count)

        for row_idx, row in enumerate(rows):
            for col_idx, cell_data in enumerate(row):
                item = QTableWidgetItem(cell_data)
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
