import sys
import os
import ctypes
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar,
    QDialog, QFormLayout, QLineEdit
)
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import Qt

# Using the updated excel_processor that references the new algorithm
from excel_processor import ExcelProcessor

def get_absolute_path(filename):
    if getattr(sys, '_MEIPASS', False):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.abspath(filename)

def set_taskbar_icon(window, icon_path):
    hwnd = int(window.winId())
    hicon = ctypes.windll.user32.LoadImageW(0, icon_path, 1, 0, 0, 0x00000010)
    ctypes.windll.user32.SendMessageW(hwnd, 0x80, 1, hicon)

class SettingsDialog(QDialog):
    def __init__(self, parent=None, current_values=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setGeometry(400, 300, 350, 500)
        self.setStyleSheet("background-color: #f4f4f4; border-radius: 10px;")

        self.layout = QVBoxLayout()

        title_label = QLabel("Modify Processing Variables", self)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #003366;")
        self.layout.addWidget(title_label)

        self.variables = {}
        self.variable_names = [
            "Multiplier", "Offset", "Threshold", "Scaling Factor", "Precision",
            "Adjustment", "Limit", "Ratio", "Modifier", "Factor"
        ]

        form_layout = QFormLayout()
        for var_name in self.variable_names:
            self.variables[var_name] = QLineEdit(self)
            self.variables[var_name].setText(str(int(current_values.get(var_name, 1))))
            self.variables[var_name].setStyleSheet("""
                QLineEdit {
                    border: 1px solid #bbb;
                    border-radius: 5px;
                    padding: 5px;
                    font-size: 14px;
                }
                QLineEdit:focus {
                    border: 2px solid #003366;
                }
            """)
            form_layout.addRow(var_name, self.variables[var_name])

        self.layout.addLayout(form_layout)

        self.save_button = QPushButton("Save", self)
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #FFC72C;
                color: black;
                font-size: 14px;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E6B800;
            }
        """)
        self.save_button.clicked.connect(self.accept)
        self.layout.addWidget(self.save_button, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(self.layout)

    def get_values(self):
        values = {}
        for key in self.variables:
            text_value = self.variables[key].text().strip()
            if text_value == "":
                values[key] = 1
                continue
            try:
                num = float(text_value)
                values[key] = int(num) if num.is_integer() else num
            except ValueError:
                values[key] = 1
        return values

class ExcelParserApp(QWidget):
    def __init__(self):
        super().__init__()
        self.settings_values = {var_name: 1 for var_name in [
            "Multiplier", "Offset", "Threshold", "Scaling Factor", "Precision",
            "Adjustment", "Limit", "Ratio", "Modifier", "Factor"
        ]}
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Lumberjack Balancing™")
        self.setGeometry(300, 200, 500, 500)
        icon_path = get_absolute_path("favicon.ico")
        self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.logo_label = QLabel(self)
        logo_path = get_absolute_path("Logo.png")
        pixmap = QPixmap(logo_path)
        if not pixmap.isNull():
            self.logo_label.setPixmap(pixmap.scaled(200, 200,
                                        Qt.AspectRatioMode.KeepAspectRatio,
                                        Qt.TransformationMode.SmoothTransformation))
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(self.logo_label, alignment=Qt.AlignmentFlag.AlignCenter)

        self.label = QLabel("Select the Excel files to process:")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        # ---------------------------
        # ADDED FOR MULTI-FILE SUPPORT
        # ---------------------------
        self.btn_select_raw = QPushButton("Upload Raw Data File")
        self.btn_select_raw.setStyleSheet("""
            QPushButton {
                background-color: #003366;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #00274D;
            }
        """)
        self.btn_select_raw.clicked.connect(self.select_raw_file)
        layout.addWidget(self.btn_select_raw)

        self.btn_select_policy = QPushButton("Select Policy File")
        self.btn_select_policy.setStyleSheet("""
            QPushButton {
                background-color: #003366;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #00274D;
            }
        """)
        self.btn_select_policy.clicked.connect(self.select_policy_file)
        layout.addWidget(self.btn_select_policy)

        self.btn_select_track = QPushButton("Select Instructor Track File")
        self.btn_select_track.setStyleSheet("""
            QPushButton {
                background-color: #003366;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #00274D;
            }
        """)
        self.btn_select_track.clicked.connect(self.select_track_file)
        layout.addWidget(self.btn_select_track)

        self.btn_select_special = QPushButton("Select Special Courses File")
        self.btn_select_special.setStyleSheet("""
            QPushButton {
                background-color: #003366;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #00274D;
            }
        """)
        self.btn_select_special.clicked.connect(self.select_special_file)
        layout.addWidget(self.btn_select_special)

        # This button triggers the actual processing (once all files are chosen)
        self.browse_button = QPushButton("Run Workload Calculation")
        self.browse_button.setStyleSheet("""
            QPushButton {
                background-color: #007E00;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #005A00;
            }
        """)
        self.browse_button.clicked.connect(self.process_excel)
        layout.addWidget(self.browse_button)

        # ---------------------------
        # Settings Button
        # ---------------------------
        self.settings_button = QPushButton("Settings")
        self.settings_button.setStyleSheet("""
            QPushButton {
                background-color: #FFC72C;
                color: black;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #E6B800;
            }
        """)
        self.settings_button.clicked.connect(self.open_settings)
        layout.addWidget(self.settings_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(35)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid white;
                border-radius: 5px;
                text-align: center;
                background-color: #1A1A1A;
                color: white;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #003366;
                width: 10px;
            }
        """)
        layout.addWidget(self.progress_bar)

        self.exit_button = QPushButton("Exit")
        self.exit_button.setStyleSheet("""
            QPushButton {
                background-color: #990000;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #7A0000;
            }
        """)
        self.exit_button.clicked.connect(self.close)
        layout.addWidget(self.exit_button)

        self.setLayout(layout)
        set_taskbar_icon(self, icon_path)

        # Store file paths
        self.raw_file_path = None
        self.policy_file_path = None
        self.track_file_path = None
        self.special_file_path = None

    # ------------
    # ADDED methods to pick each file
    # ------------
    def select_raw_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Raw Data Excel File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.raw_file_path = file_path
            QMessageBox.information(self, "Raw File Selected",
                                    f"Raw data file:\n{file_path}")

    def select_policy_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Workload Policy File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.policy_file_path = file_path
            QMessageBox.information(self, "Policy File Selected",
                                    f"Policy file:\n{file_path}")

    def select_track_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Instructor Track File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.track_file_path = file_path
            QMessageBox.information(self, "Track File Selected",
                                    f"Track file:\n{file_path}")

    def select_special_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Special Courses File", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            self.special_file_path = file_path
            QMessageBox.information(self, "Special Courses File Selected",
                                    f"Special courses file:\n{file_path}")

    def process_excel(self):
        if not self.raw_file_path:
            QMessageBox.warning(self, "Error", "Please select the Raw Data file.")
            return
        
        # We can optionally allow policy, track, special to be None if user chooses not to pick them
        # Or we can force them to pick all. For now, let's just proceed with whatever they provided.

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # Initialize our QThread-based processor
        self.thread = ExcelProcessor(
            raw_file_path=self.raw_file_path,
            policy_file_path=self.policy_file_path,
            track_file_path=self.track_file_path,
            special_file_path=self.special_file_path
        )

        # Connect signals
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.completed.connect(self.show_success)
        self.thread.error.connect(self.show_error)

        # Start the thread
        self.thread.start()

    def open_settings(self):
        settings_dialog = SettingsDialog(self, self.settings_values)
        if settings_dialog.exec():
            self.settings_values = settings_dialog.get_values()

    def show_success(self, output_file):
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Success",
                                f"Workload calculations complete.\nOutput file created at:\n{output_file}")

    def show_error(self, error_message):
        self.progress_bar.setValue(0)
        QMessageBox.critical(self, "Error", f"Failed to process the file:\n{error_message}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    icon_path = get_absolute_path("favicon.ico")
    app.setWindowIcon(QIcon(icon_path))

    window = ExcelParserApp()
    window.show()
    sys.exit(app.exec())