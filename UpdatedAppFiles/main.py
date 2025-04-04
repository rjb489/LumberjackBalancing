import sys
import os
import ctypes
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar, QDialog, QFormLayout, QLineEdit
)
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import Qt
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
        self.raw_file = None
        self.policy_file = None
        self.instructor_file = None
        self.special_file = None
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
            self.logo_label.setPixmap(pixmap.scaled(200, 200, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(self.logo_label, alignment=Qt.AlignmentFlag.AlignCenter)
        else:
            print(f"Error: Failed to load {logo_path}")
        
        self.label = QLabel("Select the necessary Excel files for processing:", self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)
        
        self.raw_button = QPushButton("Upload Raw Data File", self)
        self.raw_button.setStyleSheet("background-color: #003366; color: white; padding: 10px; border-radius: 5px;")
        self.raw_button.clicked.connect(self.select_raw_file)
        layout.addWidget(self.raw_button)
        
        self.policy_button = QPushButton("Upload Workload Policy File", self)
        self.policy_button.setStyleSheet("background-color: #003366; color: white; padding: 10px; border-radius: 5px;")
        self.policy_button.clicked.connect(self.select_policy_file)
        layout.addWidget(self.policy_button)
        
        self.instructor_button = QPushButton("Upload Instructor Track File", self)
        self.instructor_button.setStyleSheet("background-color: #003366; color: white; padding: 10px; border-radius: 5px;")
        self.instructor_button.clicked.connect(self.select_instructor_file)
        layout.addWidget(self.instructor_button)
        
        self.special_button = QPushButton("Upload Special Courses File", self)
        self.special_button.setStyleSheet("background-color: #003366; color: white; padding: 10px; border-radius: 5px;")
        self.special_button.clicked.connect(self.select_special_file)
        layout.addWidget(self.special_button)
        
        self.settings_button = QPushButton("Settings", self)
        self.settings_button.setStyleSheet("background-color: #FFC72C; color: black; padding: 10px; border-radius: 5px;")
        self.settings_button.clicked.connect(self.open_settings)
        layout.addWidget(self.settings_button)
        
        self.process_button = QPushButton("Process Files", self)
        self.process_button.setStyleSheet("background-color: #006633; color: white; padding: 10px; border-radius: 5px;")
        self.process_button.clicked.connect(self.process_excel)
        layout.addWidget(self.process_button)
        
        self.progress_bar = QProgressBar(self)
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
        
        self.exit_button = QPushButton("Exit", self)
        self.exit_button.setStyleSheet("background-color: #990000; color: white; padding: 10px; border-radius: 5px;")
        self.exit_button.clicked.connect(self.close)
        layout.addWidget(self.exit_button)
        
        self.setLayout(layout)
        set_taskbar_icon(self, icon_path)
    
    def select_raw_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Raw Data Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.raw_file = file_path
    
    def select_policy_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Workload Policy Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.policy_file = file_path
    
    def select_instructor_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Instructor Track Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.instructor_file = file_path
    
    def select_special_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Special Courses Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if file_path:
            self.special_file = file_path
    
    def process_excel(self):
        if not (self.raw_file and self.policy_file and self.instructor_file and self.special_file):
            QMessageBox.warning(self, "Missing File", "Please ensure all required files are selected.")
            return
        
        file_paths = {
            "raw": self.raw_file,
            "policy": self.policy_file,
            "instructorTrack": self.instructor_file,
            "specialCourses": self.special_file
        }
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.thread = ExcelProcessor(file_paths, self.settings_values)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.completed.connect(self.show_success)
        self.thread.error.connect(self.show_error)
        self.thread.start()
    
    def open_settings(self):
        settings_dialog = SettingsDialog(self, self.settings_values)
        if settings_dialog.exec():
            self.settings_values = settings_dialog.get_values()
    
    def show_success(self, output_file):
        self.progress_bar.setValue(100)
        QMessageBox.information(self, "Success", f"Calculations applied. Output file created at:\n{output_file}")
    
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
