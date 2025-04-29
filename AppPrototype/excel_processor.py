# excel_processor.py
# This is where you should put your caluclations @Cristan

import pandas as pd
import openpyxl
import time
import random
from PyQt6.QtCore import QThread, pyqtSignal

class ExcelProcessor(QThread):
    progress = pyqtSignal(int)
    completed = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, file_path, settings):
        super().__init__()
        self.file_path = file_path
        self.settings = settings
    
    def run(self):
        try:
            df = pd.read_excel(self.file_path)

            for key, value in self.settings.items():
                df.loc[key, df.columns[0]] = value

            df.loc["Verification", df.columns[0]] = "Processing successful"

            progress = 0
            while progress < 100:
                increment = random.randint(1, 5)  # random increment between 1 and 5
                progress += increment
                progress = min(progress, 100)  # don't exceed 100%
                self.progress.emit(progress)
                time.sleep(random.uniform(0.1, 0.3))  # random delay between 0.1 and 0.3 seconds


            # Save the processed file
            output_file = self.file_path.replace(".xlsx", "_processed.xlsx")
            df.to_excel(output_file, index=True)  
            
            self.completed.emit(output_file)
        except Exception as e:
            self.error.emit(str(e))
