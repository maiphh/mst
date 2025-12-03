import sys
import os
import shutil
import openpyxl
import platform
import pytesseract
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QProgressBar, QTextEdit, QMessageBox, QCheckBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from check_tax_official import check_cccd_official

class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal()
    
    def __init__(self, input_path, open_browser=False):
        super().__init__()
        self.input_path = input_path
        self.open_browser = open_browser
        self.is_running = True

    def log_wrapper(self, message):
        self.log_signal.emit(message)

    def run(self):
        try:
            # Create output filename
            dir_name = os.path.dirname(self.input_path)
            base_name = os.path.basename(self.input_path)
            name, ext = os.path.splitext(base_name)
            output_path = os.path.join(dir_name, f"{name}_processed{ext}")
            
            # Copy file first
            shutil.copy2(self.input_path, output_path)
            self.log_signal.emit(f"Created working copy: {output_path}")
            
            # Load workbook
            wb = openpyxl.load_workbook(output_path)
            sheet = wb.active
            
            # Map columns
            headers = {}
            for cell in sheet[1]:
                if cell.value:
                    headers[str(cell.value).strip()] = cell.column
            
            required_cols = ["CMND/CCCD", "MST 1", "Tên người nộp thuế", "Cơ quan thuế", "Ghi chú MST 1"]
            missing = [c for c in required_cols if c not in headers]
            
            if missing:
                self.log_signal.emit(f"Error: Missing columns: {', '.join(missing)}")
                self.finished_signal.emit()
                return

            col_map = {col: headers[col] for col in required_cols}
            
            # Find rows to process
            rows_to_process = []
            max_row = sheet.max_row
            
            for row_idx in range(2, max_row + 1):
                mst_val = sheet.cell(row=row_idx, column=col_map["MST 1"]).value
                cccd_val = sheet.cell(row=row_idx, column=col_map["CMND/CCCD"]).value
                
                if cccd_val and not mst_val:
                    rows_to_process.append(row_idx)
            
            total = len(rows_to_process)
            self.log_signal.emit(f"Found {total} rows to process")
            
            for i, row_idx in enumerate(rows_to_process):
                if not self.is_running:
                    break
                    
                cccd_val = sheet.cell(row=row_idx, column=col_map["CMND/CCCD"]).value
                cccd_str = str(cccd_val).strip()
                
                progress_pct = int(((i) / total) * 100)
                self.progress_signal.emit(progress_pct, f"Processing {i+1}/{total}: {cccd_str}")
                
                try:
                    result = check_cccd_official(cccd_str, open_browser=self.open_browser, log_callback=self.log_wrapper)
                    
                    # Update cells
                    if result.get("tax_id"):
                        sheet.cell(row=row_idx, column=col_map["MST 1"]).value = result["tax_id"]
                    if result.get("name"):
                        sheet.cell(row=row_idx, column=col_map["Tên người nộp thuế"]).value = result["name"]
                    if result.get("place"):
                        sheet.cell(row=row_idx, column=col_map["Cơ quan thuế"]).value = result["place"]
                    if result.get("status"):
                        sheet.cell(row=row_idx, column=col_map["Ghi chú MST 1"]).value = result["status"]
                    
                    # Save immediately
                    wb.save(output_path)
                    self.log_signal.emit(f"Processed {cccd_str}: {result.get('status')} - {result.get('tax_id')}")
                    
                except Exception as e:
                    self.log_signal.emit(f"Error processing {cccd_str}: {str(e)}")
            
            self.progress_signal.emit(100, "Completed")
            self.log_signal.emit("Processing complete. File saved.")
            
        except Exception as e:
            self.log_signal.emit(f"Critical Error: {str(e)}")
        finally:
            self.finished_signal.emit()

    def stop(self):
        self.is_running = False

class TaxCheckerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tax Checker Tool")
        self.setGeometry(100, 100, 800, 600)
        
        self.file_path = ""
        self.worker = None
        
        self.init_ui()
        self.check_tesseract()
        
    def check_tesseract(self):
        # Check if tesseract is in PATH
        if shutil.which("tesseract"):
            return

        # Common Windows paths
        if platform.system() == "Windows":
            common_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe")
            ]
            
            for path in common_paths:
                if os.path.exists(path):
                    pytesseract.tesseract_cmd = path
                    self.log(f"Found Tesseract at: {path}")
                    return

        # If not found, ask user
        reply = QMessageBox.question(self, "Tesseract Not Found", 
                                   "Tesseract OCR was not found in your PATH or common locations.\n"
                                   "Would you like to locate tesseract.exe manually?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            filename, _ = QFileDialog.getOpenFileName(self, "Locate Tesseract", "", "Executables (*.exe);;All Files (*)")
            if filename:
                pytesseract.tesseract_cmd = filename
                self.log(f"User selected Tesseract at: {filename}")
                return

        QMessageBox.warning(self, "Missing Dependency", 
                            "Tesseract OCR is not found on your system.\n\n"
                            "Please install it to use this application:\n"
                            "Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
                            "Mac: brew install tesseract")
        self.log("Warning: Tesseract not found. Please install it or add to PATH.")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # File Selection
        file_layout = QHBoxLayout()
        self.path_label = QLabel("No file selected")
        self.path_label.setStyleSheet("border: 1px solid #ccc; padding: 5px; background: white;")
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_file)
        
        file_layout.addWidget(self.path_label, stretch=1)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)
        
        # Options
        self.browser_checkbox = QCheckBox("Show Browser")
        self.browser_checkbox.setChecked(False)
        layout.addWidget(self.browser_checkbox)
        
        # Action Buttons
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start Processing")
        self.start_btn.clicked.connect(self.start_processing)
        self.start_btn.setEnabled(False)
        
        self.status_label = QLabel("Ready")
        
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.status_label)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # Log Area
        layout.addWidget(QLabel("Logs:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)
        
    def browse_file(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if filename:
            self.file_path = filename
            self.path_label.setText(filename)
            self.start_btn.setEnabled(True)
            self.log(f"Selected file: {filename}")
            
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_area.append(f"[{timestamp}] {message}")
        # Scroll to bottom
        cursor = self.log_area.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log_area.setTextCursor(cursor)
        
    def start_processing(self):
        if not self.file_path or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Error", "Invalid file path")
            return
            
        self.start_btn.setEnabled(False)
        self.log("Starting processing...")
        self.progress_bar.setValue(0)
        
        open_browser = self.browser_checkbox.isChecked()
        self.worker = WorkerThread(self.file_path, open_browser=open_browser)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.processing_finished)
        self.worker.start()
        
    def update_progress(self, value, status):
        self.progress_bar.setValue(value)
        self.status_label.setText(status)
        
    def processing_finished(self):
        self.start_btn.setEnabled(True)
        self.status_label.setText("Finished")
        self.worker = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TaxCheckerApp()
    window.show()
    sys.exit(app.exec())
