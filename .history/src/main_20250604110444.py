import sys
import os

# Add the parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QPushButton, QLabel, QFileDialog, QProgressBar,
                            QMessageBox, QListWidget)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from src.converter import PDFConverter

class ConversionWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, files, output_dir):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.converter = PDFConverter()

    def run(self):
        try:
            total_files = len(self.files)
            for i, file in enumerate(self.files):
                self.converter.convert(file, self.output_dir)
                progress = int((i + 1) / total_files * 100)
                self.progress.emit(progress)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to DOCX Converter")
        self.setMinimumSize(600, 400)
        self.setup_ui()
        self.files_to_convert = []

    def setup_ui(self):
        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # File list
        self.file_list = QListWidget()
        layout.addWidget(QLabel("Files to Convert:"))
        layout.addWidget(self.file_list)

        # Buttons
        self.select_files_btn = QPushButton("Select PDF Files")
        self.select_files_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_files_btn)

        self.select_output_btn = QPushButton("Select Output Directory")
        self.select_output_btn.clicked.connect(self.select_output_dir)
        layout.addWidget(self.select_output_btn)

        self.convert_btn = QPushButton("Convert")
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)

        # Progress bar
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)

        # Enable drag and drop
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        self.add_files(files)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF Files",
            "",
            "PDF Files (*.pdf)"
        )
        self.add_files(files)

    def add_files(self, files):
        for file in files:
            if file.lower().endswith('.pdf'):
                if file not in self.files_to_convert:
                    self.files_to_convert.append(file)
                    self.file_list.addItem(os.path.basename(file))

    def select_output_dir(self):
        self.output_dir = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory"
        )

    def start_conversion(self):
        if not self.files_to_convert:
            QMessageBox.warning(self, "Warning", "Please select files to convert.")
            return

        if not hasattr(self, 'output_dir'):
            QMessageBox.warning(self, "Warning", "Please select an output directory.")
            return

        self.convert_btn.setEnabled(False)
        self.select_files_btn.setEnabled(False)
        self.select_output_btn.setEnabled(False)
        self.progress_bar.setValue(0)

        self.worker = ConversionWorker(self.files_to_convert, self.output_dir)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        self.status_label.setText(f"Converting... {value}%")

    def conversion_finished(self):
        self.convert_btn.setEnabled(True)
        self.select_files_btn.setEnabled(True)
        self.select_output_btn.setEnabled(True)
        self.status_label.setText("Conversion completed!")
        QMessageBox.information(self, "Success", "All files have been converted successfully!")

    def conversion_error(self, error_msg):
        self.convert_btn.setEnabled(True)
        self.select_files_btn.setEnabled(True)
        self.select_output_btn.setEnabled(True)
        self.status_label.setText("Error occurred during conversion.")
        QMessageBox.critical(self, "Error", f"An error occurred: {error_msg}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 