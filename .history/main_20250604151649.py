import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QLineEdit, QMessageBox
)
from PyQt5.QtCore import Qt

class PDFtoDOCXConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('PDF to DOCX Converter')
        self.setGeometry(100, 100, 500, 200)
        layout = QVBoxLayout()

        self.pdf_label = QLabel('Select PDF file:')
        layout.addWidget(self.pdf_label)
        self.pdf_path = QLineEdit()
        self.pdf_path.setReadOnly(True)
        layout.addWidget(self.pdf_path)
        self.pdf_btn = QPushButton('Browse PDF')
        self.pdf_btn.clicked.connect(self.browse_pdf)
        layout.addWidget(self.pdf_btn)

        self.out_label = QLabel('Select Output Folder:')
        layout.addWidget(self.out_label)
        self.out_path = QLineEdit()
        self.out_path.setReadOnly(True)
        layout.addWidget(self.out_path)
        self.out_btn = QPushButton('Browse Output Folder')
        self.out_btn.clicked.connect(self.browse_output)
        layout.addWidget(self.out_btn)

        self.convert_btn = QPushButton('Convert to DOCX')
        self.convert_btn.clicked.connect(self.convert_pdf_to_docx)
        layout.addWidget(self.convert_btn)

        self.setLayout(layout)

    def browse_pdf(self):
        file_name, _ = QFileDialog.getOpenFileName(self, 'Select PDF File', '', 'PDF Files (*.pdf)')
        if file_name:
            self.pdf_path.setText(file_name)

    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
        if folder:
            self.out_path.setText(folder)

    def convert_pdf_to_docx(self):
        pdf_file = self.pdf_path.text()
        output_folder = self.out_path.text()
        if not pdf_file or not output_folder:
            QMessageBox.warning(self, 'Missing Information', 'Please select both a PDF file and an output folder.')
            return
        # Stub: Conversion logic will be implemented here
        QMessageBox.information(self, 'Conversion', 'Conversion started! (Stub)')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFtoDOCXConverter()
    window.show()
    sys.exit(app.exec_()) 