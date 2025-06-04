import os
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
from werkzeug.utils import secure_filename
import PyPDF2
from docx import Document
from docx.shared import Inches
import arabic_reshaper
from bidi.algorithm import get_display

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_word(pdf_path, docx_path):
    # Create a new Word document
    doc = Document()
    
    # Open the PDF file
    with open(pdf_path, 'rb') as file:
        # Create PDF reader object
        pdf_reader = PyPDF2.PdfReader(file)
        
        # Get number of pages
        num_pages = len(pdf_reader.pages)
        
        # Process each page
        for page_num in range(num_pages):
            # Get the page
            page = pdf_reader.pages[page_num]
            
            # Extract text
            text = page.extract_text()
            
            # Handle Arabic/Urdu text
            if any(ord(c) > 127 for c in text):
                text = get_display(arabic_reshaper.reshape(text))
            
            # Add text to document
            doc.add_paragraph(text)
            
            # Add page break if not the last page
            if page_num < num_pages - 1:
                doc.add_page_break()
    
    # Save the document
    doc.save(docx_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(upload_path)
            
            # Convert PDF to DOCX
            output_filename = filename.rsplit('.', 1)[0] + '.docx'
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            
            try:
                convert_pdf_to_word(upload_path, output_path)
                return redirect(url_for('download_file', filename=output_filename))
            except Exception as e:
                flash(f'Error during conversion: {str(e)}')
                return redirect(request.url)
        else:
            flash('Allowed file type is PDF')
            return redirect(request.url)
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True) 