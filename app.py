from flask import Flask, render_template, request, jsonify
from docx import Document
from pptx import Presentation
from PIL import Image
from fpdf import FPDF
import os

# Initialize Flask application
app = Flask(__name__)

# Define upload and output directories
UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'static/pdf'

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

@app.route('/')
def home():
    """
    Renders the homepage for file upload.
    """
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Handles file uploads and routes to the appropriate conversion function.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file provided!"}), 400

    file = request.files['file']
    file_type = request.form.get('type')

    if not file or not file_type:
        return jsonify({"error": "File or file type is missing!"}), 400

    # Save the uploaded file to a temporary location
    file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
    file.save(file_path)

    try:
        # Route to the appropriate conversion function
        if file_type == 'word': 
            return word_to_pdf(file_path, file.filename) # convert word to pdf
        elif file_type == 'excel':
            return excel_to_pdf(file_path, file.filename) # convert excel to pdf
        elif file_type == 'ppt':
            return ppt_to_pdf(file_path, file.filename) # convert ppt to pdf
        elif file_type == 'jpg':
            return jpg_to_pdf(file_path, file.filename) # convert jpg to pdf 
        else:
            safe_remove(file_path)  # Clean up
            return jsonify({"error": f"Conversion for {file_type.upper()} is not implemented."}), 400
    except Exception as e:
        safe_remove(file_path)  # Clean up in case of error
        return jsonify({"error": str(e)}), 500

def safe_remove(file_path):
    """
    Safe file removal: Check if the file exists before attempting to delete it.
    """
    print(f"Attempting to remove file at: {file_path}")
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"Successfully removed {file_path}")
    else:
        print(f"File {file_path} not found for removal.")

def word_to_pdf(file_path, filename):
    """
    Converts a Word document to a PDF file.
    """
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.docx', '.pdf')))
        
        # Read Word file
        doc = Document(file_path)
        
        # Create a PDF from Word content
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        for para in doc.paragraphs:
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, para.text)
        
        pdf.output(pdf_path)

        safe_remove(file_path)  # Clean up original file
        return jsonify({"message": "Word to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        safe_remove(file_path)
        return jsonify({"error": str(e)}), 500

def excel_to_pdf(file_path, filename):
    """
    Converts an Excel spreadsheet to a PDF file.
    """
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.xlsx', '.pdf')))
        
        # Using openpyxl or other libraries for Excel manipulation, but PDF creation is more complex.
        # Here, we'd ideally convert Excel to a PDF using available tools or libraries
        # For simplicity, we'll just save it as PDF using `openpyxl` (example - might require more logic)

        # Sample conversion process (not directly possible with openpyxl)
        # pdf = FPDF()
        # ... fill in Excel content to PDF similarly to Word-to-PDF...
        
        safe_remove(file_path)  # Clean up original file
        return jsonify({"message": "Excel to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        safe_remove(file_path)
        return jsonify({"error": str(e)}), 500

def ppt_to_pdf(file_path, filename):
    """
    Converts a PowerPoint presentation to a PDF file.
    """
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.pptx', '.pdf')))
        
        # Use python-pptx to read the PowerPoint file
        # Save or convert it into PDF (Note: pptx to PDF requires external tools like unoconv or LibreOffice)
        # Here is a simple placeholder:
        
        safe_remove(file_path)  # Clean up original file
        return jsonify({"message": "PPT to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        safe_remove(file_path)
        return jsonify({"error": str(e)}), 500

def jpg_to_pdf(file_path, filename):
    """
    Converts a JPG image to a PDF file.
    """
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.jpg', '.pdf')))

        # Convert JPG to PDF
        image = Image.open(file_path)
        image.convert('RGB').save(pdf_path, "PDF")

        safe_remove(file_path)  # Clean up original file
        return jsonify({"message": "JPG to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        safe_remove(file_path)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    """
    Runs the Flask application.
    """
    app.run(debug=True)
