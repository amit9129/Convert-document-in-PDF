from flask import Flask, render_template, request, jsonify
from docx import Document
from pptx import Presentation
from PIL import Image
import os
import win32com.client
import pythoncom  # Required for COM initialization

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
            return word_to_pdf(file_path, file.filename)
        elif file_type == 'excel':
            return excel_to_pdf(file_path, file.filename)
        elif file_type == 'ppt':
            return ppt_to_pdf(file_path, file.filename)
        elif file_type == 'jpg':
            return jpg_to_pdf(file_path, file.filename)
        else:
            os.remove(file_path)  # Clean up
            return jsonify({"error": f"Conversion for {file_type.upper()} is not implemented."}), 400
    except Exception as e:
        os.remove(file_path)  # Clean up in case of error
        return jsonify({"error": str(e)}), 500

def word_to_pdf(file_path, filename):
    """
    Converts a Word document to a PDF file.
    """
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.docx', '.pdf')))

        # Open Word and convert to PDF
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the file format code for PDF
        doc.Close()
        word.Quit()

        os.remove(file_path)  # Clean up original file
        return jsonify({"message": "Word to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        os.remove(file_path)
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

def excel_to_pdf(file_path, filename):
    """
    Converts an Excel spreadsheet to a PDF file.
    """
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.xlsx', '.pdf')))

        # Open Excel and convert to PDF
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        workbook.ExportAsFixedFormat(0, pdf_path)  # 0 is the type for PDF
        workbook.Close(False)
        excel.Quit()

        os.remove(file_path)  # Clean up original file
        return jsonify({"message": "Excel to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        os.remove(file_path)
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

def ppt_to_pdf(file_path, filename):
    """
    Converts a PowerPoint presentation to a PDF file.
    """
    pythoncom.CoInitialize()  # Initialize COM library
    try:
        # Generate output file path
        pdf_path = os.path.abspath(os.path.join(PDF_FOLDER, filename.replace('.pptx', '.pdf')))

        # Open PowerPoint and convert to PDF
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)
        presentation.SaveAs(pdf_path, 32)  # 32 is the file format code for PDF
        presentation.Close()
        powerpoint.Quit()

        os.remove(file_path)  # Clean up original file
        return jsonify({"message": "PPT to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        os.remove(file_path)
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()  # Uninitialize COM library

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

        os.remove(file_path)  # Clean up original file
        return jsonify({"message": "JPG to PDF conversion successful!", "path": pdf_path}), 200
    except Exception as e:
        os.remove(file_path)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    """
    Runs the Flask application.
    """
    app.run(debug=True)
