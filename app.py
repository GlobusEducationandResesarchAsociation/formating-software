from flask import Flask, render_template, request, send_file
from docx import Document
from fpdf import FPDF
import os
from PyPDF2 import PdfReader

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def process_docx(file_path, output_path):
    doc = Document(file_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page(format='A4')
    pdf.set_font("Times", size=12)
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            pdf.multi_cell(0, 10, text)
            pdf.ln(5)
    pdf.output(output_path)

def process_pdf(file_path, output_path):
    reader = PdfReader(file_path)
    pdf = FPDF()
    pdf.add_page(format='A4')
    pdf.set_font("Times", size=12)
    for page in reader.pages:
        text = page.extract_text()
        if text:
            for line in text.split("\n"):
                pdf.multi_cell(0, 10, line)
                pdf.ln(5)
    pdf.output(output_path)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return "No file uploaded", 400

        filename = file.filename
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        output_path = os.path.join(UPLOAD_FOLDER, "formatted.pdf")

        if filename.endswith(".docx"):
            process_docx(file_path, output_path)
        elif filename.endswith(".pdf"):
            process_pdf(file_path, output_path)
        else:
            return "Please upload a .docx or .pdf file only", 400

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
