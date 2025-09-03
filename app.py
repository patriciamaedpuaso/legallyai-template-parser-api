from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import pdfkit
from docx import Document
from html2docx import html2docx

app = Flask(__name__)
CORS(app)


# --- DOCX Export (with formatting + tables) ---
@app.route('/convert/html-to-docx', methods=['POST'])
def html_to_docx():
    try:
        data = request.get_json()
        html = data.get("html")

        if not html:
            return jsonify({"error": "No HTML content provided"}), 400

        # Create a new DOCX file
        output = io.BytesIO()
        doc = Document()
        html2docx(html, doc)
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="document.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# --- PDF Export (with formatting + tables) ---
@app.route('/convert/html-to-pdf', methods=['POST'])
def html_to_pdf():
    try:
        data = request.get_json()
        html = data.get("html")

        if not html:
            return jsonify({"error": "No HTML content provided"}), 400

        # Convert HTML to PDF using pdfkit (wkhtmltopdf backend)
        pdf_bytes = pdfkit.from_string(html, False)

        return send_file(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            download_name="document.pdf",
            mimetype="application/pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/')
def home():
    return jsonify({
        "status": "API running",
        "endpoints": ["/convert/html-to-docx", "/convert/html-to-pdf"]
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
