from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
import fitz  # PyMuPDF
import requests
from io import BytesIO
import os
from urllib.parse import urlparse
from docx.enum.text import WD_ALIGN_PARAGRAPH


app = Flask(__name__)
CORS(app)  # Enable CORS for all domains

def extract_docx_to_delta(file_bytes):
    doc = Document(file_bytes)
    delta = []

    for para in doc.paragraphs:
        paragraph_alignment = para.alignment
        align_str = None
        if paragraph_alignment == WD_ALIGN_PARAGRAPH.CENTER:
            align_str = "center"
        elif paragraph_alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            align_str = "right"
        elif paragraph_alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            align_str = "justify"
        # default is left or None

        for run in para.runs:
            attrs = {}

            if run.bold:
                attrs["bold"] = True
            if run.italic:
                attrs["italic"] = True
            if run.underline:
                attrs["underline"] = True
            if run.font and run.font.strike:
                attrs["strike"] = True
            if run.font and run.font.superscript:
                attrs["script"] = "super"
            if run.font and run.font.subscript:
                attrs["script"] = "sub"
            if run.font.name:
                attrs["font"] = run.font.name

            if run.font and run.font.size:
                try:
                    pt = run.font.size.pt
                    if pt <= 10:
                        attrs["size"] = "small"
                    elif pt >= 16:
                        attrs["size"] = "large"
                except:
                    pass

            if run.font and run.font.color and run.font.color.rgb:
                attrs["color"] = f"#{run.font.color.rgb}"

            insert_obj = {
                "insert": run.text
            }
            if attrs:
                insert_obj["attributes"] = attrs
            delta.append(insert_obj)

        # Append alignment as a block-level attribute on the paragraph break
        paragraph_break = {"insert": "\n"}
        if align_str:
            paragraph_break["attributes"] = {"align": align_str}

        delta.append(paragraph_break)

    return delta

def extract_pdf_to_delta(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    delta = []

    for page in doc:
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] != 0:  # only text blocks
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"]
                    if not text.strip():
                        continue

                    attrs = {}

                    # Bold and Italic detection (based on font name)
                    font_name = span.get("font", "")
                    if "Bold" in font_name:
                        attrs["bold"] = True
                    if "Italic" in font_name or "Oblique" in font_name:
                        attrs["italic"] = True
                    if font_name:
                        attrs["font"] = font_name

                    # Font size approximation
                    size = span.get("size", 12)
                    if size <= 10:
                        attrs["size"] = "small"
                    elif size >= 16:
                        attrs["size"] = "large"

                    # Text color (optional)
                    color = span.get("color", None)
                    if color:
                        hex_color = "#{:06x}".format(color)
                        attrs["color"] = hex_color

                    insert_obj = {
                        "insert": text
                    }
                    if attrs:
                        insert_obj["attributes"] = attrs

                    delta.append(insert_obj)

                delta.append({"insert": "\n"})

    return delta


@app.route('/extract', methods=['POST'])
def extract():
    try:
        file_url = request.json.get('fileUrl')
        if not file_url:
            return jsonify({"error": "No fileUrl provided"}), 400

        # Download the file
        response = requests.get(file_url)
        if response.status_code != 200:
            return jsonify({"error": "Failed to fetch the file"}), 400

        file_bytes = BytesIO(response.content)

        # Determine file extension
        parsed_url = urlparse(file_url)
        filename = os.path.basename(parsed_url.path)
        ext = os.path.splitext(filename)[1].lower()

        if ext == ".docx":
            delta = extract_docx_to_delta(file_bytes)
        elif ext == ".pdf":
            delta = extract_pdf_to_delta(file_bytes)
        else:
            return jsonify({"error": f"Unsupported file type: {ext}"}), 400

        return jsonify({"delta": delta})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
