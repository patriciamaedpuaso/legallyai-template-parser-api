from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import pdfkit
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from html2docx import html2docx
import tempfile
import os

app = Flask(__name__)
CORS(app)


# ---------- Helpers for Delta → DOCX ----------
def apply_attributes(run, attrs):
    """Apply text styles like bold, italic, underline, color."""
    if not attrs:
        return
    if attrs.get("bold"):
        run.bold = True
    if attrs.get("italic"):
        run.italic = True
    if attrs.get("underline"):
        run.underline = True
    if attrs.get("color"):
        hex_color = attrs["color"].lstrip("#")
        if len(hex_color) == 6:
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            run.font.color.rgb = RGBColor(r, g, b)


def add_paragraph(doc, text, attrs, align=None):
    """Insert a paragraph with attributes + alignment."""
    para = doc.add_paragraph()
    run = para.add_run(text)
    apply_attributes(run, attrs)

    if align:
        if align == "center":
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif align == "right":
            para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        elif align == "justify":
            para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        else:
            para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    return para


def add_table(doc, table_data):
    """Insert a table from Quill Delta custom embed { insert: { table: ... } }."""
    rows = table_data.get("rows", [])
    if not rows:
        return

    # Handle uneven rows by finding max columns
    num_cols = max(len(row.keys()) for row in rows)
    table = doc.add_table(rows=len(rows), cols=num_cols)

    for r_idx, row in enumerate(rows):
        for c_idx, (col_key, cell_ops) in enumerate(row.items()):
            cell = table.cell(r_idx, c_idx)
            para = cell.paragraphs[0]
            for op in cell_ops:
                if isinstance(op.get("insert"), str):
                    run = para.add_run(op["insert"])
                    apply_attributes(run, op.get("attributes"))


def delta_to_docx(delta, output_path):
    """Convert a Quill Delta JSON into a DOCX file."""
    doc = Document()

    for op in delta:
        insert_val = op.get("insert")
        attrs = op.get("attributes", {})

        # Plain text
        if isinstance(insert_val, str):
            if insert_val == "\n":
                doc.add_paragraph("")  # line break
            else:
                add_paragraph(doc, insert_val, attrs, attrs.get("align"))

        # Table
        elif isinstance(insert_val, dict) and "table" in insert_val:
            add_table(doc, insert_val["table"])

    doc.save(output_path)


# ---------- Endpoints ----------

# HTML → DOCX
@app.route('/convert/html-to-docx', methods=['POST'])
def html_to_docx():
    try:
        data = request.get_json()
        html = data.get("html")

        if not html:
            return jsonify({"error": "No HTML content provided"}), 400

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


# HTML → PDF
@app.route('/convert/html-to-pdf', methods=['POST'])
def html_to_pdf():
    try:
        data = request.get_json()
        html = data.get("html")

        if not html:
            return jsonify({"error": "No HTML content provided"}), 400

        pdf_bytes = pdfkit.from_string(html, False)

        return send_file(
            io.BytesIO(pdf_bytes),
            as_attachment=True,
            download_name="document.pdf",
            mimetype="application/pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Delta → DOCX
@app.route('/convert/delta-to-docx', methods=['POST'])
def convert_delta_to_docx():
    try:
        delta = request.json.get("delta")
        if not delta:
            return jsonify({"error": "No delta provided"}), 400

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            delta_to_docx(delta, tmp.name)
            tmp_path = tmp.name

        response = send_file(
            tmp_path,
            as_attachment=True,
            download_name="document.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # Clean up temp file after sending
        @response.call_on_close
        def cleanup():
            try:
                os.remove(tmp_path)
            except Exception:
                pass

        return response
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Home
@app.route('/')
def home():
    return jsonify({
        "status": "API running",
        "endpoints": [
            "/convert/html-to-docx",
            "/convert/html-to-pdf",
            "/convert/delta-to-docx"
        ]
    })

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))  # Render provides PORT env var
    app.run(host="0.0.0.0", port=port)

