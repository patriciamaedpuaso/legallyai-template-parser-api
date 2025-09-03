from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from io import BytesIO
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
import traceback

app = Flask(__name__)
CORS(app)

# ===== DOCX Helper Functions =====
def apply_docx_attributes(run, attrs):
    """Apply Quill attributes to a docx run."""
    if not attrs:
        return
    if attrs.get("bold"):
        run.bold = True
    if attrs.get("italic"):
        run.italic = True
    if attrs.get("underline"):
        run.underline = True
    if attrs.get("color"):
        try:
            hex_color = attrs["color"].lstrip("#")
            run.font.color.rgb = RGBColor.from_string(hex_color)
        except Exception:
            pass


def set_paragraph_alignment_docx(para, attrs):
    """Set paragraph alignment for DOCX."""
    if not attrs:
        return
    align = attrs.get("align")
    if align == "center":
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

# ===== DOCX Conversion =====
@app.route("/convert/delta-to-docx", methods=["POST"])
def delta_to_docx():
    try:
        data = request.get_json()
        delta = data.get("delta")

        if not delta:
            return jsonify({"error": "No delta provided"}), 400

        doc = Document()
        current_para = doc.add_paragraph()

        for op in delta:
            insert_data = op.get("insert")
            attrs = op.get("attributes", {})

            # Handle tables
            if isinstance(insert_data, dict) and "table" in insert_data:
                table_data = insert_data["table"]["rows"]
                if not table_data:
                    continue

                first_row = table_data[0]
                col_keys = sorted(first_row.keys(), key=lambda x: int(x.replace("col", "")))

                table = doc.add_table(rows=len(table_data), cols=len(col_keys))

                for r_idx, row in enumerate(table_data):
                    for c_idx, col_key in enumerate(col_keys):
                        cell_ops = row.get(col_key, [])
                        cell_para = table.cell(r_idx, c_idx).paragraphs[0]
                        for cell_op in cell_ops:
                            text = cell_op.get("insert", "")
                            run = cell_para.add_run(text)
                            apply_docx_attributes(run, cell_op.get("attributes", {}))

                current_para = doc.add_paragraph()
                continue

            # Handle plain text
            if isinstance(insert_data, str):
                lines = insert_data.split("\n")
                for i, line in enumerate(lines):
                    run = current_para.add_run(line)
                    apply_docx_attributes(run, attrs)
                    set_paragraph_alignment_docx(current_para, attrs)
                    if i < len(lines) - 1:
                        current_para = doc.add_paragraph()

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="document.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("Error:", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

# ===== PDF Helper Functions =====
def get_pdf_alignment(align):
    """Map Quill alignment to ReportLab alignment."""
    if align == "center":
        return TA_CENTER
    elif align == "right":
        return TA_RIGHT
    elif align == "justify":
        return TA_JUSTIFY
    return TA_LEFT

def parse_color(color_code):
    """Convert hex color to ReportLab RGB tuple."""
    try:
        color_code = color_code.lstrip("#")
        return colors.Color(int(color_code[0:2], 16)/255.0,
                            int(color_code[2:4], 16)/255.0,
                            int(color_code[4:6], 16)/255.0)
    except:
        return colors.black

# ===== PDF Conversion =====
@app.route("/convert/delta-to-pdf", methods=["POST"])
def delta_to_pdf():
    try:
        data = request.get_json()
        delta = data.get("delta")

        if not delta:
            return jsonify({"error": "No delta provided"}), 400

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []
        styles = getSampleStyleSheet()

        for op in delta:
            insert_data = op.get("insert")
            attrs = op.get("attributes", {})

            # Handle tables
            if isinstance(insert_data, dict) and "table" in insert_data:
                table_data = insert_data["table"]["rows"]
                if not table_data:
                    continue

                first_row = table_data[0]
                col_keys = sorted(first_row.keys(), key=lambda x: int(x.replace("col", "")))

                data_rows = []
                for r_idx, row in enumerate(table_data):
                    row_cells = []
                    for c_idx, col_key in enumerate(col_keys):
                        cell_ops = row.get(col_key, [])
                        cell_text = ""
                        for cell_op in cell_ops:
                            text = cell_op.get("insert", "")
                            if "attributes" in cell_op:
                                if cell_op["attributes"].get("bold"):
                                    text = f"<b>{text}</b>"
                                if cell_op["attributes"].get("italic"):
                                    text = f"<i>{text}</i>"
                                if cell_op["attributes"].get("underline"):
                                    text = f"<u>{text}</u>"
                                if cell_op["attributes"].get("color"):
                                    color = cell_op["attributes"]["color"]
                                    text = f'<font color="{color}">{text}</font>'
                            cell_text += text
                        row_cells.append(Paragraph(cell_text, styles["Normal"]))
                    data_rows.append(row_cells)

                table = Table(data_rows, style=TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ]))
                story.append(table)
                story.append(Spacer(1, 12))
                continue

            # Handle plain text
            if isinstance(insert_data, str):
                text = insert_data
                if attrs.get("bold"):
                    text = f"<b>{text}</b>"
                if attrs.get("italic"):
                    text = f"<i>{text}</i>"
                if attrs.get("underline"):
                    text = f"<u>{text}</u>"
                if attrs.get("color"):
                    text = f'<font color="{attrs["color"]}">{text}</font>'

                alignment = get_pdf_alignment(attrs.get("align"))
                paragraph_style = ParagraphStyle(
                    name="Custom",
                    parent=styles["Normal"],
                    alignment=alignment
                )
                story.append(Paragraph(text, paragraph_style))
                story.append(Spacer(1, 6))

        doc.build(story)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name="document.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        print("Error:", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route("/routes")
def list_routes():
    return jsonify([str(rule) for rule in app.url_map.iter_rules()])

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
