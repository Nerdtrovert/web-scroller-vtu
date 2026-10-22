"""
Universal File I/O module for VTU Result Automation.
Supports multi-format reading (Excel, CSV, DOCX, PDF, TXT)
and multi-format exporting (Excel .xlsx / .xls, CSV, Word .docx, PDF).
"""

import os
import re
import csv
import logging
from datetime import datetime
from logger import setup_logger

# Excel libraries
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import xlwt
    HAS_XLWT = True
except ImportError:
    HAS_XLWT = False

# Word document library
try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# PDF libraries
try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from reportlab.lib.pagesizes import landscape, letter, A4, A3
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

logger = setup_logger(__name__)

# Standard VTU USN regex pattern: e.g., 1HC24CS058, 1MS21IS045, 1RV19EC001, 1DS22CS401
USN_REGEX = re.compile(r"\b([0-9][A-Za-z]{2}\d{2}[A-Za-z0-9]{2,5})\b")


# ==============================================================================
# INPUT READERS (USN List Extraction)
# ==============================================================================

def load_usns_from_file(file_path):
    """
    Extract a list of student USNs from an input file.
    Supports .xlsx, .xlsm, .xls, .csv, .docx, .pdf, and .txt formats.
    
    Args:
        file_path: Absolute or relative path to the input file.
        
    Returns:
        List of cleaned, uppercase USN strings (preserving order, deduplicated).
    """
    if not os.path.exists(file_path):
        logger.error(f"Input file not found: {file_path}")
        return []

    ext = os.path.splitext(file_path)[1].lower()
    raw_usns = []

    try:
        if ext in [".xlsx", ".xlsm"]:
            raw_usns = _read_usns_from_excel(file_path)
        elif ext == ".csv":
            raw_usns = _read_usns_from_csv(file_path)
        elif ext == ".docx":
            if not HAS_DOCX:
                raise ImportError("python-docx is required to read .docx files.")
            raw_usns = _read_usns_from_docx(file_path)
        elif ext == ".pdf":
            if not HAS_PYPDF:
                raise ImportError("pypdf is required to read .pdf files.")
            raw_usns = _read_usns_from_pdf(file_path)
        elif ext == ".xls":
            raw_usns = _read_usns_from_legacy_xls(file_path)
        else:
            # Fallback for .txt or unrecognized plain text files
            raw_usns = _read_usns_from_txt(file_path)

    except Exception as e:
        logger.error(f"Error reading USNs from {file_path}: {e}")
        return []

    # Clean and deduplicate while preserving original order
    seen = set()
    cleaned_usns = []
    for item in raw_usns:
        if not item or not isinstance(item, str):
            continue
        val = item.strip().upper()
        if val and val not in seen:
            seen.add(val)
            cleaned_usns.append(val)

    logger.info(f"Loaded {len(cleaned_usns)} unique USNs from {file_path}")
    return cleaned_usns


def _read_usns_from_excel(file_path):
    usns = []
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        found = False
        for cell in row:
            if cell and isinstance(cell, str):
                matches = USN_REGEX.findall(cell)
                if matches:
                    usns.extend(matches)
                    found = True
                    break
        if not found and row[0] is not None:
            first_val = str(row[0]).strip()
            if first_val and not first_val.lower().startswith(("usn", "seat", "roll", "sl", "no")):
                usns.append(first_val)
    return usns


def _read_usns_from_legacy_xls(file_path):
    usns = []
    try:
        import xlrd
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(0)
        for rx in range(sheet.nrows):
            val = sheet.cell_value(rx, 0)
            if val:
                s_val = str(val).strip()
                matches = USN_REGEX.findall(s_val)
                if matches:
                    usns.extend(matches)
                elif not s_val.lower().startswith(("usn", "seat", "roll", "sl")):
                    usns.append(s_val)
    except Exception as e:
        logger.warning(f"Could not read .xls via xlrd: {e}. Trying raw text scan.")
        usns = _read_usns_from_txt(file_path)
    return usns


def _read_usns_from_csv(file_path):
    usns = []
    with open(file_path, mode="r", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            found = False
            for cell in row:
                matches = USN_REGEX.findall(cell)
                if matches:
                    usns.extend(matches)
                    found = True
                    break
            if not found and row[0].strip():
                val = row[0].strip()
                if not val.lower().startswith(("usn", "seat", "roll", "sl", "no")):
                    usns.append(val)
    return usns


def _read_usns_from_docx(file_path):
    usns = []
    doc = docx.Document(file_path)
    # Check tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = USN_REGEX.findall(cell.text)
                if matches:
                    usns.extend(matches)
    # Check paragraphs
    for para in doc.paragraphs:
        matches = USN_REGEX.findall(para.text)
        if matches:
            usns.extend(matches)
    return usns


def _read_usns_from_pdf(file_path):
    usns = []
    reader = PdfReader(file_path)
    for page in reader.pages:
        text = page.extract_text() or ""
        matches = USN_REGEX.findall(text)
        if matches:
            usns.extend(matches)
    return usns


def _read_usns_from_txt(file_path):
    usns = []
    with open(file_path, mode="r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            matches = USN_REGEX.findall(line)
            if matches:
                usns.extend(matches)
            elif line.strip() and not line.strip().lower().startswith(("usn", "seat", "roll")):
                usns.append(line.strip())
    return usns


# ==============================================================================
# DATA NORMALIZATION HELPERS
# ==============================================================================

def extract_unique_subjects(results_data):
    """
    Extract a canonical list of subject dictionaries across all student results.
    
    Returns:
        List of dicts: [{'code': 'BCS401', 'name': 'ADA'}, ...]
    """
    subject_map = {}
    for res in results_data:
        for sub in res.get("subjects", []):
            code = sub.get("code") or sub.get("name")
            name = sub.get("name") or code
            if code and code not in subject_map:
                subject_map[code] = name
    
    return [{"code": c, "name": n} for c, n in subject_map.items()]


# ==============================================================================
# OUTPUT EXPORTERS
# ==============================================================================

def export_results(results_data, output_file_path):
    """
    Export scraped student results to the requested file format.
    Automatically determines format from output_file_path extension:
      - .xlsx : Modern Excel workbook (rich styled)
      - .xls  : Legacy Excel 97-2003 workbook
      - .csv  : Comma-separated values
      - .docx : Microsoft Word document with styled results table
      - .pdf  : Printable PDF report with colored pass/fail indicators
      
    Args:
        results_data: List of student result dicts
        output_file_path: Destination file path
    """
    if not results_data:
        logger.warning("No results to export.")
        return

    ext = os.path.splitext(output_file_path)[1].lower()
    if not ext:
        ext = ".xlsx"
        output_file_path += ext

    logger.info(f"Exporting {len(results_data)} results to {ext.upper()} format at {output_file_path}")

    if ext == ".csv":
        export_to_csv(results_data, output_file_path)
    elif ext == ".docx":
        export_to_docx(results_data, output_file_path)
    elif ext == ".pdf":
        export_to_pdf(results_data, output_file_path)
    elif ext == ".xls":
        export_to_xls(results_data, output_file_path)
    else:
        # Default to .xlsx
        export_to_xlsx(results_data, output_file_path)


def export_to_csv(results_data, output_file_path):
    """Export results to CSV file."""
    unique_subjects = extract_unique_subjects(results_data)

    hdr1 = ["USN", "STUDENT NAME", "SEMESTER"]
    hdr2 = ["", "", ""]
    for s in unique_subjects:
        label = f"{s['code']} - {s['name']}" if s['code'] != s['name'] else s['name']
        hdr1.extend([label, "", "", ""])
        hdr2.extend(["IA", "SEE", "TOTAL", "RES"])

    with open(output_file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(hdr1)
        writer.writerow(hdr2)

        for res in results_data:
            row = [res.get("usn", ""), res.get("name", ""), res.get("semester", "")]
            # Map subjects by code or name
            stud_subs = {s.get("code") or s.get("name"): s for s in res.get("subjects", [])}
            for u_sub in unique_subjects:
                key = u_sub["code"]
                if key in stud_subs:
                    s_data = stud_subs[key]
                    row.extend([
                        s_data.get("ia", "-"),
                        s_data.get("see", "-"),
                        s_data.get("total", "-"),
                        s_data.get("res", "-")
                    ])
                else:
                    row.extend(["-", "-", "-", "-"])
            writer.writerow(row)

    logger.info(f"CSV results successfully saved to {output_file_path}")


def export_to_xlsx(results_data, output_file_path):
    """Export results to a styled Excel (.xlsx) workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VTU Results"

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    subheader_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    unique_subjects = extract_unique_subjects(results_data)

    # Static Header Columns
    ws.cell(row=1, column=1, value="USN").font = header_font
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=2, value="STUDENT NAME").font = header_font
    ws.cell(row=1, column=2).fill = header_fill
    ws.cell(row=1, column=3, value="SEMESTER").font = header_font
    ws.cell(row=1, column=3).fill = header_fill

    ws.cell(row=2, column=1, value="").fill = subheader_fill
    ws.cell(row=2, column=2, value="").fill = subheader_fill
    ws.cell(row=2, column=3, value="").fill = subheader_fill

    col_idx = 4
    for s in unique_subjects:
        sub_label = f"{s['code']} - {s['name']}" if s['code'] != s['name'] else s['name']
        ws.cell(row=1, column=col_idx, value=sub_label).font = header_font
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)

        ws.cell(row=2, column=col_idx, value="IA").font = bold_font
        ws.cell(row=2, column=col_idx).fill = subheader_fill
        ws.cell(row=2, column=col_idx + 1, value="SEE").font = bold_font
        ws.cell(row=2, column=col_idx + 1).fill = subheader_fill
        ws.cell(row=2, column=col_idx + 2, value="TOTAL").font = bold_font
        ws.cell(row=2, column=col_idx + 2).fill = subheader_fill
        ws.cell(row=2, column=col_idx + 3, value="RES").font = bold_font
        ws.cell(row=2, column=col_idx + 3).fill = subheader_fill
        col_idx += 4

    # Populate Student Rows
    current_row = 3
    for res in results_data:
        c_usn = ws.cell(row=current_row, column=1, value=res.get("usn", ""))
        c_usn.font = bold_font
        c_usn.border = thin_border

        c_name = ws.cell(row=current_row, column=2, value=res.get("name", ""))
        c_name.font = regular_font
        c_name.border = thin_border

        c_sem = ws.cell(row=current_row, column=3, value=res.get("semester", ""))
        c_sem.font = regular_font
        c_sem.border = thin_border

        stud_subs = {s.get("code") or s.get("name"): s for s in res.get("subjects", [])}
        col_idx = 4
        for u_sub in unique_subjects:
            key = u_sub["code"]
            if key in stud_subs:
                s_data = stud_subs[key]
                c_ia = ws.cell(row=current_row, column=col_idx, value=s_data.get("ia", "-"))
                c_ia.font = regular_font
                c_ia.border = thin_border

                c_see = ws.cell(row=current_row, column=col_idx + 1, value=s_data.get("see", "-"))
                c_see.font = regular_font
                c_see.border = thin_border

                c_tot = ws.cell(row=current_row, column=col_idx + 2, value=s_data.get("total", "-"))
                c_tot.font = bold_font
                c_tot.border = thin_border

                c_res = ws.cell(row=current_row, column=col_idx + 3, value=s_data.get("res", "-"))
                c_res.font = bold_font
                c_res.border = thin_border

                res_val = str(s_data.get("res", "")).upper()
                if res_val == "P":
                    c_res.fill = pass_fill
                elif res_val in ["F", "A"]:
                    c_res.fill = fail_fill
            else:
                for offset in range(4):
                    cell = ws.cell(row=current_row, column=col_idx + offset, value="-")
                    cell.border = thin_border
            col_idx += 4
        current_row += 1

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(output_file_path)
    logger.info(f"Excel .xlsx results successfully saved to {output_file_path}")


def export_to_xls(results_data, output_file_path):
    """Export results to legacy Excel (.xls) using xlwt."""
    if not HAS_XLWT:
        logger.warning("xlwt not installed; falling back to .xlsx format.")
        export_to_xlsx(results_data, output_file_path.replace(".xls", ".xlsx"))
        return

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("VTU Results", cell_overwrite_ok=True)
    pass_style = xlwt.easyxf("pattern: pattern solid, fore_colour light_green; font: bold on")
    fail_style = xlwt.easyxf("pattern: pattern solid, fore_colour rose; font: bold on")
    header_style = xlwt.easyxf("font: bold on, color white; pattern: pattern solid, fore_colour ocean_blue")
    subhdr_style = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour ice_blue")

    unique_subjects = extract_unique_subjects(results_data)

    sheet.write(0, 0, "USN", header_style)
    sheet.write(0, 1, "NAME", header_style)
    sheet.write(0, 2, "SEMESTER", header_style)

    col = 3
    for s in unique_subjects:
        sheet.write(0, col, s["name"], header_style)
        sheet.write(1, col, "IA", subhdr_style)
        sheet.write(1, col + 1, "SEE", subhdr_style)
        sheet.write(1, col + 2, "TOTAL", subhdr_style)
        sheet.write(1, col + 3, "RES", subhdr_style)
        col += 4

    for r_idx, res in enumerate(results_data, start=2):
        sheet.write(r_idx, 0, res.get("usn", ""))
        sheet.write(r_idx, 1, res.get("name", ""))
        sheet.write(r_idx, 2, res.get("semester", ""))

        stud_subs = {s.get("code") or s.get("name"): s for s in res.get("subjects", [])}
        col = 3
        for u_sub in unique_subjects:
            key = u_sub["code"]
            if key in stud_subs:
                s_data = stud_subs[key]
                sheet.write(r_idx, col, str(s_data.get("ia", "-")))
                sheet.write(r_idx, col + 1, str(s_data.get("see", "-")))
                sheet.write(r_idx, col + 2, str(s_data.get("total", "-")))
                res_val = str(s_data.get("res", "")).upper()
                style = pass_style if res_val == "P" else (fail_style if res_val in ["F", "A"] else xlwt.Style.default_style)
                sheet.write(r_idx, col + 3, res_val, style)
            else:
                for off in range(4):
                    sheet.write(r_idx, col + off, "-")
            col += 4

    workbook.save(output_file_path)
    logger.info(f"Legacy Excel .xls saved to {output_file_path}")


def export_to_docx(results_data, output_file_path):
    """Export results to a Microsoft Word (.docx) document."""
    if not HAS_DOCX:
        raise ImportError("python-docx is required to export to .docx")

    doc = docx.Document()
    
    # Page Margins (compact landscape for wide table)
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    title = doc.add_heading("Visvesvaraya Technological University", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph(f"Student Examination Results Report | Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    unique_subjects = extract_unique_subjects(results_data)
    total_cols = 3 + len(unique_subjects) * 4

    table = doc.add_table(rows=2, cols=total_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    # Header Row 1
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "USN"
    hdr_cells[1].text = "NAME"
    hdr_cells[2].text = "SEM"

    # Set background color helper
    def set_cell_background(cell, hex_color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    col_idx = 3
    for s in unique_subjects:
        hdr_cells[col_idx].text = f"{s['code']} - {s['name']}"
        col_idx += 4

    # Header Row 2 (Subheaders)
    subhdr_cells = table.rows[1].cells
    col_idx = 3
    for s in unique_subjects:
        subhdr_cells[col_idx].text = "IA"
        subhdr_cells[col_idx + 1].text = "SEE"
        subhdr_cells[col_idx + 2].text = "TOT"
        subhdr_cells[col_idx + 3].text = "RES"
        col_idx += 4

    # Style Header Rows
    for cell in table.rows[0].cells:
        set_cell_background(cell, "1F497D")
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(8.5)

    for cell in table.rows[1].cells:
        set_cell_background(cell, "DCE6F1")
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(31, 73, 125)
                run.font.size = Pt(8)

    # Populate Student Data Rows
    for res in results_data:
        row_cells = table.add_row().cells
        row_cells[0].text = res.get("usn", "")
        row_cells[1].text = res.get("name", "")
        row_cells[2].text = res.get("semester", "")

        stud_subs = {s.get("code") or s.get("name"): s for s in res.get("subjects", [])}
        col_idx = 3
        for u_sub in unique_subjects:
            key = u_sub["code"]
            if key in stud_subs:
                s_data = stud_subs[key]
                row_cells[col_idx].text = str(s_data.get("ia", "-"))
                row_cells[col_idx + 1].text = str(s_data.get("see", "-"))
                row_cells[col_idx + 2].text = str(s_data.get("total", "-"))
                res_val = str(s_data.get("res", "-")).upper()
                row_cells[col_idx + 3].text = res_val

                # Color pass/fail
                if res_val == "P":
                    set_cell_background(row_cells[col_idx + 3], "C6EFCE")
                elif res_val in ["F", "A"]:
                    set_cell_background(row_cells[col_idx + 3], "FFC7CE")
            else:
                for off in range(4):
                    row_cells[col_idx + off].text = "-"
            col_idx += 4

        # Style data cells font size
        for cell in row_cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)

    doc.save(output_file_path)
    logger.info(f"Word .docx report successfully saved to {output_file_path}")


def export_to_pdf(results_data, output_file_path):
    """Export results to a PDF report with formatted tables."""
    if not HAS_REPORTLAB:
        raise ImportError("reportlab is required to export to .pdf")

    unique_subjects = extract_unique_subjects(results_data)
    total_cols = 3 + len(unique_subjects) * 4

    # Dynamically select page size (A4 landscape or A3 landscape for wide tables)
    page_size = landscape(A3) if total_cols > 25 else landscape(A4)

    doc = SimpleDocTemplate(
        output_file_path,
        pagesize=page_size,
        leftMargin=15,
        rightMargin=15,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PDFTitle",
        parent=styles["Title"],
        fontSize=15,
        textColor=colors.HexColor("#1F497D"),
        spaceAfter=4
    )
    meta_style = ParagraphStyle(
        "PDFMeta",
        parent=styles["Normal"],
        fontSize=8.5,
        textColor=colors.HexColor("#555555"),
        alignment=1, # Center
        spaceAfter=12
    )

    elements = [
        Paragraph("<b>Visvesvaraya Technological University - Student Results Report</b>", title_style),
        Paragraph(f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')} | Total Students: {len(results_data)}", meta_style)
    ]

    # Two-level Table Headers
    hdr1 = ["USN", "Student Name", "Sem"]
    hdr2 = ["", "", ""]
    for s in unique_subjects:
        hdr1.extend([s["code"], "", "", ""])
        hdr2.extend(["IA", "SEE", "TOT", "RES"])

    table_data = [hdr1, hdr2]

    # Table styling commands
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F497D")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#DCE6F1")),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#1F497D")),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (1, 2), (1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0C4DE")),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]

    # Span header 1 subject cells across 4 columns
    for idx in range(len(unique_subjects)):
        sc = 3 + idx * 4
        style_cmds.append(("SPAN", (sc, 0), (sc + 3, 0)))

    # Populate Data
    for row_idx, res in enumerate(results_data, start=2):
        row = [res.get("usn", ""), res.get("name", ""), res.get("semester", "")]
        stud_subs = {s.get("code") or s.get("name"): s for s in res.get("subjects", [])}
        
        # Alternating background
        if row_idx % 2 == 1:
            style_cmds.append(("BACKGROUND", (0, row_idx), (-1, row_idx), colors.HexColor("#F9FBFD")))

        for u_idx, u_sub in enumerate(unique_subjects):
            key = u_sub["code"]
            sc = 3 + u_idx * 4
            if key in stud_subs:
                s_data = stud_subs[key]
                row.extend([
                    s_data.get("ia", "-"),
                    s_data.get("see", "-"),
                    s_data.get("total", "-"),
                    s_data.get("res", "-")
                ])
                res_val = str(s_data.get("res", "")).upper()
                if res_val == "P":
                    style_cmds.append(("BACKGROUND", (sc + 3, row_idx), (sc + 3, row_idx), colors.HexColor("#C6EFCE")))
                elif res_val in ["F", "A"]:
                    style_cmds.append(("BACKGROUND", (sc + 3, row_idx), (sc + 3, row_idx), colors.HexColor("#FFC7CE")))
            else:
                row.extend(["-", "-", "-", "-"])
        table_data.append(row)

    # Calculate column widths
    col_widths = [65, 90, 50] + [19] * (len(unique_subjects) * 4)
    t = Table(table_data, colWidths=col_widths, repeatRows=2)
    t.setStyle(TableStyle(style_cmds))
    elements.append(t)

    doc.build(elements)
    logger.info(f"PDF report successfully saved to {output_file_path}")
