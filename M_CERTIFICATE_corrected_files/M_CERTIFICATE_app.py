import io
import os
import re
import zipfile
from datetime import date, datetime
from pathlib import Path

import streamlit as st
import openpyxl
import fitz
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


# ============================================================
# APP CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="M-CERTIFICATE | Smart Certificate Studio",
    page_icon="📄",
    layout="wide",
)

DEFAULTS = {
    "date_format": "DD-MMM-YY",
    "company": "OptoPixel Pvt. Ltd.",
    "paper_type": "Letterhead",
    "top_spacing_mm": 10,
    "show_make": True,
    "show_model": True,
    "show_serial": True,
    "show_mac": True,
    "show_win": True,
    "make_heading": "MAKE",
    "model_heading": "MODEL NO",
    "serial_heading": "SERIAL NO",
    "mac_heading": "MAC ID",
    "win_heading": "WIN KEY",
}

DATE_FORMATS = {
    "DD-MMM-YY": "%d-%b-%y",
    "DD-MMM-YYYY": "%d-%b-%Y",
    "DD/MM/YYYY": "%d/%m/%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
}

EMPTY_VALUES = {
    "", "n/a", "na", "n.a.", "not applicable",
    "-", "--", "—", "none", "null"
}

if "smart_settings" not in st.session_state:
    st.session_state.smart_settings = DEFAULTS.copy()

for key, value in DEFAULTS.items():
    st.session_state.smart_settings.setdefault(key, value)

st.session_state.setdefault("page", "M-SMART PDF")
st.session_state.setdefault("smart_generated", None)
st.session_state.setdefault("manual_generated", None)


# ============================================================
# PREMIUM UI
# ============================================================

st.markdown("""
<style>
.stApp {
    background: #000000 !important;
    color: #FFFFFF !important;
}

.block-container {
    max-width: 1500px;
    padding-top: 3.2rem;
}

.brand-title {
    text-align: center;
    font-size: 34px;
    font-weight: 850;
    letter-spacing: 2px;
    color: #FFFFFF !important;
}

.brand-subtitle {
    text-align: center;
    color: #FFFFFF !important;
    margin-bottom: 22px;
}

div[data-testid="stButton"] button {
    min-height: 45px;
    border-radius: 12px;
    font-weight: 700;
}

div[data-testid="stDownloadButton"] button {
    border-radius: 10px;
    font-weight: 700;
}

div[data-testid="stFileUploader"] {
    border-radius: 12px;
}

.stApp, .stApp p, .stApp span, .stApp label,
.stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp h5, .stApp h6,
[data-testid="stMarkdownContainer"],
[data-testid="stWidgetLabel"],
[data-testid="stCaptionContainer"] {
    color: #FFFFFF !important;
}

.stApp input,
.stApp textarea,
.stApp [data-baseweb="select"] {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="brand-title">M-CERTIFICATE</div>'
    '<div class="brand-subtitle">'
    'SMART CERTIFICATE STUDIO · EXCEL TO PDF + WORD'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# COMMON HELPERS
# ============================================================

def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def valid_value(value):
    value = clean(value)
    return "" if value.casefold() in EMPTY_VALUES else value


def safe_stem(value):
    value = re.sub(r"[^A-Za-z0-9.-]+", "_", str(value or ""))
    value = value.strip("._-")
    return value[:100] or "TEMP"


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = clean(value)

    formats = [
        "%d-%b-%y", "%d-%b-%Y",
        "%d/%m/%Y", "%d-%m-%Y",
        "%Y-%m-%d", "%m/%d/%Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    return date.today()


def item_filename(item):
    item_name = safe_stem(
        valid_value(item.get("item")) or "TEMP"
    ).upper()

    parts = ["TEST_CERTIFICATE", item_name]

    model = valid_value(item.get("model_no"))
    if model:
        parts.append(safe_stem(model).upper())

    serial = valid_value(item.get("serial_no"))
    if serial:
        parts.append(safe_stem(serial).upper())

    return "_".join(parts)


def combined_filename(company, cert_date):
    company_part = safe_stem(company).upper()

    if isinstance(cert_date, (date, datetime)):
        date_part = cert_date.strftime("%d-%b-%y").upper()
    else:
        date_part = safe_stem(cert_date).upper()

    return f"TEST_CERTIFICATE_{company_part}_{date_part}"


def merged_cell_value(ws, row, column):
    if not column:
        return None

    cell = ws.cell(row, column)

    if cell.value is not None:
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(
                merged_range.min_row,
                merged_range.min_col,
            ).value

    return None


def wrap_text(text, font, size, width):
    words = str(text or "").split()

    if not words:
        return [""]

    lines = []
    current = ""

    for word in words:
        candidate = word if not current else current + " " + word

        if stringWidth(candidate, font, size) <= width:
            current = candidate
        else:
            if current:
                lines.append(current)

            current = word

            if stringWidth(current, font, size) > width:
                chunk = ""

                for char in current:
                    if stringWidth(chunk + char, font, size) <= width:
                        chunk += char
                    else:
                        if chunk:
                            lines.append(chunk)
                        chunk = char

                current = chunk

    if current:
        lines.append(current)

    return lines or [""]


# ============================================================
# EXCEL READER
# ============================================================

def read_excel(uploaded_file):
    wb = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    ws = wb[wb.sheetnames[0]]

    company = ""
    date_raw = ""
    header_row = None

    for row in range(1, min(ws.max_row, 15) + 1):
        values = [
            clean(ws.cell(row, col).value)
            .lower().replace(".", "").strip()
            for col in range(1, ws.max_column + 1)
        ]

        for col in range(1, ws.max_column + 1):
            value = clean(ws.cell(row, col).value)

            if value.lower().startswith("company name"):
                company = re.sub(
                    r"^company name\s*:?\s*",
                    "",
                    value,
                    flags=re.I,
                ).strip()

            if value.lower().startswith("date"):
                date_raw = re.sub(
                    r"^date\s*:?\s*",
                    "",
                    value,
                    flags=re.I,
                ).strip()

        has_name = any(
            v in values
            for v in ("name", "item", "material", "item name")
        )

        has_model_or_serial = any(
            v in values
            for v in (
                "mn", "model", "model no", "model number",
                "sn", "serial", "serial no", "serial number",
            )
        )

        if has_name and has_model_or_serial:
            header_row = row
            break

    if header_row is None:
        raise ValueError(
            "Excel headings find nathi thata. "
            "Name/Item sathe Model/MN athva Serial/SN heading rakho."
        )

    headers = {}

    for col in range(1, ws.max_column + 1):
        heading = (
            clean(ws.cell(header_row, col).value)
            .lower().replace(".", "").strip()
        )

        if heading:
            headers[heading] = col

    def find_col(*names):
        for name in names:
            if name in headers:
                return headers[name]
        return None

    name_col = find_col(
        "name", "item", "material", "item name", "description"
    )
    make_col = find_col("make", "brand", "manufacturer")
    model_col = find_col(
        "mn", "model", "model no", "model number", "modelno"
    )
    serial_col = find_col(
        "sn", "serial", "serial no", "serial number", "serialno"
    )
    mac_col = find_col("mac", "mac id", "mac address")
    win_col = find_col(
        "win key", "windows key", "product key",
        "license key", "licence key",
    )
    qty_col = find_col("qty", "quantity")

    if not name_col:
        raise ValueError("Excel ma Name/Item column jaruri chhe.")

    items = []
    previous_item = None

    for row in range(header_row + 1, ws.max_row + 1):
        name = valid_value(
            merged_cell_value(ws, row, name_col)
        )

        make = valid_value(
            merged_cell_value(ws, row, make_col)
        )

        model = (
            valid_value(merged_cell_value(ws, row, model_col))
            if model_col else ""
        )

        serial = (
            valid_value(ws.cell(row, serial_col).value)
            if serial_col else ""
        )

        mac = (
            valid_value(ws.cell(row, mac_col).value)
            if mac_col else ""
        )

        win_key = (
            valid_value(ws.cell(row, win_col).value)
            if win_col else ""
        )

        # CHANGE ONLY:
        # Remove "(" and ")" from Model No., Serial No. and MAC ID.
        # No other characters are removed.
        model = model.replace("(", "").replace(")", "")
        serial = serial.replace("(", "").replace(")", "")
        mac = mac.replace("(", "").replace(")", "")

        qty_raw = ws.cell(row, qty_col).value if qty_col else 1

        try:
            qty = max(1, int(float(qty_raw or 1)))
        except (TypeError, ValueError):
            qty = 1

        # MAC ID sometimes appears inside the Serial cell.
        if serial:
            match = re.search(
                r"\bMAC(?:\s+ID)?\s*[:=]\s*([0-9A-Fa-f:-]{11,})",
                serial,
                flags=re.I,
            )

            if match:
                mac = mac or match.group(1)

                serial = re.sub(
                    r"\s*[,;]?\s*\bMAC(?:\s+ID)?\s*[:=]\s*"
                    r"[0-9A-Fa-f:-]{11,}",
                    "",
                    serial,
                    flags=re.I,
                ).strip()

        # Support a WIN continuation row when model contains WIN.
        if model and (
            model.upper().startswith("WIN ")
            or model.upper().startswith("WINDOWS ")
        ):
            if previous_item:
                previous_item["win_key"] = win_key or serial
            continue

        # Skip completely blank rows.
        if not any([name, make, model, serial, mac, win_key]):
            continue

        # Carry forward the previous item's identifying fields
        # when a row has a separate serial/MAC/license value.
        if not name:
            if previous_item and any([serial, mac, win_key]):
                name = previous_item.get("item", "")
                make = make or previous_item.get("make", "")
                model = model or previous_item.get("model_no", "")
            else:
                if previous_item and win_key:
                    previous_item["win_key"] = win_key
                continue

        if not name:
            continue

        item = {
            "item": name,
            "make": make,
            "model_no": model,
            "serial_no": serial,
            "mac": mac,
            "win_key": win_key,
            "qty": qty,
        }

        items.append(item)
        previous_item = item

    if not items:
        raise ValueError("Excel ma koi item row mali nathi.")

    company = company or "Customer"

    company = re.sub(
        r"\s*TVC\s*\d+.*$",
        "",
        company,
        flags=re.I,
    ).strip()

    company = re.sub(
        r"\s*\([^)]*\)\s*$",
        "",
        company,
    ).strip() or "Customer"

    return {
        "company": company,
        "date": parse_date(date_raw),
        "items": items,
    }


# ============================================================
# DYNAMIC TABLE COLUMNS
# ============================================================

def table_fields(item, settings):
    fields = [
        ("ITEM", valid_value(item.get("item"))),
    ]

    optional_fields = [
        (
            "make",
            settings.get("make_heading", "MAKE"),
            "show_make",
        ),
        (
            "model_no",
            settings.get("model_heading", "MODEL NO"),
            "show_model",
        ),
        (
            "serial_no",
            settings.get("serial_heading", "SERIAL NO"),
            "show_serial",
        ),
        (
            "mac",
            settings.get("mac_heading", "MAC ID"),
            "show_mac",
        ),
        (
            "win_key",
            settings.get("win_heading", "WIN KEY"),
            "show_win",
        ),
    ]

    for key, heading, setting_key in optional_fields:
        value = valid_value(item.get(key))

        if settings.get(setting_key, True) and value:
            fields.append((heading, value))

    return fields


# ============================================================
# PDF GENERATION
# ============================================================

def draw_certificate_page(pdf, item, company, cert_date, settings):
    page_w, page_h = A4

    if settings.get("paper_type") == "Letterhead":
        template_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "optopixel_template.png",
        )

        if os.path.isfile(template_path):
            pdf.drawImage(
                template_path,
                0, 0,
                width=page_w,
                height=page_h,
                preserveAspectRatio=False,
                mask="auto",
            )
        else:
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(48, page_h - 35, "OptoPixel")

    pdf.saveState()

    shift = (
        float(settings.get("top_spacing_mm", 10))
        * 72 / 25.4
    )

    pdf.translate(0, -shift)

    left = 48
    right = page_w - 48
    content_w = right - left

    fmt = DATE_FORMATS.get(
        settings.get("date_format"),
        "%d-%b-%y",
    )

    if isinstance(cert_date, (date, datetime)):
        date_text = cert_date.strftime(fmt)
    else:
        date_text = str(cert_date)

    pdf.setFillColorRGB(0.08, 0.08, 0.08)

    pdf.setFont("Helvetica", 10)
    pdf.drawRightString(
        right,
        page_h - 145,
        "Date: " + date_text,
    )

    title = "TEST CERTIFICATE"
    title_y = page_h - 95

    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawCentredString(page_w / 2, title_y, title)

    title_width = stringWidth(
        title,
        "Helvetica-Bold",
        16,
    )

    pdf.line(
        page_w / 2 - title_width / 2,
        title_y - 3,
        page_w / 2 + title_width / 2,
        title_y - 3,
    )

    concern = "TO WHOM IT MAY CONCERN"
    concern_y = page_h - 190

    pdf.setFont("Helvetica", 12)
    pdf.drawCentredString(
        page_w / 2,
        concern_y,
        concern,
    )

    concern_width = stringWidth(
        concern,
        "Helvetica",
        12,
    )

    pdf.line(
        page_w / 2 - concern_width / 2,
        concern_y - 3,
        page_w / 2 + concern_width / 2,
        concern_y - 3,
    )

    item_name = valid_value(item.get("item")) or "TEMP"
    fields = table_fields(item, settings)

    present_labels = [
        heading
        for heading, value in fields[1:]
        if value
    ]

    intro = f"This test certificate is for {item_name}"

    if present_labels:
        intro += " having " + ", ".join(present_labels)

    intro += " as per below table"

    intro_lines = wrap_text(
        intro,
        "Helvetica",
        10.5,
        content_w,
    )

    intro_y = page_h - 245

    pdf.setFont("Helvetica", 10.5)

    for index, line in enumerate(intro_lines):
        pdf.drawString(
            left,
            intro_y - index * 15,
            line,
        )

    count = len(fields)

    if count == 1:
        widths = [content_w]
    else:
        widths = (
            [content_w * 0.25]
            + [content_w * 0.75 / (count - 1)] * (count - 1)
        )

    wrapped = []
    max_header = 1
    max_value = 1

    for index, (heading, value) in enumerate(fields):
        available = widths[index] - 8

        header_lines = wrap_text(
            heading,
            "Helvetica-Bold",
            8.5,
            available,
        )

        value_lines = wrap_text(
            value,
            "Helvetica",
            8.5,
            available,
        )

        wrapped.append((header_lines, value_lines))

        max_header = max(max_header, len(header_lines))
        max_value = max(max_value, len(value_lines))

    header_h = max(30, max_header * 11 + 12)
    row_h = max(34, max_value * 11 + 12)

    table_top = (
        intro_y - len(intro_lines) * 15 - 25
    )

    table_bottom = table_top - header_h - row_h

    xs = [left]

    for width in widths:
        xs.append(xs[-1] + width)

    pdf.setLineWidth(0.6)

    pdf.line(left, table_top, right, table_top)
    pdf.line(
        left,
        table_top - header_h,
        right,
        table_top - header_h,
    )
    pdf.line(left, table_bottom, right, table_bottom)

    for x in xs:
        pdf.line(x, table_top, x, table_bottom)

    for index, (header_lines, value_lines) in enumerate(wrapped):
        center_x = (xs[index] + xs[index + 1]) / 2

        y = (
            table_top
            - (header_h - len(header_lines) * 11) / 2
            - 8
        )

        pdf.setFont("Helvetica-Bold", 8.5)

        for line in header_lines:
            pdf.drawCentredString(center_x, y, line)
            y -= 11

        y = (
            table_top - header_h
            - (row_h - len(value_lines) * 11) / 2
            - 8
        )

        pdf.setFont("Helvetica", 8.5)

        for line in value_lines:
            pdf.drawCentredString(center_x, y, line)
            y -= 11

    y = table_bottom - 35

    statement = (
        "Specified component is tested for all its functionalities "
        "and is performing satisfactorily and found OK."
    )

    pdf.setFont("Helvetica", 10)

    for line in wrap_text(
        statement, "Helvetica", 10, content_w
    ):
        pdf.drawString(left, y, line)
        y -= 15

    y -= 20

    pdf.setFont("Helvetica-Oblique", 10)

    sold_to = company or "Customer"

    sold_text = (
        f"The above said item is sold to {sold_to}."
    )

    for line in wrap_text(
        sold_text, "Helvetica-Oblique", 10, content_w
    ):
        pdf.drawString(left, y, line)
        y -= 15

    y -= 20

    pdf.setFont("Helvetica", 10)

    pdf.drawString(
        left,
        y,
        f'For {settings.get("company", DEFAULTS["company"])},',
    )

    pdf.drawString(left, y - 45, "Authorized")

    pdf.restoreState()
    pdf.showPage()


def create_pdf(data, settings):
    output = io.BytesIO()

    pdf = canvas.Canvas(
        output,
        pagesize=A4,
    )

    for item in data["items"]:
        copies = max(1, int(item.get("qty", 1)))

        item_company = item.get("_company", data["company"])
        item_date = item.get("_date", data["date"])

        for _ in range(copies):
            draw_certificate_page(
                pdf,
                item,
                item_company,
                item_date,
                settings,
            )

    pdf.save()

    return output.getvalue()


# ============================================================
# WORD GENERATION
# ============================================================

def create_word(data, settings):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    first_page = True

    for item in data["items"]:
        copies = max(1, int(item.get("qty", 1)))

        company = item.get("_company", data["company"])
        cert_date = item.get("_date", data["date"])

        for _ in range(copies):
            if not first_page:
                doc.add_page_break()

            first_page = False

            fmt = DATE_FORMATS.get(
                settings.get("date_format"),
                "%d-%b-%y",
            )

            if isinstance(cert_date, (date, datetime)):
                date_text = cert_date.strftime(fmt)
            else:
                date_text = str(cert_date)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run("Date: " + date_text)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = p.add_run("TEST CERTIFICATE")
            run.bold = True
            run.font.size = Pt(16)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = p.add_run("TO WHOM IT MAY CONCERN")
            run.bold = True
            run.font.size = Pt(11)

            item_name = valid_value(item.get("item")) or "TEMP"

            doc.add_paragraph(
                f"This test certificate is for {item_name} as per below table."
            )

            fields = table_fields(item, settings)

            table = doc.add_table(
                rows=2,
                cols=len(fields),
            )

            table.style = "Table Grid"
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            for col_index, (heading, value) in enumerate(fields):
                table.cell(0, col_index).text = heading
                table.cell(1, col_index).text = value

                for row_index in (0, 1):
                    cell = table.cell(row_index, col_index)
                    cell.vertical_alignment = (
                        WD_CELL_VERTICAL_ALIGNMENT.CENTER
                    )

                    for paragraph in cell.paragraphs:
                        paragraph.alignment = (
                            WD_ALIGN_PARAGRAPH.CENTER
                        )

                        for run in paragraph.runs:
                            run.font.size = Pt(8)

                for run in table.cell(
                    0, col_index
                ).paragraphs[0].runs:
                    run.bold = True

            doc.add_paragraph(
                "Specified component is tested for all its functionalities "
                "and is performing satisfactorily and found OK."
            )

            doc.add_paragraph(
                f"The above said item is sold to {company}."
            )

            doc.add_paragraph(
                f'For {settings.get("company", DEFAULTS["company"])},'
            )

            doc.add_paragraph("\nAuthorized")

    output = io.BytesIO()
    doc.save(output)

    return output.getvalue()


# ============================================================
# LOCAL WINDOWS AUTO-SAVE
# ============================================================

def save_files_locally(files, company, cert_date):
    if os.name != "nt":
        return None

    if isinstance(cert_date, (date, datetime)):
        date_folder = cert_date.strftime("%d-%b-%y")
    else:
        date_folder = safe_stem(cert_date)

    company_folder = safe_stem(company)

    candidates = [
        r"D:\M-CERTIFICATE",
        r"C:\M-CERTIFICATE",
    ]

    for root in candidates:
        try:
            os.makedirs(root, exist_ok=True)

            target = os.path.join(
                root,
                date_folder,
                company_folder,
            )

            os.makedirs(target, exist_ok=True)

            for filename, content in files.items():
                with open(
                    os.path.join(target, filename),
                    "wb",
                ) as handle:
                    handle.write(content)

            return target

        except OSError:
            continue

    return None


# ============================================================
# GENERATE INDIVIDUAL + COMBINED FILES
# ============================================================

def make_entries(parsed, settings):
    entries = []
    counts = {}

    for original in parsed["items"]:
        copies = max(1, int(original.get("qty", 1)))

        for _ in range(copies):
            item = dict(original)

            item["_company"] = item.get(
                "_company", parsed["company"]
            )

            item["_date"] = item.get(
                "_date", parsed["date"]
            )

            stem = item_filename(item)

            counts[stem] = counts.get(stem, 0) + 1

            if counts[stem] > 1:
                stem = f"{stem}_{counts[stem]}"

            one_data = {
                "company": item["_company"],
                "date": item["_date"],
                "items": [item],
            }

            pdf_bytes = create_pdf(one_data, settings)
            word_bytes = create_word(one_data, settings)

            entries.append({
                "label": (
                    f'{item.get("item") or "TEMP"} | '
                    f'Model: {item.get("model_no") or "—"} | '
                    f'Serial: {item.get("serial_no") or "—"}'
                ),
                "stem": stem,
                "pdf_name": stem + ".pdf",
                "docx_name": stem + ".docx",
                "pdf": pdf_bytes,
                "docx": word_bytes,
                "item": item,
            })

    return entries


def build_bundle(entries, company, cert_date, settings):
    combined_items = []

    for entry in entries:
        item = dict(entry["item"])
        item["qty"] = 1
        combined_items.append(item)

    combined_data = {
        "company": company,
        "date": cert_date,
        "items": combined_items,
    }

    combined_pdf = create_pdf(combined_data, settings)
    combined_word = create_word(combined_data, settings)

    combined_stem = combined_filename(company, cert_date)

    all_files = {}

    for entry in entries:
        all_files[entry["pdf_name"]] = entry["pdf"]
        all_files[entry["docx_name"]] = entry["docx"]

    all_files[combined_stem + ".pdf"] = combined_pdf
    all_files[combined_stem + ".docx"] = combined_word

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED,
    ) as archive:
        for filename, content in all_files.items():
            archive.writestr(filename, content)

    saved_folder = save_files_locally(
        all_files,
        company,
        cert_date,
    )

    return {
        "entries": entries,
        "combined_pdf": combined_pdf,
        "combined_docx": combined_word,
        "combined_pdf_name": combined_stem + ".pdf",
        "combined_docx_name": combined_stem + ".docx",
        "zip": zip_buffer.getvalue(),
        "zip_name": combined_stem + "_ALL_FILES.zip",
        "company": company,
        "date": cert_date,
        "saved_folder": saved_folder,
    }


# ============================================================
# DOWNLOAD PAGE
# ============================================================

def show_downloads(generated, key_prefix):
    if not generated:
        st.info(
            "Pehla M-SMART PDF athva M-CERTIFICATE page par "
            "certificate generate karo."
        )
        return

    st.subheader("Combined downloads")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.download_button(
            "⬇ DOWNLOAD ALL — ONE PDF",
            data=generated["combined_pdf"],
            file_name=generated["combined_pdf_name"],
            mime="application/pdf",
            key=f"{key_prefix}_all_pdf",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "⬇ DOWNLOAD ALL — ONE WORD",
            data=generated["combined_docx"],
            file_name=generated["combined_docx_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{key_prefix}_all_word",
            use_container_width=True,
        )

    with c3:
        st.download_button(
            "⬇ DOWNLOAD ZIP — ALL FILES",
            data=generated["zip"],
            file_name=generated["zip_name"],
            mime="application/zip",
            key=f"{key_prefix}_all_zip",
            use_container_width=True,
        )

    st.divider()
    st.subheader("Individual PDF and Word files")

    for index, entry in enumerate(generated["entries"]):
        with st.container(border=True):
            left, mid, right = st.columns([3, 1, 1])

            with left:
                st.markdown(
                    f"**{entry['item'].get('item') or 'TEMP'}**"
                )
                st.caption(entry["pdf_name"])

            with mid:
                st.download_button(
                    "PDF",
                    data=entry["pdf"],
                    file_name=entry["pdf_name"],
                    mime="application/pdf",
                    key=f"{key_prefix}_pdf_{index}",
                    use_container_width=True,
                )

            with right:
                st.download_button(
                    "WORD",
                    data=entry["docx"],
                    file_name=entry["docx_name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"{key_prefix}_word_{index}",
                    use_container_width=True,
                )

    st.divider()
    st.subheader("Download only selected certificates")

    labels = [
        f"{index + 1}. {entry['label']}"
        for index, entry in enumerate(generated["entries"])
    ]

    selected = st.multiselect(
        "Select only the certificates you need",
        options=labels,
        key=f"{key_prefix}_selected",
    )

    if selected:
        chosen = [
            entry
            for index, entry in enumerate(generated["entries"])
            if labels[index] in selected
        ]

        selected_pdf_data = {
            "company": generated["company"],
            "date": generated["date"],
            "items": [],
        }

        for entry in chosen:
            item = dict(entry["item"])
            item["qty"] = 1
            selected_pdf_data["items"].append(item)

        selected_pdf = create_pdf(
            selected_pdf_data,
            st.session_state.smart_settings,
        )

        selected_word = create_word(
            selected_pdf_data,
            st.session_state.smart_settings,
        )

        selected_stem = (
            combined_filename(
                generated["company"],
                generated["date"],
            )
            + f"_SELECTED_{len(chosen)}"
        )

        c1, c2, c3 = st.columns(3)

        with c1:
            st.download_button(
                "Download selected — PDF",
                selected_pdf,
                selected_stem + ".pdf",
                "application/pdf",
                key=f"{key_prefix}_selected_pdf",
                use_container_width=True,
            )

        with c2:
            st.download_button(
                "Download selected — Word",
                selected_word,
                selected_stem + ".docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"{key_prefix}_selected_word",
                use_container_width=True,
            )

        with c3:
            selected_zip = io.BytesIO()

            with zipfile.ZipFile(
                selected_zip,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as archive:
                for entry in chosen:
                    archive.writestr(
                        entry["pdf_name"],
                        entry["pdf"],
                    )
                    archive.writestr(
                        entry["docx_name"],
                        entry["docx"],
                    )

            st.download_button(
                "Download selected — ZIP",
                selected_zip.getvalue(),
                selected_stem + ".zip",
                "application/zip",
                key=f"{key_prefix}_selected_zip",
                use_container_width=True,
            )


# ============================================================
# NAVIGATION
# ============================================================

nav_cols = st.columns(4)

pages = [
    "M-SMART PDF",
    "DOWNLOAD PDF WORD",
    "M-CERTIFICATE",
    "PAGE SETTINGS",
]

for col, page_name in zip(nav_cols, pages):
    with col:
        if st.button(
            page_name,
            key="nav_" + page_name,
            type=(
                "primary"
                if st.session_state.page == page_name
                else "secondary"
            ),
            use_container_width=True,
        ):
            st.session_state.page = page_name
            st.rerun()

st.divider()

settings = st.session_state.smart_settings


# ============================================================
# PAGE 1 — EXCEL TO CERTIFICATES
# ============================================================

if st.session_state.page == "M-SMART PDF":
    st.header("M-SMART PDF")
    st.write(
        "Excel upload karo. Darek item mate alag PDF ane editable "
        "Word file generate thase."
    )

    uploads = st.file_uploader(
        "Upload Excel files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="excel_upload",
    )

    if uploads:
        try:
            parsed_files = [
                read_excel(uploaded)
                for uploaded in uploads
            ]

            all_items = []

            for parsed_file in parsed_files:
                for original in parsed_file["items"]:
                    item = dict(original)
                    item["_company"] = parsed_file["company"]
                    item["_date"] = parsed_file["date"]
                    all_items.append(item)

            company = parsed_files[0]["company"]
            cert_date = parsed_files[0]["date"]

            parsed = {
                "company": company,
                "date": cert_date,
                "items": all_items,
            }

            total_certificates = sum(
                max(1, int(item.get("qty", 1)))
                for item in all_items
            )

            st.success(
                f"{len(uploads)} Excel file(s) read. "
                f"Total item rows: {len(all_items)}. "
                f"Certificates to generate: {total_certificates}."
            )

            st.subheader("Excel data preview")

            st.dataframe(
                [
                    {
                        "Item": item.get("item", ""),
                        "Make": item.get("make", ""),
                        "Model No": item.get("model_no", ""),
                        "Serial No": item.get("serial_no", ""),
                        "MAC ID": item.get("mac", ""),
                        "WIN KEY": item.get("win_key", ""),
                        "Qty": item.get("qty", 1),
                    }
                    for item in all_items
                ],
                use_container_width=True,
                hide_index=True,
            )

            if st.button(
                "✨ GENERATE CERTIFICATES",
                type="primary",
                use_container_width=True,
            ):
                entries = make_entries(parsed, settings)

                st.session_state.smart_generated = build_bundle(
                    entries,
                    company,
                    cert_date,
                    settings,
                )

                st.success(
                    f"{len(entries)} individual certificates generated."
                )

                saved = st.session_state.smart_generated["saved_folder"]

                if saved:
                    st.info("Files saved locally: " + saved)

                st.rerun()

        except Exception as exc:
            st.error(f"Excel error: {exc}")

    generated = st.session_state.smart_generated

    if generated:
        st.divider()
        st.subheader("Generated certificate preview")

        st.caption(
            "PDF preview niche dekho. Downloads mate "
            "DOWNLOAD PDF WORD page kholo."
        )

        try:
            pdf_doc = fitz.open(
                stream=generated["combined_pdf"],
                filetype="pdf",
            )

            for index, page in enumerate(pdf_doc):
                with st.expander(
                    f"Certificate preview — Page {index + 1}",
                    expanded=(index == 0),
                ):
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(1.1, 1.1),
                        alpha=False,
                    )

                    image = Image.open(
                        io.BytesIO(pix.tobytes("png"))
                    )

                    st.image(
                        image,
                        use_container_width=True,
                    )

            pdf_doc.close()

        except Exception as exc:
            st.warning(f"Preview error: {exc}")

        st.dataframe(
            [
                {
                    "Item": entry["item"].get("item", ""),
                    "PDF filename": entry["pdf_name"],
                    "Word filename": entry["docx_name"],
                }
                for entry in generated["entries"]
            ],
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# PAGE 2 — DOWNLOADS
# ============================================================

elif st.session_state.page == "DOWNLOAD PDF WORD":
    st.header("DOWNLOAD PDF WORD")

    show_downloads(
        st.session_state.smart_generated,
        "smart",
    )

    generated = st.session_state.smart_generated

    if generated and generated["saved_folder"]:
        st.success(
            "PC par files saved: " + generated["saved_folder"]
        )
    elif generated:
        st.info(
            "D:/C: local auto-save Windows par Streamlit locally "
            "run karsho tyare j available chhe. Cloud par download "
            "buttons vapro."
        )


# ============================================================
# PAGE 3 — MANUAL CERTIFICATE
# ============================================================

elif st.session_state.page == "M-CERTIFICATE":
    st.header("M-CERTIFICATE")
    st.write(
        "Excel vagar manually certificate banao. Khali fields "
        "table ma dekhase nahi."
    )

    with st.form("manual_certificate_form"):
        left, right = st.columns(2)

        with left:
            item_name = st.text_input(
                "Item name",
                placeholder="CAMERA / LENSE",
            )

            make = st.text_input("Make")
            model = st.text_input("Model No")
            serial = st.text_input("Serial No")

        with right:
            company = st.text_input(
                "Sold To / Company",
                value="Customer",
            )

            cert_date = st.date_input(
                "Certificate date",
                value=date.today(),
            )

            mac = st.text_input("MAC ID")
            win_key = st.text_input("WIN KEY")

        st.caption(
            "Model No, Serial No, Make, MAC ID ane WIN KEY "
            "optional chhe."
        )

        submit_manual = st.form_submit_button(
            "GENERATE MANUAL CERTIFICATE",
            type="primary",
            use_container_width=True,
        )

    if submit_manual:
        if not item_name.strip():
            st.error("Item name jaruri chhe.")
        else:
            company = company.strip() or "Customer"

            item = {
                "item": item_name.strip(),
                "make": valid_value(make),
                "model_no": valid_value(model).replace("(", "").replace(")", ""),
                "serial_no": valid_value(serial).replace("(", "").replace(")", ""),
                "mac": valid_value(mac).replace("(", "").replace(")", ""),
                "win_key": valid_value(win_key),
                "qty": 1,
                "_company": company,
                "_date": cert_date,
            }

            parsed = {
                "company": company,
                "date": cert_date,
                "items": [item],
            }

            entries = make_entries(parsed, settings)

            st.session_state.manual_generated = build_bundle(
                entries,
                company,
                cert_date,
                settings,
            )

            st.success("Manual certificate generated.")
            st.rerun()

    if st.session_state.manual_generated:
        generated = st.session_state.manual_generated

        st.subheader("Manual certificate preview")

        try:
            pdf_doc = fitz.open(
                stream=generated["combined_pdf"],
                filetype="pdf",
            )

            page = pdf_doc[0]

            pix = page.get_pixmap(
                matrix=fitz.Matrix(1.2, 1.2),
                alpha=False,
            )

            image = Image.open(
                io.BytesIO(pix.tobytes("png"))
            )

            st.image(
                image,
                use_container_width=False,
            )

            pdf_doc.close()

        except Exception as exc:
            st.warning(f"Preview error: {exc}")

        show_downloads(
            generated,
            "manual",
        )


# ============================================================
# PAGE 4 — SETTINGS
# ============================================================

elif st.session_state.page == "PAGE SETTINGS":
    st.header("PAGE SETTINGS")

    st.info(
        "Aa settings current Streamlit session ma auto-save thase. "
        "App bandh kari fari kholsho to settings permanent rehse "
        "teni guarantee aa version ma nathi."
    )

    settings = st.session_state.smart_settings

    with st.form("settings_form"):
        left, right = st.columns(2)

        with left:
            date_format = st.selectbox(
                "Date format",
                list(DATE_FORMATS),
                index=list(DATE_FORMATS).index(
                    settings["date_format"]
                ),
            )

            paper_type = st.selectbox(
                "Paper type",
                ["Letterhead", "Blank paper"],
                index=[
                    "Letterhead", "Blank paper"
                ].index(settings["paper_type"]),
            )

            top_spacing = st.number_input(
                "Content top spacing (mm)",
                min_value=0,
                max_value=60,
                value=int(settings["top_spacing_mm"]),
                step=1,
            )

            issuing_company = st.text_input(
                "Certificate issued by",
                value=settings["company"],
            )

        with right:
            make_heading = st.text_input(
                "Make heading",
                value=settings["make_heading"],
            )

            model_heading = st.text_input(
                "Model heading",
                value=settings["model_heading"],
            )

            serial_heading = st.text_input(
                "Serial heading",
                value=settings["serial_heading"],
            )

            mac_heading = st.text_input(
                "MAC heading",
                value=settings["mac_heading"],
            )

            win_heading = st.text_input(
                "WIN KEY heading",
                value=settings["win_heading"],
            )

        st.subheader("Optional table columns")

        a, b, c, d, e = st.columns(5)

        with a:
            show_make = st.checkbox(
                "Show Make",
                value=settings["show_make"],
            )

        with b:
            show_model = st.checkbox(
                "Show Model No",
                value=settings["show_model"],
            )

        with c:
            show_serial = st.checkbox(
                "Show Serial No",
                value=settings["show_serial"],
            )

        with d:
            show_mac = st.checkbox(
                "Show MAC ID",
                value=settings["show_mac"],
            )

        with e:
            show_win = st.checkbox(
                "Show WIN KEY",
                value=settings["show_win"],
            )

        save_settings = st.form_submit_button(
            "SAVE SETTINGS",
            type="primary",
            use_container_width=True,
        )

    if save_settings:
        st.session_state.smart_settings.update({
            "date_format": date_format,
            "paper_type": paper_type,
            "top_spacing_mm": int(top_spacing),
            "company": issuing_company.strip() or DEFAULTS["company"],
            "make_heading": make_heading.strip() or "MAKE",
            "model_heading": model_heading.strip() or "MODEL NO",
            "serial_heading": serial_heading.strip() or "SERIAL NO",
            "mac_heading": mac_heading.strip() or "MAC ID",
            "win_heading": win_heading.strip() or "WIN KEY",
            "show_make": show_make,
            "show_model": show_model,
            "show_serial": show_serial,
            "show_mac": show_mac,
            "show_win": show_win,
        })

        st.session_state.smart_generated = None
        st.session_state.manual_generated = None

        st.success("Settings saved for this session.")
        st.rerun()

    if st.button("RESTORE DEFAULT SETTINGS"):
        st.session_state.smart_settings = DEFAULTS.copy()
        st.session_state.smart_generated = None
        st.session_state.manual_generated = None
        st.rerun()


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.markdown(
    '<div style="text-align:center;color:#728096;font-size:12px">'
    'M-CERTIFICATE · SMART CERTIFICATE STUDIO'
    '</div>',
    unsafe_allow_html=True,
)
