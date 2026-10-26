
import io
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.utils import ImageReader


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="🖨️",
    layout="wide"
)

st.markdown("""
<style>
.stApp {
    background: #0b1220;
    color: #eef2ff;
}
.block-container {
    max-width: 1450px;
    padding-top: 1.2rem;
}
h1, h2, h3, h4, label, p {
    color: #eef2ff !important;
}
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("A4 Test Certificate • PDF • Linux CUPS Printing • Multi-Page Scanner")


# ============================================================
# PATHS AND FILE HELPERS
# ============================================================

def get_mcertificate_root():
    """Save under ~/M-CERTIFICATE on Linux."""
    root = Path.home() / "M-CERTIFICATE"
    root.mkdir(parents=True, exist_ok=True)
    return root


def safe_folder_name(name, default="Customer"):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "", str(name).strip())
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or default


def safe_filename(name, default="document"):
    name = safe_folder_name(name, default)
    name = re.sub(r"\s+", "_", name)
    if not name:
        name = default
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name


def get_certificate_folder(cert_date, sold_to):
    root = get_mcertificate_root()
    date_folder = cert_date.strftime("%d-%b-%y")
    customer_folder = safe_folder_name(sold_to)
    folder = root / date_folder / customer_folder
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def parse_date(value):
    for fmt in ("%d-%b-%y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            pass
    return None


# ============================================================
# LINUX PRINTERS - CUPS
# ============================================================

def get_printers():
    try:
        result = subprocess.run(
            ["lpstat", "-p"],
            capture_output=True,
            text=True,
            timeout=10
        )
        printers = []

        for line in result.stdout.splitlines():
            match = re.match(r"printer\s+(\S+)", line)
            if match:
                printers.append(match.group(1))

        return sorted(set(printers))

    except (FileNotFoundError, subprocess.TimeoutExpired):
        return []


def get_default_printer():
    try:
        result = subprocess.run(
            ["lpstat", "-d"],
            capture_output=True,
            text=True,
            timeout=10
        )
        match = re.search(
            r"system default destination:\s*(\S+)",
            result.stdout
        )
        return match.group(1) if match else ""
    except Exception:
        return ""


def print_pdf(pdf_bytes, printer, copies=1):
    if not printer:
        raise RuntimeError("Please select a printer.")

    fd, temp_path = tempfile.mkstemp(suffix=".pdf")

    try:
        with os.fdopen(fd, "wb") as f:
            f.write(pdf_bytes)

        result = subprocess.run(
            [
                "lp",
                "-d", printer,
                "-n", str(int(copies)),
                "-o", "media=A4",
                temp_path
            ],
            capture_output=True,
            text=True,
            timeout=90
        )

        if result.returncode != 0:
            raise RuntimeError(
                result.stderr.strip() or "CUPS printing failed."
            )

        return result.stdout.strip()

    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass


# ============================================================
# LINUX SCANNER - SANE
# ============================================================

def find_scanners():
    try:
        result = subprocess.run(
            ["scanimage", "-L"],
            capture_output=True,
            text=True,
            timeout=20
        )

        output = result.stdout + "\n" + result.stderr
        scanners = []

        for line in output.splitlines():
            line = line.strip()

            # Typical output:
            # device `escl:http://.../eSCL' is a ...
            match = re.match(r"device\s+`(.+?)'\s+is a\s+(.+)", line)

            if match:
                scanners.append({
                    "device": match.group(1),
                    "name": match.group(2)
                })

        return scanners

    except FileNotFoundError:
        raise RuntimeError(
            "scanimage is not installed. Install sane-utils."
        )


def scan_one_page(device, resolution=200, mode="Color"):
    """Scan one page. Returns image bytes in memory."""
    cmd = [
        "scanimage",
        "--device-name", device,
        "--format=jpeg",
        "--resolution", str(int(resolution))
    ]

    if mode == "Gray":
        cmd += ["--mode", "Gray"]
    elif mode == "Lineart":
        cmd += ["--mode", "Lineart"]
    else:
        cmd += ["--mode", "Color"]

    result = subprocess.run(
        cmd,
        capture_output=True,
        timeout=180
    )

    if result.returncode != 0 or not result.stdout:
        error = result.stderr.decode(errors="replace").strip()
        raise RuntimeError(error or "Scanner returned no image data.")

    # Verify the image before accepting it.
    try:
        with Image.open(io.BytesIO(result.stdout)) as image:
            image.verify()
    except Exception as exc:
        raise RuntimeError(f"Invalid scanner image: {exc}")

    return result.stdout


# ============================================================
# CERTIFICATE PDF
# ============================================================

def wrap_words(text, font, size, max_width):
    lines = []
    current = ""

    for word in str(text).split():
        test = word if not current else current + " " + word

        if stringWidth(test, font, size) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines


def make_certificate_pdf(data):
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    width, height = A4

    pdf.setFont("Helvetica", 11)
    pdf.drawRightString(
        width - 55,
        height - 100,
        "Date: " + data["date"].strftime("%d-%b-%y")
    )

    title = "TEST CERTIFICATE"
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawCentredString(width / 2, height - 145, title)

    tw = stringWidth(title, "Helvetica-Bold", 16)
    pdf.line(
        width / 2 - tw / 2, height - 149,
        width / 2 + tw / 2, height - 149
    )

    subtitle = "TO WHOM IT MAY CONCERN"
    pdf.setFont("Helvetica", 12)
    pdf.drawCentredString(width / 2, height - 190, subtitle)

    labels = [label for _, label in data["selected_columns"]]

    if not labels:
        intro = (
            f"This test certificate is for {data['item_name']} "
            "as per below table"
        )
    else:
        intro = (
            f"This test certificate is for {data['item_name']} "
            f"having {', '.join(labels)} as per below table"
        )

    pdf.setFont("Helvetica", 11)
    intro_y = height - 230

    for i, line in enumerate(wrap_words(intro, "Helvetica", 11, 490)[:3]):
        pdf.drawString(55, intro_y - i * 16, line)

    columns = [("ITEM", data["item_name"])]

    for key, label in data["selected_columns"]:
        columns.append((label, data[key]))

    left = 55
    right = width - 55
    top = height - 300
    header_h = 30
    row_h = 32
    bottom = top - header_h - row_h
    col_width = (right - left) / len(columns)

    pdf.setLineWidth(0.6)

    for i in range(len(columns) + 1):
        x = left + i * col_width
        pdf.line(x, top, x, bottom)

    for y in (top, top - header_h, bottom):
        pdf.line(left, y, right, y)

    for i, (label, value) in enumerate(columns):
        x1 = left + i * col_width
        x2 = x1 + col_width
        center = (x1 + x2) / 2

        pdf.setFont("Helvetica-Bold", 9)
        pdf.drawCentredString(center, top - 19, str(label)[:40])

        value_text = str(value)
        font_size = 10

        while (
            font_size > 6
            and stringWidth(value_text, "Helvetica", font_size)
            > col_width - 8
        ):
            font_size -= 0.5

        pdf.setFont("Helvetica", font_size)
        pdf.drawCentredString(center, bottom + 11, value_text)

    body_y = height - 375
    pdf.setFont("Helvetica", 11)

    pdf.drawString(
        55, body_y,
        "Specified component is tested for all its functionalities and is performing"
    )
    pdf.drawString(
        55, body_y - 18,
        "satisfactorily and found OK."
    )

    sold_text = f"The above said item is sold to {data['sold_to']}."
    pdf.setFont("Helvetica-Oblique", 11)

    for i, line in enumerate(
        wrap_words(sold_text, "Helvetica-Oblique", 11, 490)[:3]
    ):
        pdf.drawString(55, body_y - 60 - i * 16, line)

    pdf.setFont("Helvetica", 11)
    pdf.drawString(55, height - 525, "For OptoPixel Pvt. Ltd.,")
    pdf.drawString(55, height - 580, "Authorized")

    pdf.showPage()
    pdf.save()

    output.seek(0)
    return output.getvalue()


# ============================================================
# WORD CERTIFICATE
# ============================================================

def certificate_word_bytes(data):
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Date: " + data["date"].strftime("%d-%b-%y"))
    run.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("TEST CERTIFICATE")
    run.bold = True
    run.font.size = Pt(16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("TO WHOM IT MAY CONCERN")

    labels = [label for _, label in data["selected_columns"]]
    intro = f"This test certificate is for {data['item_name']}"

    if labels:
        intro += " having " + ", ".join(labels)

    intro += " as per below table"
    doc.add_paragraph(intro)

    columns = [("ITEM", data["item_name"])]
    for key, label in data["selected_columns"]:
        columns.append((label, data[key]))

    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"

    for i, (label, _) in enumerate(columns):
        table.rows[0].cells[i].text = str(label)

    row = table.add_row()
    for i, (_, value) in enumerate(columns):
        row.cells[i].text = str(value)

    doc.add_paragraph(
        "Specified component is tested for all its functionalities "
        "and is performing satisfactorily and found OK."
    )
    doc.add_paragraph(
        f"The above said item is sold to {data['sold_to']}."
    )
    doc.add_paragraph()
    doc.add_paragraph("For OptoPixel Pvt. Ltd.,")
    doc.add_paragraph()
    doc.add_paragraph("Authorized")

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def save_certificate_files(data, pdf_bytes, filename):
    folder = get_certificate_folder(data["date"], data["sold_to"])

    base = safe_folder_name(filename, "test_certificate")
    base = os.path.splitext(base)[0]

    pdf_path = folder / (base + ".pdf")
    docx_path = folder / (base + ".docx")

    pdf_path.write_bytes(pdf_bytes)
    docx_path.write_bytes(certificate_word_bytes(data))

    return folder, pdf_path, docx_path


def make_scanned_pdf(pages):
    if not pages:
        raise RuntimeError("No scanned pages available.")

    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    page_w, page_h = A4
    margin = 20

    for page_bytes in pages:
        with Image.open(io.BytesIO(page_bytes)) as source:
            image = source.convert("RGB")
            img_w, img_h = image.size

            scale = min(
                (page_w - 2 * margin) / img_w,
                (page_h - 2 * margin) / img_h
            )

            draw_w = img_w * scale
            draw_h = img_h * scale
            x = (page_w - draw_w) / 2
            y = (page_h - draw_h) / 2

            pdf.drawImage(
                ImageReader(image),
                x, y,
                width=draw_w,
                height=draw_h,
                preserveAspectRatio=True
            )
            pdf.showPage()

    pdf.save()
    output.seek(0)
    return output.getvalue()


def save_scanned_pdf(pdf_bytes, sold_to, filename):
    folder = get_certificate_folder(
        datetime.now().date(),
        sold_to or "Scanned"
    )
    path = folder / safe_filename(filename, "scanned_document")
    path.write_bytes(pdf_bytes)
    return folder, path


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "mcert_pdf": None,
    "certificate_data": None,
    "mcert_filename": "test_certificate.pdf",
    "mcert_copies": 1,
    "printer_connected": True,
    "available_scanners": [],
    "scan_pages": [],
    "scan_pdf": None,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# MAIN NAVIGATION
# ============================================================

menu = st.radio(
    "MENU",
    ["PRINT", "SCANNER"],
    horizontal=True,
    key="main_menu"
)
st.divider()


# ============================================================
# PRINT PAGE
# ============================================================

if menu == "PRINT":
    st.header("🖨️ PRINT")

    left, right = st.columns(2)

    with left:
        date_text = st.text_input(
            "Date (DD-MMM-YY)",
            value=datetime.now().strftime("%d-%b-%y"),
            key="date_input"
        )
        item_name = st.text_input("Item Name", key="item_input")

        show_model = st.checkbox("Model No.", key="show_model")
        show_serial = st.checkbox("Serial No.", key="show_serial")
        show_make = st.checkbox("Make", key="show_make")

    with right:
        model_no = st.text_input("Model No. Value") if show_model else ""
        serial_no = st.text_input("Serial No. Value") if show_serial else ""
        make = st.text_input("Make Value") if show_make else ""
        sold_to = st.text_input("Sold To / Company Name", key="sold_to")

    selected_columns = []

    if show_model:
        selected_columns.append(("model_no", "Model No."))
    if show_serial:
        selected_columns.append(("serial_no", "Serial No."))
    if show_make:
        selected_columns.append(("make", "Make"))

    copies = st.number_input(
        "Number of print copies",
        min_value=1,
        max_value=20,
        value=1
    )

    if st.button("📄 GENERATE CERTIFICATE", type="primary",
                 use_container_width=True):
        parsed_date = parse_date(date_text)

        required = {
            "Item Name": item_name,
            "Sold To": sold_to,
        }

        if show_model:
            required["Model No."] = model_no
        if show_serial:
            required["Serial No."] = serial_no
        if show_make:
            required["Make"] = make

        missing = [k for k, v in required.items() if not str(v).strip()]

        if not parsed_date:
            st.error("Invalid date. Use DD-MMM-YY, e.g. 11-Oct-26.")
        elif missing:
            st.error("Please fill in: " + ", ".join(missing))
        else:
            data = {
                "date": parsed_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns,
            }

            try:
                st.session_state["mcert_pdf"] = make_certificate_pdf(data)
                st.session_state["certificate_data"] = data
                st.session_state["mcert_copies"] = int(copies)
                st.session_state["mcert_filename"] = safe_filename(
                    item_name, "test_certificate"
                )
                st.success("Certificate generated.")
            except Exception as exc:
                st.error(f"Certificate generation failed: {exc}")

    pdf_bytes = st.session_state.get("mcert_pdf")
    data = st.session_state.get("certificate_data")

    if pdf_bytes and data:
        st.divider()
        st.subheader("💾 SAVE CERTIFICATE")

        filename = st.text_input(
            "File name",
            value=st.session_state["mcert_filename"],
            key="save_filename"
        )

        st.download_button(
            "⬇️ DOWNLOAD PDF",
            data=pdf_bytes,
            file_name=safe_filename(filename, "test_certificate"),
            mime="application/pdf",
            use_container_width=True
        )

        if st.button("💾 SAVE PDF + WORD", use_container_width=True):
            try:
                folder, pdf_path, word_path = save_certificate_files(
                    data, pdf_bytes, filename
                )
                st.success("PDF and Word files saved.")
                st.code(str(folder))
                st.write("PDF:", str(pdf_path))
                st.write("Word:", str(word_path))
            except Exception as exc:
                st.error(f"Save failed: {exc}")

        st.divider()
        st.subheader("🖨️ PRINT CERTIFICATE")

        if not st.session_state["printer_connected"]:
            st.warning("Printer is disconnected from this application's print workflow.")
        else:
            printers = get_printers()
            default = get_default_printer()

            if printers:
                index = printers.index(default) if default in printers else 0
                printer = st.selectbox("Select printer", printers, index=index)

                if st.button("🖨️ PRINT NOW", type="primary",
                             use_container_width=True):
                    try:
                        message = print_pdf(
                            pdf_bytes, printer,
                            st.session_state["mcert_copies"]
                        )
                        st.success("Print job submitted to CUPS.")
                        if message:
                            st.code(message)
                    except Exception as exc:
                        st.error(f"Printing failed: {exc}")
            else:
                st.warning("No CUPS printer found. Check `lpstat -p`.")


# ============================================================
# SCANNER PAGE
# ============================================================

else:
    st.header("📷 SCANNER")

    st.subheader("Printer Connection")

    if st.session_state["printer_connected"]:
        st.success("Printer workflow: ON")
        if st.button("🔌 DISCONNECT PRINTER"):
            st.session_state["printer_connected"] = False
            st.rerun()
    else:
        st.warning("Printer workflow: OFF")
        if st.button("🔌 RECONNECT PRINTER"):
            st.session_state["printer_connected"] = True
            st.rerun()

    st.caption(
        "This switch enables/disables printing inside this app. "
        "It does not physically disconnect the USB/network printer."
    )

    st.divider()
    st.subheader("Find Scanner")

    if st.button("🔄 FIND SCANNERS", use_container_width=True):
        try:
            st.session_state["available_scanners"] = find_scanners()
            if st.session_state["available_scanners"]:
                st.success(
                    f"{len(st.session_state['available_scanners'])} scanner(s) found."
                )
            else:
                st.warning("No scanner found. Check `scanimage -L`.")
        except Exception as exc:
            st.error(f"Scanner detection failed: {exc}")

    scanners = st.session_state["available_scanners"]

    if scanners:
        scanner_labels = [
            f"{item['name']} — {item['device']}"
            for item in scanners
        ]
        selected_label = st.selectbox("Select scanner", scanner_labels)
        selected_index = scanner_labels.index(selected_label)
        selected_device = scanners[selected_index]["device"]

        resolution = st.selectbox(
            "Resolution (DPI)", [100, 150, 200, 300, 600], index=2
        )
        scan_mode = st.selectbox("Scan mode", ["Color", "Gray", "Lineart"])

        col1, col2 = st.columns(2)

        with col1:
            if st.button("📷 START SCAN / NEW DOCUMENT",
                         type="primary", use_container_width=True):
                try:
                    with st.spinner("Scanning page..."):
                        page = scan_one_page(
                            selected_device, resolution, scan_mode
                        )
                    st.session_state["scan_pages"] = [page]
                    st.session_state["scan_pdf"] = make_scanned_pdf(
                        st.session_state["scan_pages"]
                    )
                    st.success("First page scanned.")
                except Exception as exc:
                    st.error(f"Scan failed: {exc}")

        with col2:
            if st.button("➕ SCAN NEXT PAGE", use_container_width=True):
                try:
                    with st.spinner("Scanning next page..."):
                        page = scan_one_page(
                            selected_device, resolution, scan_mode
                        )
                    st.session_state["scan_pages"].append(page)
                    st.session_state["scan_pdf"] = make_scanned_pdf(
                        st.session_state["scan_pages"]
                    )
                    st.success(
                        f"Page {len(st.session_state['scan_pages'])} added."
                    )
                except Exception as exc:
                    st.error(f"Scan failed: {exc}")

        if st.button("🗑️ CLEAR SCANNED PAGES"):
            st.session_state["scan_pages"] = []
            st.session_state["scan_pdf"] = None
            st.rerun()

    else:
        st.info("Click FIND SCANNERS first.")

    pages = st.session_state.get("scan_pages", [])
    scanned_pdf = st.session_state.get("scan_pdf")

    if pages and scanned_pdf:
        st.divider()
        st.subheader(f"📑 Combined PDF — {len(pages)} page(s)")

        for i, page_bytes in enumerate(pages):
            st.markdown(f"**Page {i + 1}**")
            st.image(page_bytes, use_container_width=True)

        scan_name = st.text_input(
            "Scanned PDF file name",
            value="scanned_document.pdf",
            key="scan_filename"
        )
        scan_customer = st.text_input(
            "Sold To / Customer Folder",
            key="scan_sold_to"
        )

        st.download_button(
            "⬇️ DOWNLOAD SCANNED PDF",
            data=scanned_pdf,
            file_name=safe_filename(scan_name, "scanned_document"),
            mime="application/pdf",
            use_container_width=True
        )

        if st.button("💾 SAVE SCANNED PDF", type="primary",
                     use_container_width=True):
            try:
                folder, path = save_scanned_pdf(
                    scanned_pdf, scan_customer, scan_name
                )
                st.success("Scanned PDF saved.")
                st.code(str(folder))
                st.write(str(path))
            except Exception as exc:
                st.error(f"Could not save scanned PDF: {exc}")


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption("M-CERTIFICATE • PRINT | SCANNER • Linux CUPS + SANE")
