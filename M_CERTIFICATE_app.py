import io
import os
import re
import subprocess
import tempfile
from datetime import datetime

import streamlit as st
import win32print
import win32com.client

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.utils import ImageReader

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(page_title="M-CERTIFICATE", page_icon="🖨️", layout="wide")

st.markdown("""
<style>
.stApp { background:#0b1220; color:#eef2ff; }
.block-container { max-width:1250px; padding-top:1.2rem; }
h1,h2,h3,h4,label,p { color:#eef2ff !important; }
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("A4 Test Certificate • Scanner • Multiple Scan Pages • Direct HP Printing")

# ============================================================
# SESSION STATE - APP STARTS BLANK
# ============================================================
for key, value in {
    "cert_pdf": None,
    "cert_filename": "test_certificate.pdf",
    "cert_copies": 1,
    "scan_pages": [],
    "scanned_pdf": None,
    "scanned_filename": "scanned_document.pdf",
}.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ============================================================
# SUMATRA PDF
# ============================================================
SUMATRA_PATHS = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
]
sumatra = next((p for p in SUMATRA_PATHS if os.path.isfile(p)), None)

# ============================================================
# WINDOWS PRINTERS
# ============================================================
try:
    printers = sorted({
        p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    })
except Exception:
    printers = []

try:
    default_printer = win32print.GetDefaultPrinter()
except Exception:
    default_printer = printers[0] if printers else ""

printer_index = printers.index(default_printer) if default_printer in printers else 0

# ============================================================
# WIA SCANNER
# ============================================================
WIA_SCANNER_TYPE = 1
WIA_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"


def get_scanners():
    """Return Windows WIA scanners as [(name, device_id), ...]."""
    result = []
    try:
        manager = win32com.client.Dispatch("WIA.DeviceManager")
        for info in manager.DeviceInfos:
            try:
                if int(info.Type) == WIA_SCANNER_TYPE:
                    result.append((
                        str(info.Properties("Name").Value),
                        str(info.DeviceID),
                    ))
            except Exception:
                pass
    except Exception as exc:
        return [], str(exc)
    return result, ""


def scan_one_page(device_id):
    """Scan one page and return JPEG bytes."""
    manager = win32com.client.Dispatch("WIA.DeviceManager")
    device = None

    for info in manager.DeviceInfos:
        try:
            if str(info.DeviceID) == str(device_id):
                device = info.Connect()
                break
        except Exception:
            pass

    if device is None:
        raise RuntimeError("Selected scanner could not be connected.")
    if device.Items.Count < 1:
        raise RuntimeError("Scanner has no scan item available.")

    item = device.Items(1)
    try:
        image = item.Transfer(WIA_JPEG)
    except Exception:
        image = item.Transfer()

    try:
        data = bytes(image.FileData.BinaryData)
    except Exception:
        fd, path = tempfile.mkstemp(suffix=".jpg")
        os.close(fd)
        try:
            image.SaveFile(path)
            with open(path, "rb") as f:
                data = f.read()
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    if not data:
        raise RuntimeError("Scanner returned an empty image.")
    return data

# ============================================================
# HELPERS
# ============================================================
def wrap_words(text, font, size, max_width):
    lines, line = [], ""
    for word in str(text).split():
        test = word if not line else line + " " + word
        if stringWidth(test, font, size) <= max_width:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


def fit_text(c, text, x, y, max_width, font="Helvetica", size=10, min_size=7):
    current = size
    while current > min_size and stringWidth(text, font, current) > max_width:
        current -= 0.25
    c.setFont(font, current)
    c.drawString(x, y, text)


def safe_filename(name, prefix="test_certificate"):
    clean = re.sub(r"[^A-Za-z0-9]+", "_", str(name).strip()).strip("_")
    return f"{prefix}_{clean or 'item'}.pdf"


def parse_date(text):
    try:
        return datetime.strptime(text.strip(), "%d-%b-%y").date()
    except ValueError:
        return None


def intro_text(item_name, selected):
    labels = [label for _, label in selected]
    if not labels:
        return f"This test certificate is for {item_name} as per below table"
    if len(labels) == 1:
        return f"This test certificate is for {item_name} having {labels[0]} as per below table"
    if len(labels) == 2:
        return f"This test certificate is for {item_name} having {labels[0]} and {labels[1]} as per below table"
    return f"This test certificate is for {item_name} having {', '.join(labels[:-1])} and {labels[-1]} as per below table"

# ============================================================
# CERTIFICATE PDF
# ============================================================
def make_certificate_pdf(data):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    c.setFillColorRGB(0.08, 0.08, 0.08)

    # Date
    fit_text(c, "Date: " + data["date"].strftime("%d-%b-%y"), 428, H - 204, 100, size=12, min_size=10)

    # Title
    title = "TEST CERTIFICATE"
    title_y = H - 142
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)
    tw = stringWidth(title, "Helvetica-Bold", 16)
    c.setLineWidth(0.8)
    c.line(W / 2 - tw / 2, title_y - 3, W / 2 + tw / 2, title_y - 3)

    # To whom
    whom = "TO WHOM IT MAY CONCERN"
    whom_y = H - 262
    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, whom_y, whom)
    ww = stringWidth(whom, "Helvetica", 12)
    c.setLineWidth(0.55)
    c.line(W / 2 - ww / 2, whom_y - 2.5, W / 2 + ww / 2, whom_y - 2.5)

    # Intro
    lines = wrap_words(intro_text(data["item_name"], data["selected_columns"]), "Helvetica", 12, 470)
    c.setFont("Helvetica", 12)
    for i, line in enumerate(lines[:2]):
        c.drawString(72, H - 310 - i * 17, line)

    # Dynamic table
    x_left, x_right = 72, 550
    table_top = H - 359
    header_h, row_h = 31, 31
    table_bottom = table_top - header_h - row_h

    columns = [("item_name", "ITEM", data["item_name"])]
    columns += [(key, label, data[key]) for key, label in data["selected_columns"]]

    col_w = (x_right - x_left) / len(columns)
    xs = [x_left + i * col_w for i in range(len(columns) + 1)]

    c.setLineWidth(0.55)
    for x in xs:
        c.line(x, table_top, x, table_bottom)
    c.line(x_left, table_top, x_right, table_top)
    c.line(x_left, table_top - header_h, x_right, table_top - header_h)
    c.line(x_left, table_bottom, x_right, table_bottom)

    c.setFont("Helvetica-Bold", 11)
    for i, (_, label, _) in enumerate(columns):
        c.drawCentredString((xs[i] + xs[i + 1]) / 2, table_top - 20, label)

    for i, (_, _, value) in enumerate(columns):
        left, right = xs[i], xs[i + 1]
        text = str(value)
        size = 10.5
        while size > 7 and stringWidth(text, "Helvetica", size) > right - left - 10:
            size -= 0.25
        c.setFont("Helvetica", size)
        c.drawCentredString((left + right) / 2, table_bottom + 9.5, text)

    # Body
    body_y = H - 452
    c.setFont("Helvetica", 11.8)
    c.drawString(72, body_y, "Specified component is tested for all its functionalities and is performing")
    c.drawString(72, body_y - 18, "satisfactorily and found OK.")

    sold_lines = wrap_words(f'The above said item is sold to {data["sold_to"]}.', "Helvetica-Oblique", 11.8, 470)
    c.setFont("Helvetica-Oblique", 11.8)
    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, H - 528 - i * 17, line)

    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# SCANNED PAGES -> ONE PDF
# ============================================================
def make_scanned_pdf(pages):
    if not pages:
        return None
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    margin = 18

    for image_bytes in pages:
        image = ImageReader(io.BytesIO(image_bytes))
        iw, ih = image.getSize()
        scale = min((W - 2 * margin) / iw, (H - 2 * margin) / ih)
        dw, dh = iw * scale, ih * scale
        x, y = (W - dw) / 2, (H - dh) / 2
        c.drawImage(image, x, y, width=dw, height=dh, preserveAspectRatio=True, mask="auto")
        c.showPage()

    c.save()
    buf.seek(0)
    return buf.getvalue()


def print_pdf(pdf_bytes, printer, copies=1, prefix="mcertificate"):
    if not sumatra:
        raise RuntimeError("SumatraPDF was not found on this Windows PC.")

    fd, path = tempfile.mkstemp(suffix=".pdf", prefix=prefix + "_")
    os.close(fd)
    with open(path, "wb") as f:
        f.write(pdf_bytes)

    try:
        result = subprocess.run(
            [
                sumatra,
                "-print-to", printer,
                "-print-settings", f"copies={copies},paper=A4",
                "-silent", path,
            ],
            capture_output=True,
            text=True,
            timeout=90,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or f"Print failed. Exit code: {result.returncode}")
    finally:
        # Sumatra normally queues the job before returning. Keep a small delay
        # unnecessary here; Windows has already accepted the print job.
        try:
            os.remove(path)
        except OSError:
            pass

# ============================================================
# TWO SIDE-BY-SIDE PANELS
# ============================================================
left, right = st.columns([1.55, 1], gap="large")

# ============================================================
# LEFT: CERTIFICATE
# ============================================================
with left:
    st.subheader("Certificate Details")
    a, b = st.columns(2)

    with a:
        date_text = st.text_input("Date", placeholder="DD-MMM-YY (e.g. 23-Jun-26)")
        item_name = st.text_input("Item Name", placeholder="Enter item name")
        st.markdown("**Select fields to show in table**")
        show_model = st.checkbox("Model No.", key="show_model")
        show_serial = st.checkbox("Serial No.", key="show_serial")
        show_make = st.checkbox("Make", key="show_make")

    with b:
        model_no = st.text_input("Model No. Value", key="model_no") if show_model else ""
        serial_no = st.text_input("Serial No. Value", key="serial_no") if show_serial else ""
        make = st.text_input("Make Value", key="make") if show_make else ""
        sold_to = st.text_input("Sold To", placeholder="Enter customer name")

    copies = st.number_input("Number of Print Copies", min_value=1, max_value=20, value=1, step=1)

    selected_columns = []
    if show_model:
        selected_columns.append(("model_no", "MODEL NO."))
    if show_serial:
        selected_columns.append(("serial_no", "SERIAL NO."))
    if show_make:
        selected_columns.append(("make", "MAKE"))

    if st.button("📄 GENERATE CERTIFICATE PDF", type="primary", use_container_width=True):
        parsed_date = parse_date(date_text)
        required = {"Date": date_text, "Item Name": item_name, "Sold To": sold_to}
        if show_model:
            required["Model No."] = model_no
        if show_serial:
            required["Serial No."] = serial_no
        if show_make:
            required["Make"] = make

        missing = [name for name, value in required.items() if not str(value).strip()]
        if parsed_date is None and "Date" not in missing:
            st.error("Date must be DD-MMM-YY, e.g. 23-Jun-26.")
        elif missing:
            st.error("Please fill in: " + ", ".join(missing))
        else:
            try:
                st.session_state.cert_pdf = make_certificate_pdf({
                    "date": parsed_date,
                    "item_name": item_name.strip(),
                    "model_no": model_no.strip(),
                    "serial_no": serial_no.strip(),
                    "make": make.strip(),
                    "sold_to": sold_to.strip(),
                    "selected_columns": selected_columns,
                })
                st.session_state.cert_filename = safe_filename(item_name)
                st.session_state.cert_copies = int(copies)
                st.success("Certificate PDF generated.")
            except Exception as exc:
                st.error(f"PDF generation failed: {exc}")

    if st.session_state.cert_pdf:
        st.divider()
        st.subheader("Certificate PDF")
        st.download_button(
            "⬇️ DOWNLOAD CERTIFICATE PDF",
            data=st.session_state.cert_pdf,
            file_name=st.session_state.cert_filename,
            mime="application/pdf",
            use_container_width=True,
        )

        if printers:
            printer = st.selectbox("Select Printer", printers, index=printer_index, key="certificate_printer")
            st.caption(f"Copies: {st.session_state.cert_copies} | Paper: A4")
            if st.button("🖨️ PRINT CERTIFICATE NOW", type="primary", use_container_width=True):
                try:
                    print_pdf(st.session_state.cert_pdf, printer, st.session_state.cert_copies, "certificate")
                    st.success(f"Certificate sent to {printer}.")
                except Exception as exc:
                    st.error(f"Printing failed: {exc}")
        else:
            st.warning("No Windows printer found.")

# ============================================================
# RIGHT: SCANNER
# ============================================================
with right:
    st.subheader("📠 Scanner")
    scanners, scanner_error = get_scanners()

    if scanner_error:
        st.error("Scanner detection failed.")
        st.caption(scanner_error)
    elif not scanners:
        st.warning("No Windows WIA scanner detected.")
        st.caption("Install the HP scanner/WIA driver and make sure the MFP is available on Wi-Fi/Windows.")
    else:
        scanner_names = [name for name, _ in scanners]
        scanner_name = st.selectbox("Select Scanner", scanner_names, key="scanner_name")
        scanner_id = dict(scanners)[scanner_name]

        st.markdown("**Multiple scanning**")
        st.caption("SCAN starts a new set. ➕ adds the next page to the same PDF.")

        s1, s2 = st.columns(2)
        with s1:
            if st.button("📠 SCAN", type="primary", use_container_width=True):
                try:
                    st.session_state.scan_pages = [scan_one_page(scanner_id)]
                    st.session_state.scanned_pdf = None
                    st.success("Page 1 scanned.")
                except Exception as exc:
                    st.error(f"Scan failed: {exc}")

        with s2:
            if st.button("➕ SCAN NEXT", use_container_width=True):
                if not st.session_state.scan_pages:
                    st.warning("First click SCAN.")
                else:
                    try:
                        st.session_state.scan_pages.append(scan_one_page(scanner_id))
                        st.session_state.scanned_pdf = None
                        st.success(f"Page {len(st.session_state.scan_pages)} scanned.")
                    except Exception as exc:
                        st.error(f"Scan failed: {exc}")

        count = len(st.session_state.scan_pages)
        if count:
            st.info(f"📄 {count} page(s) scanned. All will become ONE PDF.")

            # Show first three previews only to keep the page fast.
            previews = st.columns(min(count, 3))
            for i, page in enumerate(st.session_state.scan_pages[:3]):
                with previews[i]:
                    st.image(page, caption=f"Page {i + 1}", use_container_width=True)
            if count > 3:
                st.caption(f"+ {count - 3} more scanned page(s)")

            scan_name = st.text_input(
                "PDF File Name",
                placeholder="Enter your own name, e.g. Amneal_Document",
                key="scan_name",
            )
            scan_filename = safe_filename(scan_name, "")
            if not scan_name.strip():
                scan_filename = "scanned_document.pdf"
            st.caption(f"File: {scan_filename}")

            x, y = st.columns(2)
            with x:
                if st.button("📑 CREATE SCANNED PDF", type="primary", use_container_width=True):
                    try:
                        st.session_state.scanned_pdf = make_scanned_pdf(st.session_state.scan_pages)
                        st.session_state.scanned_filename = scan_filename
                        st.success("All scanned pages combined into one PDF.")
                    except Exception as exc:
                        st.error(f"PDF creation failed: {exc}")

            with y:
                if st.button("🗑️ CLEAR SCANS", use_container_width=True):
                    st.session_state.scan_pages = []
                    st.session_state.scanned_pdf = None
                    st.rerun()

        if st.session_state.scanned_pdf:
            st.divider()
            st.subheader("Scanned PDF Ready")
            st.download_button(
                "⬇️ DOWNLOAD SCANNED PDF",
                data=st.session_state.scanned_pdf,
                file_name=st.session_state.scanned_filename,
                mime="application/pdf",
                use_container_width=True,
            )

            if printers:
                scan_printer = st.selectbox("Printer for Scanned PDF", printers, index=printer_index, key="scan_printer")
                if st.button("🖨️ PRINT SCANNED PDF", use_container_width=True):
                    try:
                        print_pdf(st.session_state.scanned_pdf, scan_printer, 1, "scanned")
                        st.success(f"Scanned PDF sent to {scan_printer}.")
                    except Exception as exc:
                        st.error(f"Printing failed: {exc}")

# ============================================================
# NOTES
# ============================================================
st.divider()
with st.expander("ℹ️ Scanner / Printer Setup"):
    st.write("Scanner: this app uses Windows WIA through pywin32. Your HP LaserJet Pro MFP must be installed in Windows with its scanner driver and be reachable on Wi-Fi.")
    st.write("Scanning: SCAN starts a new document. ➕ SCAN NEXT adds pages. CREATE SCANNED PDF makes one PDF containing all scanned pages in order.")
    st.write("Printer: direct printing uses SumatraPDF and the Windows printer queue. Select your HP LaserJet Pro MFP from the printer list.")
