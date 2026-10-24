
import io
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime

import streamlit as st
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth


# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="🖨️",
    layout="wide",
)

st.markdown("""
<style>
.stApp {
    background: #0b1220;
    color: #eef2ff;
}
.block-container {
    max-width: 1150px;
    padding-top: 1.5rem;
}
h1, h2, h3, label, p {
    color: #eef2ff !important;
}
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("A4 Test Certificate • PDF • Linux CUPS Printing • Scan")


# --------------------------------------------------
# LINUX COMMAND HELPERS
# --------------------------------------------------
def run_command(args, timeout=30):
    """Run a system command safely and return its result."""
    try:
        return subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError:
        return None
    except subprocess.TimeoutExpired:
        return None


def get_printers():
    """Read printers configured in the Linux CUPS environment."""
    if not shutil.which("lpstat"):
        return [], None, "Linux command 'lpstat' is not installed."

    result = run_command(["lpstat", "-p"])

    if result is None:
        return [], None, "Could not execute lpstat."

    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip()
        return [], None, detail or "CUPS printer listing failed."

    printers = []

    for line in result.stdout.splitlines():
        # Example: printer HP_LaserJet is idle. enabled since ...
        match = re.match(r"^printer\s+(\S+)", line)
        if match:
            printers.append(match.group(1))

    default_printer = None
    default_result = run_command(["lpstat", "-d"])

    if default_result and default_result.returncode == 0:
        match = re.search(
            r"system default destination:\s*(\S+)",
            default_result.stdout,
        )
        if match:
            default_printer = match.group(1)

    return sorted(set(printers)), default_printer, None


def get_scanners():
    """Find scanners visible to the Linux SANE subsystem."""
    if not shutil.which("scanimage"):
        return [], "Linux command 'scanimage' is not installed."

    result = run_command(["scanimage", "-L"], timeout=30)

    if result is None:
        return [], "Could not execute scanimage."

    output = (result.stdout or "").strip()
    errors = (result.stderr or "").strip()

    devices = []

    for line in output.splitlines():
        match = re.search(r"device\s+`([^']+)'", line)
        if match:
            devices.append((match.group(1), line.strip()))

    if devices:
        return devices, None

    return [], errors or "No scanner devices found by SANE."


# --------------------------------------------------
# PDF HELPERS
# --------------------------------------------------
def wrap_words(text, font, size, max_width):
    words = str(text).split()
    lines = []
    line = ""

    for word in words:
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


def fit_text(c, text, x, y, max_width,
             font="Helvetica", size=10, min_size=7):
    current_size = size

    while (
        current_size > min_size
        and stringWidth(text, font, current_size) > max_width
    ):
        current_size -= 0.25

    c.setFont(font, current_size)
    c.drawString(x, y, text)


def make_pdf(data):
    """Generate the A4 certificate, retaining the original layout."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4

    c.setFillColorRGB(0.08, 0.08, 0.08)

    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    fit_text(
        c, date_text, 428, H - 204, 100,
        size=12, min_size=10,
    )

    title = "TEST CERTIFICATE"
    title_y = H - 142

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)

    title_width = stringWidth(title, "Helvetica-Bold", 16)
    c.setLineWidth(0.8)
    c.line(
        W / 2 - title_width / 2,
        title_y - 3,
        W / 2 + title_width / 2,
        title_y - 3,
    )

    whom = "TO WHOM IT MAY CONCERN"
    whom_y = H - 262

    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, whom_y, whom)

    whom_width = stringWidth(whom, "Helvetica", 12)
    c.setLineWidth(0.55)
    c.line(
        W / 2 - whom_width / 2,
        whom_y - 2.5,
        W / 2 + whom_width / 2,
        whom_y - 2.5,
    )

    item_name = data["item_name"]
    selected = data["selected_columns"]
    labels = [label for _, label in selected]

    if not labels:
        intro = (
            f"This test certificate is for {item_name} "
            "as per below table"
        )
    elif len(labels) == 1:
        intro = (
            f"This test certificate is for {item_name} "
            f"having {labels[0]} as per below table"
        )
    elif len(labels) == 2:
        intro = (
            f"This test certificate is for {item_name} "
            f"having {labels[0]} and {labels[1]} "
            "as per below table"
        )
    else:
        intro = (
            f"This test certificate is for {item_name} "
            f"having {', '.join(labels[:-1])} and "
            f"{labels[-1]} as per below table"
        )

    intro_y = H - 310
    intro_lines = wrap_words(intro, "Helvetica", 12, 470)

    c.setFont("Helvetica", 12)
    for i, line in enumerate(intro_lines[:2]):
        c.drawString(72, intro_y - i * 17, line)

    # Dynamic table
    x_left, x_right = 72, 550
    table_top = H - 359
    header_h, row_h = 31, 31
    table_bottom = table_top - header_h - row_h

    columns = [
        ("item_name", "ITEM", data["item_name"])
    ]

    for key, label in selected:
        columns.append((key, label, data[key]))

    col_width = (x_right - x_left) / len(columns)
    x_positions = [
        x_left + i * col_width
        for i in range(len(columns) + 1)
    ]

    c.setLineWidth(0.55)

    for x in x_positions:
        c.line(x, table_top, x, table_bottom)

    c.line(x_left, table_top, x_right, table_top)
    c.line(
        x_left, table_top - header_h,
        x_right, table_top - header_h,
    )
    c.line(x_left, table_bottom, x_right, table_bottom)

    c.setFont("Helvetica-Bold", 11)

    for i, (_, label, _) in enumerate(columns):
        cx = (x_positions[i] + x_positions[i + 1]) / 2
        c.drawCentredString(cx, table_top - 20, label)

    for i, (_, _, value) in enumerate(columns):
        left = x_positions[i]
        right = x_positions[i + 1]
        txt = str(value)

        size = 10.5
        while (
            size > 7
            and stringWidth(txt, "Helvetica", size)
            > (right - left - 10)
        ):
            size -= 0.25

        c.setFont("Helvetica", size)
        c.drawCentredString(
            (left + right) / 2,
            table_bottom + 9.5,
            txt,
        )

    body_y = H - 452
    c.setFont("Helvetica", 11.8)

    c.drawString(
        72, body_y,
        "Specified component is tested for all its functionalities and is performing",
    )
    c.drawString(
        72, body_y - 18,
        "satisfactorily and found OK.",
    )

    sold = f'The above said item is sold to {data["sold_to"]}.'
    sold_lines = wrap_words(sold, "Helvetica-Oblique", 11.8, 470)

    c.setFont("Helvetica-Oblique", 11.8)
    sold_y = H - 528

    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, sold_y - i * 17, line)

    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.showPage()
    c.save()

    buffer.seek(0)
    return buffer.getvalue()


def parse_date(value):
    try:
        return datetime.strptime(
            value.strip(), "%d-%b-%y"
        ).date()
    except ValueError:
        return None


def safe_filename(item_name):
    name = re.sub(
        r"[^A-Za-z0-9]+", "_", item_name.strip()
    ).strip("_")

    return f"test_certificate_{name or 'item'}.pdf"


def print_pdf(pdf_bytes, printer, copies):
    """Submit a PDF to a configured Linux CUPS printer."""
    if not shutil.which("lp"):
        return False, "Linux command 'lp' is not installed."

    pdf_path = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".pdf",
            prefix="mcertificate_",
            delete=False,
        ) as temp_file:
            temp_file.write(pdf_bytes)
            pdf_path = temp_file.name

        result = subprocess.run(
            [
                "lp",
                "-d", printer,
                "-n", str(int(copies)),
                "-o", "media=A4",
                pdf_path,
            ],
            capture_output=True,
            text=True,
            timeout=90,
            check=False,
        )

        if result.returncode == 0:
            return True, (
                result.stdout.strip()
                or "Print job submitted to CUPS."
            )

        detail = result.stderr.strip() or result.stdout.strip()
        return False, detail or "CUPS could not submit the print job."

    except subprocess.TimeoutExpired:
        return False, "Printing timed out. Check the CUPS queue."

    except Exception as exc:
        return False, f"Printing failed: {exc}"

    finally:
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except OSError:
                pass


def scan_to_pdf(device, resolution=200):
    """Scan one page using SANE and convert it to a PDF."""
    if not shutil.which("scanimage"):
        return None, "Linux command 'scanimage' is not installed."

    image_path = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".pnm",
            prefix="mcertificate_scan_",
            delete=False,
        ) as temp_file:
            image_path = temp_file.name

        result = subprocess.run(
            [
                "scanimage",
                "--device-name", device,
                "--format=pnm",
                "--resolution", str(int(resolution)),
                "--output-file", image_path,
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )

        if result.returncode != 0:
            detail = result.stderr.strip() or result.stdout.strip()
            return None, detail or "Scanner failed to scan."

        if not os.path.exists(image_path) or os.path.getsize(image_path) == 0:
            return None, "Scanner returned an empty file."

        with Image.open(image_path) as image:
            image = image.convert("RGB")
            output = io.BytesIO()
            image.save(output, format="PDF", resolution=resolution)
            return output.getvalue(), None

    except subprocess.TimeoutExpired:
        return None, "Scanning timed out."

    except Exception as exc:
        return None, f"Scanning failed: {exc}"

    finally:
        if image_path and os.path.exists(image_path):
            try:
                os.remove(image_path)
            except OSError:
                pass


# --------------------------------------------------
# CERTIFICATE INPUT FORM
# --------------------------------------------------
st.subheader("Certificate Details")

left, right = st.columns(2)

with left:
    date_text = st.text_input(
        "Date",
        placeholder="DD-MMM-YY (e.g. 23-Jun-26)",
    )

    item_name = st.text_input(
        "Item Name",
        placeholder="Enter item name",
    )

    st.markdown("**Select fields to show in table**")

    show_model = st.checkbox("Model No.")
    show_serial = st.checkbox("Serial No.")
    show_make = st.checkbox("Make")

with right:
    model_no = (
        st.text_input("Model No. Value")
        if show_model else ""
    )

    serial_no = (
        st.text_input("Serial No. Value")
        if show_serial else ""
    )

    make = (
        st.text_input("Make Value")
        if show_make else ""
    )

    sold_to = st.text_input(
        "Sold To",
        placeholder="Enter customer name",
    )

copies = st.number_input(
    "Number of Print Copies",
    min_value=1,
    max_value=20,
    value=1,
    step=1,
)

selected_columns = []

if show_model:
    selected_columns.append(("model_no", "Model No."))
if show_serial:
    selected_columns.append(("serial_no", "Serial No."))
if show_make:
    selected_columns.append(("make", "Make"))


# --------------------------------------------------
# GENERATE CERTIFICATE
# --------------------------------------------------
if st.button(
    "📄 GENERATE CERTIFICATE PDF",
    type="primary",
    use_container_width=True,
):
    parsed_date = parse_date(date_text)

    required = {
        "Date": date_text,
        "Item Name": item_name,
        "Sold To": sold_to,
    }

    if show_model:
        required["Model No."] = model_no
    if show_serial:
        required["Serial No."] = serial_no
    if show_make:
        required["Make"] = make

    missing = [
        name for name, value in required.items()
        if not value.strip()
    ]

    if missing:
        st.error("Please fill in: " + ", ".join(missing))

    elif parsed_date is None:
        st.error("Date must be DD-MMM-YY, e.g. 23-Jun-26.")

    else:
        try:
            pdf_bytes = make_pdf({
                "date": parsed_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns,
            })

            st.session_state["mcert_pdf"] = pdf_bytes
            st.session_state["mcert_filename"] = safe_filename(item_name)
            st.session_state["mcert_copies"] = int(copies)

            st.success("Certificate PDF generated successfully!")

        except Exception as exc:
            st.error(f"PDF generation failed: {exc}")


# --------------------------------------------------
# DOWNLOAD AND PRINT
# --------------------------------------------------
if "mcert_pdf" in st.session_state:
    st.divider()
    st.subheader("PDF Ready — Download or Print")

    st.download_button(
        "⬇️ DOWNLOAD PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get(
            "mcert_filename", "test_certificate.pdf"
        ),
        mime="application/pdf",
        use_container_width=True,
    )

    st.write("### Linux CUPS Printer Settings")

    printers, default_printer, printer_error = get_printers()

    if printers:
        default_index = (
            printers.index(default_printer)
            if default_printer in printers else 0
        )

        printer = st.selectbox(
            "Select printer",
            printers,
            index=default_index,
            key="mcert_printer",
        )

        st.caption(
            f"Copies: {st.session_state['mcert_copies']} | Paper: A4"
        )

        if st.button(
            "🖨️ PRINT NOW",
            type="primary",
            use_container_width=True,
        ):
            with st.spinner("Submitting print job to Linux CUPS..."):
                ok, message = print_pdf(
                    st.session_state["mcert_pdf"],
                    printer,
                    st.session_state["mcert_copies"],
                )

            if ok:
                st.success(message)
            else:
                st.error(message)

    else:
        st.warning(
            "No printer is configured in this Linux environment. "
            "PDF generation and download still work."
        )
        if printer_error:
            st.caption(printer_error)


# --------------------------------------------------
# SCANNER
# --------------------------------------------------
st.divider()
st.subheader("Scan Document")

st.caption(
    "The scanner must be physically connected to, or network-accessible "
    "from, the Linux machine running this app."
)

if st.button("🔎 CHECK SCANNERS", use_container_width=True):
    devices, scan_error = get_scanners()
    st.session_state["mcert_scanners"] = devices
    st.session_state["mcert_scan_error"] = scan_error

if "mcert_scanners" in st.session_state:
    scanners = st.session_state["mcert_scanners"]

    if scanners:
        device_names = [device for device, _ in scanners]

        selected_device = st.selectbox(
            "Select scanner",
            device_names,
            key="mcert_scanner",
        )

        resolution = st.selectbox(
            "Scan resolution",
            [100, 150, 200, 300],
            index=2,
            format_func=lambda value: f"{value} DPI",
        )

        if st.button(
            "📠 SCAN TO PDF",
            type="primary",
            use_container_width=True,
        ):
            with st.spinner("Scanning one page..."):
                scan_pdf, scan_error = scan_to_pdf(
                    selected_device, resolution
                )

            if scan_pdf:
                st.session_state["mcert_scan_pdf"] = scan_pdf
                st.success("Scan completed successfully.")
            else:
                st.error(scan_error or "Scanning failed.")

    else:
        st.warning("No scanner was detected by Linux SANE.")
        st.caption(
            st.session_state.get(
                "mcert_scan_error",
                "Check scanner connection and SANE driver support.",
            )
        )

if "mcert_scan_pdf" in st.session_state:
    st.download_button(
        "⬇️ DOWNLOAD SCANNED PDF",
        data=st.session_state["mcert_scan_pdf"],
        file_name="scanned_document.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
