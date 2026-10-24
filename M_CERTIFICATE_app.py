
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime

import streamlit as st
import win32print

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth


# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
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
    max-width: 1150px;
    padding-top: 1.5rem;
}
h1, h2, h3, label, p {
    color: #eef2ff !important;
}
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("A4 Test Certificate • Generate PDF • Direct HP Printing")


# --------------------------------------------------
# SUMATRAPDF LOCATION
# --------------------------------------------------
SUMATRA_PATHS = [
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
    os.path.expandvars(
        r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"
    ),
]

sumatra = next(
    (p for p in SUMATRA_PATHS if os.path.isfile(p)),
    None
)


# --------------------------------------------------
# WINDOWS PRINTER LIST
# --------------------------------------------------
try:
    printers = sorted({
        p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL
            | win32print.PRINTER_ENUM_CONNECTIONS
        )
    })
except Exception as e:
    st.error(f"Cannot read Windows printers: {e}")
    st.stop()

if not printers:
    st.error(
        "No Windows printers found. "
        "Please install your HP printer first."
    )
    st.stop()

try:
    default_printer = win32print.GetDefaultPrinter()
except Exception:
    default_printer = printers[0]

default_index = (
    printers.index(default_printer)
    if default_printer in printers
    else 0
)


# --------------------------------------------------
# PDF HELPER FUNCTIONS
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
    """Generate an A4 test certificate PDF."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4

    c.setFillColorRGB(0.08, 0.08, 0.08)

    # Date
    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    fit_text(
        c, date_text, 428, H - 204, 100,
        size=12, min_size=10
    )

    # Main title
    title = "TEST CERTIFICATE"
    title_y = H - 142

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)

    title_width = stringWidth(
        title, "Helvetica-Bold", 16
    )
    c.setLineWidth(0.8)
    c.line(
        W / 2 - title_width / 2,
        title_y - 3,
        W / 2 + title_width / 2,
        title_y - 3
    )

    # Subtitle
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
        whom_y - 2.5
    )

    # Intro text based on selected fields
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
    intro_lines = wrap_words(
        intro, "Helvetica", 12, 470
    )

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
        x_right, table_top - header_h
    )
    c.line(x_left, table_bottom, x_right, table_bottom)

    # Table headers
    c.setFont("Helvetica-Bold", 11)

    for i, (_, label, _) in enumerate(columns):
        cx = (x_positions[i] + x_positions[i + 1]) / 2
        c.drawCentredString(
            cx, table_top - 20, label
        )

    # Table values
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
            txt
        )

    # Certificate body
    body_y = H - 452
    c.setFont("Helvetica", 11.8)

    c.drawString(
        72, body_y,
        "Specified component is tested for all its functionalities and is performing"
    )
    c.drawString(
        72, body_y - 18,
        "satisfactorily and found OK."
    )

    sold = f'The above said item is sold to {data["sold_to"]}.'
    sold_lines = wrap_words(
        sold, "Helvetica-Oblique", 11.8, 470
    )

    c.setFont("Helvetica-Oblique", 11.8)
    sold_y = H - 528

    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, sold_y - i * 17, line)

    # Authorization area
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


# --------------------------------------------------
# CERTIFICATE INPUT FORM
# --------------------------------------------------
st.subheader("Certificate Details")

left, right = st.columns(2)

with left:
    date_text = st.text_input(
        "Date",
        placeholder="DD-MMM-YY (e.g. 23-Jun-26)"
    )

    item_name = st.text_input(
        "Item Name",
        placeholder="Enter item name"
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
        placeholder="Enter customer name"
    )

copies = st.number_input(
    "Number of Print Copies",
    min_value=1,
    max_value=20,
    value=1,
    step=1
)

selected_columns = []

if show_model:
    selected_columns.append(("model_no", "Model No."))

if show_serial:
    selected_columns.append(("serial_no", "Serial No."))

if show_make:
    selected_columns.append(("make", "Make"))


# --------------------------------------------------
# GENERATE PDF
# --------------------------------------------------
if st.button(
    "📄 GENERATE CERTIFICATE PDF",
    type="primary",
    use_container_width=True
):
    parsed_date = parse_date(date_text)

    required = {
        "Date": date_text,
        "Item Name": item_name,
        "Sold To": sold_to
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

    if parsed_date is None and "Date" not in missing:
        st.error(
            "Date must be DD-MMM-YY, e.g. 23-Jun-26."
        )

    elif missing:
        st.error(
            "Please fill in: " + ", ".join(missing)
        )

    else:
        try:
            pdf_bytes = make_pdf({
                "date": parsed_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns
            })

            st.session_state["mcert_pdf"] = pdf_bytes
            st.session_state["mcert_filename"] = (
                safe_filename(item_name)
            )
            st.session_state["mcert_copies"] = int(copies)

            st.success(
                "Certificate PDF generated successfully!"
            )

        except Exception as e:
            st.error(f"PDF generation failed: {e}")


# --------------------------------------------------
# DOWNLOAD PDF + DIRECT PRINT BUTTONS
# --------------------------------------------------
if "mcert_pdf" in st.session_state:

    st.divider()
    st.subheader("PDF Ready — Download or Print")

    st.download_button(
        "⬇️ DOWNLOAD PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get(
            "mcert_filename",
            "test_certificate.pdf"
        ),
        mime="application/pdf",
        type="secondary",
        use_container_width=True
    )

    st.write("### Printer Settings")

    printer = st.selectbox(
        "Select HP Printer",
        printers,
        index=(
            printers.index(default_printer)
            if default_printer in printers else 0
        ),
        key="mcert_printer"
    )

    st.caption(
        f"Copies to print: {st.session_state['mcert_copies']} | "
        "Paper: A4"
    )

    # PRINT BUTTON BELOW PDF DOWNLOAD
    if st.button(
        "🖨️ PRINT NOW",
        type="primary",
        use_container_width=True
    ):

        if not sumatra:
            st.error(
                "SumatraPDF not found. Install it from: "
                "https://www.sumatrapdfreader.org/download-free-pdf-viewer"
            )

        else:
            pdf_path = None

            try:
                # Save the generated certificate as a temporary PDF
                fd, pdf_path = tempfile.mkstemp(
                    suffix=".pdf",
                    prefix="mcertificate_"
                )

                with os.fdopen(fd, "wb") as f:
                    f.write(st.session_state["mcert_pdf"])

                # Send PDF directly to the selected Windows printer
                result = subprocess.run(
                    [
                        sumatra,
                        "-print-to",
                        printer,
                        "-print-settings",
                        f"copies={st.session_state['mcert_copies']},paper=A4",
                        "-silent",
                        pdf_path
                    ],
                    capture_output=True,
                    text=True,
                    timeout=90
                )

                if result.returncode == 0:
                    st.success(
                        "Print command completed. "
                        "Please check the HP printer / print queue."
                    )
                    st.caption(
                        "The temporary PDF is retained to avoid "
                        "deleting it before the print job reads it."
                    )

                else:
                    st.error(
                        f"Printing failed. Exit code: "
                        f"{result.returncode}"
                    )

                    if result.stderr:
                        st.code(result.stderr)

            except subprocess.TimeoutExpired:
                st.error(
                    "Printing timed out. Check the Windows "
                    "printer queue and printer connection."
                )

            except Exception as e:
                st.error(f"Printing failed: {e}")




give currected full code
