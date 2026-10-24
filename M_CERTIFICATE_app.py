import io
import re
from datetime import date, datetime

import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="M-CERTIFICATE", page_icon="📄", layout="wide")

# Keep the UI light so the app loads/reruns quickly.
st.markdown("""
<style>
.stApp { background: #0b1220; color: #eef2ff; }
.block-container { max-width: 1150px; padding-top: 1.5rem; }
h1, h2, h3, label, p { color: #eef2ff !important; }
.small-note { color: #a9b7ce; font-size: 0.92rem; }
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("A4 Test Certificate · Select only the fields required for this certificate")


def fit_text(c, text, x, y, max_width, font="Helvetica", size=10, min_size=7):
    size_now = size
    while size_now > min_size and stringWidth(text, font, size_now) > max_width:
        size_now -= 0.25
    c.setFont(font, size_now)
    c.drawString(x, y, text)


def wrap_words(text, font, size, max_width):
    words = text.split()
    lines, line = [], ""
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


def make_pdf(data):
    """Create one A4 certificate matching the supplied reference layout."""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4
    c.setFillColorRGB(0.08, 0.08, 0.08)

    # Date: same visual format as reference, e.g. 23-Jun-26.
    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    fit_text(c, date_text, 428, H - 204, 100, size=12.0, min_size=10)

    # TEST CERTIFICATE
    title_y = H - 142
    title = "TEST CERTIFICATE"
    c.setFont("Helvetica-Bold", 16.0)
    c.drawCentredString(W / 2, title_y, title)
    title_width = stringWidth(title, "Helvetica-Bold", 16.0)
    c.setLineWidth(0.8)
    c.line(W / 2 - title_width / 2, title_y - 3,
           W / 2 + title_width / 2, title_y - 3)

    # TO WHOM IT MAY CONCERN
    whom_y = H - 262
    whom = "TO WHOM IT MAY CONCERN"
    c.setFont("Helvetica", 12.0)
    c.drawCentredString(W / 2, whom_y, whom)
    whom_width = stringWidth(whom, "Helvetica", 12.0)
    c.setLineWidth(0.55)
    c.line(W / 2 - whom_width / 2, whom_y - 2.5,
           W / 2 + whom_width / 2, whom_y - 2.5)

    # Intro sentence changes automatically according to selected columns.
    item_name = data["item_name"]
    selected = data["selected_columns"]
    labels = [label for _, label in selected]

    if not labels:
        intro = f"This test certificate is for {item_name} as per below table"
    elif len(labels) == 1:
        intro = f"This test certificate is for {item_name} having {labels[0]} as per below table"
    elif len(labels) == 2:
        intro = f"This test certificate is for {item_name} having {labels[0]} and {labels[1]} as per below table"
    else:
        intro = f"This test certificate is for {item_name} having {', '.join(labels[:-1])} and {labels[-1]} as per below table"

    intro_y = H - 310
    intro_lines = wrap_words(intro, "Helvetica", 12.0, 470)
    c.setFont("Helvetica", 12.0)
    for i, line in enumerate(intro_lines[:2]):
        c.drawString(72, intro_y - i * 17, line)

    # Dynamic table. Full width stays identical; columns resize automatically.
    x_left, x_right = 72, 550
    table_top = H - 359
    header_h, row_h = 31, 31
    table_bottom = table_top - header_h - row_h

    columns = [("item_name", "ITEM", data["item_name"])]
    for key, label in selected:
        columns.append((key, label, data[key]))

    col_width = (x_right - x_left) / len(columns)
    x_positions = [x_left + i * col_width for i in range(len(columns) + 1)]

    c.setLineWidth(0.55)
    for x in x_positions:
        c.line(x, table_top, x, table_bottom)
    c.line(x_left, table_top, x_right, table_top)
    c.line(x_left, table_top - header_h, x_right, table_top - header_h)
    c.line(x_left, table_bottom, x_right, table_bottom)

    c.setFont("Helvetica-Bold", 11.0)
    for i, (_, label, _) in enumerate(columns):
        cx = (x_positions[i] + x_positions[i + 1]) / 2
        c.drawCentredString(cx, table_top - 20, label)

    for i, (_, _, value) in enumerate(columns):
        left, right = x_positions[i], x_positions[i + 1]
        txt = str(value)
        size = 10.5
        while size > 7.0 and stringWidth(txt, "Helvetica", size) > (right - left - 10):
            size -= 0.25
        c.setFont("Helvetica", size)
        c.drawCentredString((left + right) / 2, table_bottom + 9.5, txt)

    # Fixed text below table.
    body_y = H - 452
    c.setFont("Helvetica", 11.8)
    c.drawString(72, body_y,
                 "Specified component is tested for all its functionalities and is performing")
    c.drawString(72, body_y - 18, "satisfactorily and found OK.")

    sold = f'The above said item is sold to {data["sold_to"]}.'
    c.setFont("Helvetica-Oblique", 11.8)
    sold_lines = wrap_words(sold, "Helvetica-Oblique", 11.8, 470)
    sold_y = H - 528
    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, sold_y - i * 17, line)

    # Same authorization/signature area as reference.
    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def parse_date(value):
    """Accept the requested display format DD-MMM-YY, e.g. 23-Jun-26."""
    try:
        return datetime.strptime(value.strip(), "%d-%b-%y").date()
    except ValueError:
        return None


def safe_filename(item_name):
    name = re.sub(r"[^A-Za-z0-9]+", "_", item_name.strip()).strip("_")
    return f"test_certificate_{name or 'item'}.pdf"


# No pre-filled certificate data: every new certificate starts blank.
st.subheader("Certificate details")

left, right = st.columns(2)
with left:
    date_text = st.text_input("Date", value="", placeholder="DD-MMM-YY  (e.g. 23-Jun-26)")
    item_name = st.text_input("Item name", value="", placeholder="Enter item name")

    st.markdown("**Select fields to show in table**")
    show_model = st.checkbox("Model No.", value=False)
    show_serial = st.checkbox("Serial No.", value=False)
    show_make = st.checkbox("Make", value=False)

with right:
    # These text boxes appear/disappear immediately when checkboxes change.
    model_no = st.text_input("Model No. value", value="", placeholder="Enter Model No.") if show_model else ""
    serial_no = st.text_input("Serial No. value", value="", placeholder="Enter Serial No.") if show_serial else ""
    make = st.text_input("Make value", value="", placeholder="Enter Make") if show_make else ""
    sold_to = st.text_input("Sold to", value="", placeholder="Enter customer / sold-to name")

copies = st.number_input("Number of print copies", min_value=1, max_value=20, value=1, step=1)

selected_columns = []
if show_model:
    selected_columns.append(("model_no", "Model No."))
if show_serial:
    selected_columns.append(("serial_no", "Serial No."))
if show_make:
    selected_columns.append(("make", "Make"))

if st.button("Generate Demo / PDF", type="primary", use_container_width=True):
    parsed_date = parse_date(date_text)
    required = {"Date": date_text, "Item name": item_name, "Sold to": sold_to}
    if show_model:
        required["Model No."] = model_no
    if show_serial:
        required["Serial No."] = serial_no
    if show_make:
        required["Make"] = make

    missing = [name for name, value in required.items() if not value.strip()]
    if parsed_date is None and "Date" not in missing:
        st.error("Date must be in DD-MMM-YY format, for example 23-Jun-26.")
    elif missing:
        st.error("Please fill in: " + ", ".join(missing))
    else:
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
        st.success("Demo generated. This is how the certificate will print on A4 at 100% / Actual size.")


if "mcert_pdf" in st.session_state:
    st.download_button(
        "Download PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get("mcert_filename", "test_certificate_item.pdf"),
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )
    st.markdown(f"**Print copies:** {st.session_state.get('mcert_copies', 1)}")
    st.caption(
        "For your Wi-Fi HP LaserJet Pro MFP: open the downloaded PDF → Print → "
        "select the HP LaserJet Pro MFP → A4 → Scale 100% / Actual size."
    )
