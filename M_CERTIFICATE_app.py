import io
from datetime import date

import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="M-CERTIFICATE", page_icon="📄", layout="wide")

st.markdown("""
<style>
.stApp { background: #0b1220; color: #eef2ff; }
.block-container { max-width: 1150px; padding-top: 1.5rem; }
h1, h2, h3, label, p { color: #eef2ff !important; }
div[data-testid="stForm"] {
    border: 1px solid #26354d;
    border-radius: 14px;
    padding: 22px;
    background: #111c2e;
}
div.stButton > button, div.stDownloadButton > button {
    border-radius: 9px;
    font-weight: 700;
}
.small-note { color: #a9b7ce; font-size: 0.92rem; }
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("Editable test certificate generator · A4 PDF · Layout matched to the supplied certificate image")
st.info(
    "Select which optional columns should appear in the certificate. "
    "The table automatically resizes to use the same full table width when columns are hidden."
)


def fit_text(c, text, x, y, max_width, font="Helvetica", size=10, min_size=7):
    """Draw text at a fixed position, reducing font size only if necessary."""
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
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4

    # The coordinates below are intentionally fixed to reproduce the
    # supplied certificate image's spacing and positions.
    black = (0.08, 0.08, 0.08)
    c.setFillColorRGB(*black)

    # ---- Header / date ----
    # Date format intentionally remains exactly like the reference:
    # 23-Jun-26
    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    fit_text(c, date_text, 428, H - 204, 100, size=12.0, min_size=10)

    # TEST CERTIFICATE
    title_y = H - 142
    c.setFont("Helvetica-Bold", 16.0)
    c.drawCentredString(W / 2, title_y, "TEST CERTIFICATE")
    title_width = stringWidth("TEST CERTIFICATE", "Helvetica-Bold", 16.0)
    c.setLineWidth(0.8)
    c.line(
        W / 2 - title_width / 2,
        title_y - 3,
        W / 2 + title_width / 2,
        title_y - 3,
    )

    # TO WHOM IT MAY CONCERN
    whom_y = H - 262
    c.setFont("Helvetica", 12.0)
    whom = "TO WHOM IT MAY CONCERN"
    c.drawCentredString(W / 2, whom_y, whom)
    whom_width = stringWidth(whom, "Helvetica", 12.0)
    c.setLineWidth(0.55)
    c.line(
        W / 2 - whom_width / 2,
        whom_y - 2.5,
        W / 2 + whom_width / 2,
        whom_y - 2.5,
    )

    # ---- Intro sentence ----
    item_name = data["item_name"]
    selected = data["selected_columns"]

    if selected:
        labels = [label for key, label in selected]
        if len(labels) == 1:
            field_text = labels[0]
        elif len(labels) == 2:
            field_text = f"{labels[0]} and {labels[1]}"
        else:
            field_text = ", ".join(labels[:-1]) + " and " + labels[-1]
        intro = f"This test certificate is for {item_name} having {field_text} as per below table:"
    else:
        intro = f"This test certificate is for {item_name} as per below table:"

    intro_lines = wrap_words(intro, "Helvetica", 12.0, 470)
    intro_y = H - 310
    c.setFont("Helvetica", 12.0)
    for i, line in enumerate(intro_lines[:2]):
        c.drawString(72, intro_y - i * 17, line)

    # ---- Dynamic table ----
    # Fixed left/right edges; only the number/width of columns changes.
    x_left = 72
    x_right = 550
    table_top = H - 359
    header_h = 31
    row_h = 31
    table_bottom = table_top - header_h - row_h

    columns = [("item_name", "ITEM", data["item_name"])]
    for key, label in selected:
        columns.append((key, label, data[key]))

    col_width = (x_right - x_left) / len(columns)
    x_positions = [x_left + i * col_width for i in range(len(columns) + 1)]

    c.setLineWidth(0.55)

    # Vertical lines
    for x in x_positions:
        c.line(x, table_top, x, table_bottom)

    # Horizontal lines
    c.line(x_left, table_top, x_right, table_top)
    c.line(x_left, table_top - header_h, x_right, table_top - header_h)
    c.line(x_left, table_bottom, x_right, table_bottom)

    # Header row
    c.setFont("Helvetica-Bold", 11.0)
    for i, (_, label, _) in enumerate(columns):
        cx = (x_positions[i] + x_positions[i + 1]) / 2
        c.drawCentredString(cx, table_top - 20, label)

    # Value row
    for i, (_, _, value) in enumerate(columns):
        left = x_positions[i]
        right = x_positions[i + 1]
        txt = str(value)
        size = 10.5
        while size > 7.0 and stringWidth(txt, "Helvetica", size) > (right - left - 10):
            size -= 0.25
        c.setFont("Helvetica", size)
        c.drawCentredString((left + right) / 2, table_bottom + 9.5, txt)

    # ---- Fixed certificate text ----
    body_y = H - 452
    c.setFont("Helvetica", 11.8)
    c.drawString(
        72,
        body_y,
        "Specified component is tested for all its functionalities and is performing",
    )
    c.drawString(72, body_y - 18, "satisfactorily and found OK.")

    # ---- Sold-to ----
    sold = f'The above said item is sold to {data["sold_to"]}.'
    c.setFont("Helvetica-Oblique", 11.8)
    sold_lines = wrap_words(sold, "Helvetica-Oblique", 11.8, 470)
    sold_y = H - 528
    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, sold_y - i * 17, line)

    # ---- Authorization ----
    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


with st.form("certificate_form"):
    st.subheader("Certificate details")

    a, b = st.columns(2)

    with a:
        cert_date = st.date_input("Date", value=date.today(), format="DD/MM/YYYY")
        item_name = st.text_input("Item name", value="LENS")

        st.markdown("**Optional table columns**")
        show_model = st.checkbox("Model No.", value=True)
        show_serial = st.checkbox("Serial No.", value=False)
        show_make = st.checkbox("Make", value=True)

    with b:
        if show_model:
            model_no = st.text_input("Model No. value", value="COR-0810K-20M")
        else:
            model_no = ""

        if show_serial:
            serial_no = st.text_input("Serial No. value", value="")
        else:
            serial_no = ""

        if show_make:
            make = st.text_input("Make value", value="OPTO PIXEL")
        else:
            make = ""

        sold_to = st.text_input(
            "Sold to",
            value="Amneal Ireland Ltd., Ireland",
        )

    copies = st.number_input(
        "Number of print copies",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
    )

    submitted = st.form_submit_button(
        "Generate certificate PDF",
        type="primary",
        use_container_width=True,
    )


if submitted:
    selected_columns = []
    if show_model:
        selected_columns.append(("model_no", "MODEL NO."))
    if show_serial:
        selected_columns.append(("serial_no", "SERIAL NO."))
    if show_make:
        selected_columns.append(("make", "MAKE"))

    required = {
        "Item name": item_name,
        "Sold to": sold_to,
    }

    if show_model:
        required["Model No."] = model_no
    if show_serial:
        required["Serial No."] = serial_no
    if show_make:
        required["Make"] = make

    missing = [name for name, value in required.items() if not value.strip()]

    if missing:
        st.error("Please fill in: " + ", ".join(missing))
    else:
        pdf_bytes = make_pdf(
            {
                "date": cert_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns,
            }
        )

        st.session_state["mcert_pdf"] = pdf_bytes
        st.session_state["mcert_filename"] = (
            f"M-CERTIFICATE_{cert_date.strftime('%Y%m%d')}.pdf"
        )
        st.session_state["mcert_copies"] = int(copies)
        st.success("PDF generated. Download it, then print at 100% / Actual size.")


if "mcert_pdf" in st.session_state:
    st.download_button(
        "Download PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get("mcert_filename", "M-CERTIFICATE.pdf"),
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

    st.markdown(
        f"**Print copies selected:** {st.session_state.get('mcert_copies', 1)}"
    )
    st.caption(
        "The PDF is one A4 certificate page. In the print dialog, set "
        "Copies as required and use 100% / Actual size."
    )

    st.markdown(
        "**Print:** Open the downloaded PDF → Print → select your Wi-Fi printer "
        "→ A4 → Scale 100% / Actual size."
    )

with st.expander("Wi-Fi printer and hosting note"):
    st.write(
        "A Streamlit website cannot reliably discover or directly control printers "
        "connected to your phone/PC's Wi-Fi network. The app generates a PDF; "
        "your device's PDF viewer or browser print dialog lists printers available "
        "to that device."
    )
