import io
import re
from datetime import datetime

import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="📄",
    layout="wide"
)

st.title("M-CERTIFICATE")
st.caption("Generate, Preview and Print Certificate")

st.markdown("""
<style>
.stApp {
    background: #0b1220;
    color: #eef2ff;
}
h1, h2, h3, p, label {
    color: #eef2ff !important;
}
</style>
""", unsafe_allow_html=True)


def parse_date(value):
    try:
        return datetime.strptime(
            value.strip(), "%d-%b-%y"
        ).date()
    except ValueError:
        return None


def safe_filename(name):
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return f"test_certificate_{name or 'item'}.pdf"


def wrap_text(text, font, size, max_width):
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


def make_certificate(data):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4

    c.setFillColorRGB(0.08, 0.08, 0.08)

    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    c.setFont("Helvetica", 12)
    c.drawString(428, H - 204, date_text)

    title = "TEST CERTIFICATE"
    title_y = H - 142

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)

    tw = stringWidth(title, "Helvetica-Bold", 16)
    c.line(W / 2 - tw / 2, title_y - 3,
           W / 2 + tw / 2, title_y - 3)

    whom = "TO WHOM IT MAY CONCERN"
    whom_y = H - 262

    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, whom_y, whom)

    ww = stringWidth(whom, "Helvetica", 12)
    c.line(W / 2 - ww / 2, whom_y - 2.5,
           W / 2 + ww / 2, whom_y - 2.5)

    labels = [label for _, label in data["selected_columns"]]
    item = data["item_name"]

    if len(labels) == 0:
        intro = f"This test certificate is for {item} as per below table"
    elif len(labels) == 1:
        intro = f"This test certificate is for {item} having {labels[0]} as per below table"
    elif len(labels) == 2:
        intro = f"This test certificate is for {item} having {labels[0]} and {labels[1]} as per below table"
    else:
        intro = f"This test certificate is for {item} having {', '.join(labels[:-1])} and {labels[-1]} as per below table"

    intro_lines = wrap_text(intro, "Helvetica", 12, 470)
    c.setFont("Helvetica", 12)

    for i, line in enumerate(intro_lines[:2]):
        c.drawString(72, H - 310 - i * 17, line)

    x_left, x_right = 72, 550
    table_top = H - 359
    header_h, row_h = 31, 31
    table_bottom = table_top - header_h - row_h

    columns = [("item_name", "ITEM", item)]
    for key, label in data["selected_columns"]:
        columns.append((key, label, data[key]))

    col_width = (x_right - x_left) / len(columns)
    xs = [x_left + i * col_width for i in range(len(columns) + 1)]

    for x in xs:
        c.line(x, table_top, x, table_bottom)

    c.line(x_left, table_top, x_right, table_top)
    c.line(x_left, table_top - header_h, x_right, table_top - header_h)
    c.line(x_left, table_bottom, x_right, table_bottom)

    c.setFont("Helvetica-Bold", 11)
    for i, (_, label, _) in enumerate(columns):
        cx = (xs[i] + xs[i + 1]) / 2
        c.drawCentredString(cx, table_top - 20, label)

    for i, (_, _, value) in enumerate(columns):
        left, right = xs[i], xs[i + 1]
        txt = str(value)
        size = 10.5

        while size > 7 and stringWidth(txt, "Helvetica", size) > right - left - 10:
            size -= 0.25

        c.setFont("Helvetica", size)
        c.drawCentredString(
            (left + right) / 2,
            table_bottom + 9.5,
            txt
        )

    body_y = H - 452
    c.setFont("Helvetica", 11.8)
    c.drawString(
        72, body_y,
        "Specified component is tested for all its functionalities and is performing"
    )
    c.drawString(72, body_y - 18, "satisfactorily and found OK.")

    sold = f'The above said item is sold to {data["sold_to"]}.'
    sold_lines = wrap_text(sold, "Helvetica-Oblique", 11.8, 470)

    c.setFont("Helvetica-Oblique", 11.8)
    for i, line in enumerate(sold_lines[:2]):
        c.drawString(72, H - 528 - i * 17, line)

    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# CERTIFICATE FORM
st.header("Certificate Details")

left, right = st.columns(2)

with left:
    date_text = st.text_input(
        "Date (DD-MMM-YY)",
        placeholder="23-Jun-26"
    )
    item_name = st.text_input("Item Name")

    st.markdown("**Select fields for table**")
    show_model = st.checkbox("Model No.")
    show_serial = st.checkbox("Serial No.")
    show_make = st.checkbox("Make")

with right:
    model_no = st.text_input("Model No. Value") if show_model else ""
    serial_no = st.text_input("Serial No. Value") if show_serial else ""
    make = st.text_input("Make Value") if show_make else ""
    sold_to = st.text_input("Sold To / Customer Name")

selected_columns = []

if show_model:
    selected_columns.append(("model_no", "Model No."))
if show_serial:
    selected_columns.append(("serial_no", "Serial No."))
if show_make:
    selected_columns.append(("make", "Make"))


if st.button("📄 GENERATE CERTIFICATE", type="primary",
             use_container_width=True):

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
        if not str(value).strip()
    ]

    if missing:
        st.error("Please fill in: " + ", ".join(missing))
    elif parsed_date is None:
        st.error("Date format must be DD-MMM-YY, e.g. 23-Jun-26.")
    else:
        try:
            pdf_bytes = make_certificate({
                "date": parsed_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns
            })

            st.session_state["certificate_pdf"] = pdf_bytes
            st.session_state["certificate_filename"] = safe_filename(item_name)

        except Exception as exc:
            st.error(f"Certificate generation failed: {exc}")


# SHOW GENERATED CERTIFICATE
if "certificate_pdf" in st.session_state:
    st.success("Certificate generated successfully!")

    st.subheader("Certificate Preview")

    # Embed PDF preview in the page.
    import base64

    pdf_base64 = base64.b64encode(
        st.session_state["certificate_pdf"]
    ).decode("utf-8")

    st.components.v1.html(
        f"""
        <iframe
            src="data:application/pdf;base64,{pdf_base64}"
            width="100%"
            height="700"
            style="border:1px solid #555; border-radius:8px;"
        ></iframe>
        """,
        height=720,
        scrolling=True
    )

    st.download_button(
        "⬇️ DOWNLOAD CERTIFICATE PDF",
        data=st.session_state["certificate_pdf"],
        file_name=st.session_state["certificate_filename"],
        mime="application/pdf",
        use_container_width=True
    )

    st.info("PDF download કરો અને તેને ખોલીને Print કરો.")
