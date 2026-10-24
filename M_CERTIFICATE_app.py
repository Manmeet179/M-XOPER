import io
from datetime import date
from xml.sax.saxutils import escape

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
div[data-testid="stForm"] { border: 1px solid #26354d; border-radius: 14px; padding: 22px; background: #111c2e; }
div.stButton > button, div.stDownloadButton > button { border-radius: 9px; font-weight: 700; }
.small-note { color: #a9b7ce; font-size: 0.92rem; }
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("Editable test certificate generator · A4 PDF · Designed for pre-printed OptoPixel letterhead")
st.info("This app prints only the certificate text onto your existing letterhead. The logo, header, footer and stamp are not added to the PDF, so load the pre-printed letterhead paper in your printer.")

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
    # All coordinates are aligned to the supplied certificate's body area.
    # Pre-printed letterhead remains visible because this PDF has no background.
    black = (0.18, 0.18, 0.18)
    c.setFillColorRGB(*black)

    # Date, positioned in the upper-right body area.
    date_text = "Date: " + data["date"].strftime("%d-%b-%y")
    fit_text(c, date_text, 418, H - 91, 105, size=9.2)

    # Fixed headings.
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(W / 2, H - 61, "TEST CERTIFICATE")
    c.setLineWidth(0.7)
    c.line(W/2 - 49, H - 63, W/2 + 49, H - 63)
    c.setFont("Helvetica", 9.5)
    c.drawCentredString(W / 2, H - 116, "TO WHOM IT MAY CONCERN")
    c.setLineWidth(0.5)
    c.line(W/2 - 63, H - 118, W/2 + 63, H - 118)

    # Editable introductory sentence, wrapping to the same two-line area.
    intro = f'This test certificate is for {data["item_name"]} having Model No. and Serial No. as per'
    intro2 = "below table:"
    fit_text(c, intro, 79, H - 141, 455, size=9.1)
    c.setFont("Helvetica", 9.1)
    c.drawString(79, H - 155, intro2)

    # Editable table: fixed table footprint, four columns, no layout changes.
    x0, x1, x2, x3, x4 = 79, 192, 305, 418, 503
    y_top = H - 168
    row_h = 17
    y1 = y_top - row_h
    y2 = y1 - row_h
    c.setLineWidth(0.45)
    for x in (x0, x1, x2, x3, x4):
        c.line(x, y_top, x, y2)
    for y in (y_top, y1, y2):
        c.line(x0, y, x4, y)
    headers = [("ITEM", (x0+x1)/2), ("MODEL NO.", (x1+x2)/2),
               ("SERIAL NO.", (x2+x3)/2), ("MAKE", (x3+x4)/2)]
    c.setFont("Helvetica-Bold", 8.4)
    for label, cx in headers:
        c.drawCentredString(cx, y_top - 11.5, label)
    values = [data["item_name"], data["model_no"], data["serial_no"], data["make"]]
    bounds = [(x0, x1), (x1, x2), (x2, x3), (x3, x4)]
    for val, (left, right) in zip(values, bounds):
        txt = str(val)
        size = 8.1
        while size > 5.8 and stringWidth(txt, "Helvetica", size) > (right-left-5):
            size -= 0.2
        c.setFont("Helvetica", size)
        c.drawCentredString((left+right)/2, y2 + 5.5, txt)

    # These statements remain fixed as requested.
    c.setFont("Helvetica", 9.1)
    c.drawString(79, H - 214, "Specified component is tested for all its functionalities and is performing")
    c.drawString(79, H - 228, "satisfactorily and found OK.")

    # Editable sold-to sentence; wraps automatically but stays in its allocated area.
    sold = f'The above said item is sold to {data["sold_to"]}.'
    c.setFont("Helvetica-Oblique", 9.1)
    lines = wrap_words(sold, "Helvetica-Oblique", 9.1, 455)
    for i, line in enumerate(lines[:2]):
        c.drawString(79, H - 253 - i*13, line)

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

with st.form("certificate_form"):
    st.subheader("Certificate details")
    a, b = st.columns(2)
    with a:
        cert_date = st.date_input("Date", value=date.today(), format="DD-MMM-YYYY")
        item_name = st.text_input("Item name (editable)", value="OPT LAN CABLE")
        model_no = st.text_input("Model No.", value="OPTL_LAN3")
    with b:
        serial_no = st.text_input("Serial No.", value="OLAN3024")
        make = st.text_input("Make", value="OptoPixel")
        sold_to = st.text_input("Sold to", value="Amneal Ireland Ltd., Ireland")
    copies = st.number_input("Number of print copies", min_value=1, max_value=20, value=1, step=1)
    submitted = st.form_submit_button("Generate certificate PDF", type="primary", use_container_width=True)

if submitted:
    required = {
        "Item name": item_name,
        "Model No.": model_no,
        "Serial No.": serial_no,
        "Make": make,
        "Sold to": sold_to,
    }
    missing = [name for name, value in required.items() if not value.strip()]
    if missing:
        st.error("Please fill in: " + ", ".join(missing))
    else:
        pdf_bytes = make_pdf({
            "date": cert_date,
            "item_name": item_name.strip(),
            "model_no": model_no.strip(),
            "serial_no": serial_no.strip(),
            "make": make.strip(),
            "sold_to": sold_to.strip(),
        })
        st.session_state["mcert_pdf"] = pdf_bytes
        st.session_state["mcert_filename"] = f"M-CERTIFICATE_{cert_date.strftime('%Y%m%d')}.pdf"
        st.session_state["mcert_copies"] = int(copies)
        st.success("PDF generated. Download it, then print using your Wi-Fi printer.")

if "mcert_pdf" in st.session_state:
    st.download_button(
        "Download PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get("mcert_filename", "M-CERTIFICATE.pdf"),
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )
    st.markdown(f"**Print copies selected:** {st.session_state.get('mcert_copies', 1)}")
    st.caption("The PDF itself is one certificate page. In the print dialog, set Copies to the number shown above. The default is 1.")
    st.markdown("**Print:** Open the downloaded PDF and choose Print → select your Wi-Fi printer → set paper to A4 and scale to 100% / Actual size. Use the pre-printed letterhead tray.")

with st.expander("Wi-Fi printer and hosting note"):
    st.write(
        "A Streamlit website cannot reliably discover or directly control printers connected to your phone/PC's Wi-Fi network. "
        "This app generates a PDF; your device's PDF viewer or browser print dialog lists printers available to that device. "
        "If you host this app online, printing is initiated from your device, not from the Streamlit server."
    )
