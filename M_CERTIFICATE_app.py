
import io
import re

import streamlit as st
import fitz
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth


st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="📄",
    layout="wide"
)


# =========================
# DEFAULT SETTINGS
# =========================

DEFAULTS = {
    "date_format": "DD-MMM-YY",
    "item": "ITEM",
    "model": "Model No.",
    "serial": "Serial No.",
    "make": "Make",
    "mac": "MAC ID",
    "win": "WIN KEY",
    "company": "OptoPixel Pvt. Ltd.",
    "top_spacing_mm": 25,
}

FORMATS = {
    "DD-MMM-YY": "%d-%b-%y",
    "DD-MMM-YY": "%d-%b-%y",
    "DD-MMM-YYYY": "%d-%b-%Y",
    "DD/MM/YYYY": "%d/%m/%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
}

EXAMPLES = {
    "DD-MMM-YY": "23-Oct-25",
    "DD-MMM-YY": "22-OCT-26",
    "DD-MMM-YYYY": "22-OCT-2026",
    "DD/MM/YYYY": "22/10/2026",
    "DD-MM-YYYY": "22-10-2026",
    "MM/DD/YYYY": "10/22/2026",
    "YYYY-MM-DD": "2026-10-22",
}


# Initialize settings
if "pdf_settings" not in st.session_state:
    st.session_state.pdf_settings = DEFAULTS.copy()

else:
    # Add any newly introduced settings without
    # removing existing user preferences.
    for key, value in DEFAULTS.items():
        st.session_state.pdf_settings.setdefault(key, value)


# =========================
# TEXT WRAPPING
# =========================

def wrap_text(text, font, size, width):
    """Wrap text, including long values without spaces."""

    result = []
    line = ""

    for word in str(text).split():
        candidate = word if not line else line + " " + word

        if stringWidth(candidate, font, size) <= width:
            line = candidate
            continue

        if line:
            result.append(line)
            line = ""

        if stringWidth(word, font, size) <= width:
            line = word

        else:
            chunk = ""

            for char in word:
                if stringWidth(
                    chunk + char, font, size
                ) <= width:
                    chunk += char

                else:
                    if chunk:
                        result.append(chunk)

                    chunk = char

            line = chunk

    if line:
        result.append(line)

    return result or [""]


def safe_filename(name):
    name = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        name
    ).strip("_")

    return f"test_certificate_{name or 'item'}.pdf"


# =========================
# CREATE PDF
# =========================

def create_pdf(data, settings):
    buffer = io.BytesIO()

    pdf = canvas.Canvas(
        buffer,
        pagesize=A4
    )

    page_w, page_h = A4

    left = 48
    right = page_w - 48
    content_w = right - left

    # Adjustable top spacing
    # 1 mm = 72 / 25.4 PDF points
    top_spacing_mm = float(
        settings.get("top_spacing_mm", 25)
    )

    top_spacing_points = (
        top_spacing_mm * 72 / 25.4
    )

    # Move all certificate content downward
    pdf.translate(
        0,
        -top_spacing_points
    )

    # -------------------------
    # DATE
    # -------------------------

    date_text = data["date"].strftime(
        FORMATS[settings["date_format"]]
    )

    pdf.setFillColorRGB(
        0.08, 0.08, 0.08
    )

    pdf.setFont(
        "Helvetica",
        11
    )

    pdf.drawRightString(
        right,
        page_h - 145,
        "Date: " + date_text
    )

    # -------------------------
    # CERTIFICATE TITLE
    # -------------------------

    title = "TEST CERTIFICATE"
    title_y = page_h - 95

    pdf.setFont(
        "Helvetica-Bold",
        16
    )

    pdf.drawCentredString(
        page_w / 2,
        title_y,
        title
    )

    title_width = stringWidth(
        title,
        "Helvetica-Bold",
        16
    )

    pdf.line(
        page_w / 2 - title_width / 2,
        title_y - 3,
        page_w / 2 + title_width / 2,
        title_y - 3
    )

    # -------------------------
    # TO WHOM IT MAY CONCERN
    # -------------------------

    concern = "TO WHOM IT MAY CONCERN"
    concern_y = page_h - 190

    pdf.setFont(
        "Helvetica",
        12
    )

    pdf.drawCentredString(
        page_w / 2,
        concern_y,
        concern
    )

    concern_width = stringWidth(
        concern,
        "Helvetica",
        12
    )

    pdf.line(
        page_w / 2 - concern_width / 2,
        concern_y - 3,
        page_w / 2 + concern_width / 2,
        concern_y - 3
    )

    # -------------------------
    # INTRODUCTION
    # -------------------------

    selected = data["selected"]

    labels = [
        label for _, label in selected
    ]

    intro = (
        f'This test certificate is for {data["item"]}'
    )

    if labels:
        intro += " having " + ", ".join(labels)

    intro += " as per below table"

    intro_lines = wrap_text(
        intro,
        "Helvetica",
        11,
        content_w
    )

    intro_y = page_h - 245

    pdf.setFont(
        "Helvetica",
        11
    )

    for i, line in enumerate(intro_lines):
        pdf.drawString(
            left,
            intro_y - i * 15,
            line
        )

    # -------------------------
    # TABLE DATA
    # -------------------------

    columns = [
        (settings["item"], str(data["item"]))
    ]

    # FIXED: Read values safely from data dictionary
    columns.extend(
        (
            label,
            str(data.get(key, ""))
        )
        for key, label in selected
    )

    count = len(columns)

    if count == 1:
        widths = [content_w]

    else:
        widths = (
            [content_w * 0.27]
            + [
                content_w * 0.73 / (count - 1)
            ] * (count - 1)
        )

    font_size = 9
    line_gap = 12

    wrapped_columns = []

    max_header = 1
    max_value = 1

    for i, (heading, value) in enumerate(columns):
        available = widths[i] - 12

        header_lines = wrap_text(
            heading,
            "Helvetica-Bold",
            9,
            available
        )

        value_lines = wrap_text(
            value,
            "Helvetica",
            font_size,
            available
        )

        wrapped_columns.append(
            (header_lines, value_lines)
        )

        max_header = max(
            max_header,
            len(header_lines)
        )

        max_value = max(
            max_value,
            len(value_lines)
        )

    header_h = max(
        30,
        max_header * line_gap + 12
    )

    row_h = max(
        31,
        max_value * line_gap + 12
    )

    table_top = (
        intro_y
        - len(intro_lines) * 15
        - 25
    )

    table_bottom = (
        table_top
        - header_h
        - row_h
    )

    xs = [left]

    for width in widths:
        xs.append(
            xs[-1] + width
        )

    # -------------------------
    # TABLE BORDERS
    # -------------------------

    pdf.setLineWidth(0.6)

    pdf.line(
        left,
        table_top,
        right,
        table_top
    )

    pdf.line(
        left,
        table_top - header_h,
        right,
        table_top - header_h
    )

    pdf.line(
        left,
        table_bottom,
        right,
        table_bottom
    )

    for x in xs:
        pdf.line(
            x,
            table_top,
            x,
            table_bottom
        )

    # -------------------------
    # TABLE HEADINGS AND VALUES
    # -------------------------

    for i, (
        header_lines,
        value_lines
    ) in enumerate(wrapped_columns):

        center_x = (
            xs[i] + xs[i + 1]
        ) / 2

        y = (
            table_top
            - (
                header_h
                - len(header_lines) * line_gap
            ) / 2
            - 9
        )

        pdf.setFont(
            "Helvetica-Bold",
            9
        )

        for line in header_lines:
            pdf.drawCentredString(
                center_x,
                y,
                line
            )

            y -= line_gap

        y = (
            table_top
            - header_h
            - (
                row_h
                - len(value_lines) * line_gap
            ) / 2
            - 8
        )

        pdf.setFont(
            "Helvetica",
            font_size
        )

        for line in value_lines:
            pdf.drawCentredString(
                center_x,
                y,
                line
            )

            y -= line_gap

    # -------------------------
    # TEST STATEMENT
    # -------------------------

    y = table_bottom - 35

    pdf.setFont(
        "Helvetica",
        10.5
    )

    statement = (
        "Specified component is tested for all its functionalities "
        "and is performing satisfactorily and found OK."
    )

    for line in wrap_text(
        statement,
        "Helvetica",
        10.5,
        content_w
    ):
        pdf.drawString(
            left,
            y,
            line
        )

        y -= 15

    y -= 25

    # -------------------------
    # SOLD TO
    # -------------------------

    sold_text = (
        f'The above said item is sold to {data["sold_to"]}.'
    )

    pdf.setFont(
        "Helvetica-Oblique",
        10.5
    )

    for line in wrap_text(
        sold_text,
        "Helvetica-Oblique",
        10.5,
        content_w
    ):
        pdf.drawString(
            left,
            y,
            line
        )

        y -= 15

    # -------------------------
    # COMPANY AND AUTHORIZATION
    # -------------------------

    y -= 25

    pdf.setFont(
        "Helvetica",
        10.5
    )

    pdf.drawString(
        left,
        y,
        f'For {settings["company"]},'
    )

    pdf.drawString(
        left,
        y - 48,
        "Authorized"
    )

    # No "Made by MEET MEVADA" in PDF
    pdf.save()

    buffer.seek(0)

    return buffer.getvalue()


# =========================
# PDF PREVIEW IMAGE
# =========================

def pdf_to_image(pdf_bytes):
    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        pix = doc[0].get_pixmap(
            matrix=fitz.Matrix(1.5, 1.5),
            alpha=False
        )

        return Image.open(
            io.BytesIO(
                pix.tobytes("png")
            )
        ).copy()

    finally:
        doc.close()


# =========================
# SIDEBAR SETTINGS
# =========================

with st.sidebar:

    st.markdown(
        "<h3 style='text-align:center'>"
        "M-CERTIFICATE</h3>",
        unsafe_allow_html=True
    )

    st.markdown(
        "<p style='text-align:center'>"
        "Made by MEET MEVADA</p>",
        unsafe_allow_html=True
    )

    st.divider()

    st.header("PDF Page Settings")

    settings = st.session_state.pdf_settings

    with st.form("settings_form"):

        date_format = st.selectbox(
            "Date format",
            list(FORMATS.keys()),
            index=list(FORMATS).index(
                settings["date_format"]
            ),
            format_func=lambda f: (
                f"{f} ({EXAMPLES[f]})"
            )
        )

        # DEFAULT TOP SPACING = 25 MM
        top_spacing_mm = st.number_input(
            "PDF Content Top Spacing (mm)",
            min_value=0,
            max_value=60,
            value=int(
                settings.get("top_spacing_mm", 25)
            ),
            step=1,
            help=(
                "Increase the value to move the "
                "certificate content downward."
            )
        )

        st.caption("Table headings")

        item_heading = st.text_input(
            "ITEM heading",
            settings["item"]
        )

        model_heading = st.text_input(
            "Model No. heading",
            settings["model"]
        )

        serial_heading = st.text_input(
            "Serial No. heading",
            settings["serial"]
        )

        make_heading = st.text_input(
            "Make heading",
            settings["make"]
        )

        mac_heading = st.text_input(
            "MAC ID heading",
            settings["mac"]
        )

        win_heading = st.text_input(
            "WIN KEY heading",
            settings["win"]
        )

        company_name = st.text_input(
            "Company name",
            settings["company"]
        )

        save_settings = st.form_submit_button(
            "Save Settings",
            use_container_width=True
        )

        if save_settings:

            st.session_state.pdf_settings = {
                "date_format": date_format,
                "top_spacing_mm": int(top_spacing_mm),

                "item": (
                    item_heading.strip()
                    or DEFAULTS["item"]
                ),

                "model": (
                    model_heading.strip()
                    or DEFAULTS["model"]
                ),

                "serial": (
                    serial_heading.strip()
                    or DEFAULTS["serial"]
                ),

                "make": (
                    make_heading.strip()
                    or DEFAULTS["make"]
                ),

                "mac": (
                    mac_heading.strip()
                    or DEFAULTS["mac"]
                ),

                "win": (
                    win_heading.strip()
                    or DEFAULTS["win"]
                ),

                "company": (
                    company_name.strip()
                    or DEFAULTS["company"]
                ),
            }

            # Remove old PDF after settings change
            st.session_state.pop(
                "certificate_pdf",
                None
            )

            st.rerun()

    if st.button(
        "Restore Default Settings",
        use_container_width=True
    ):

        st.session_state.pdf_settings = (
            DEFAULTS.copy()
        )

        st.session_state.pop(
            "certificate_pdf",
            None
        )

        st.rerun()


settings = st.session_state.pdf_settings


# =========================
# HOME PAGE
# =========================

st.markdown(
    "<h1 style='text-align:center'>"
    "M-CERTIFICATE</h1>",
    unsafe_allow_html=True
)

st.markdown(
    "<p style='text-align:center'>"
    "Certificate Generator</p>",
    unsafe_allow_html=True
)

st.header("Certificate Details")

left_col, right_col = st.columns(2)


with left_col:

    selected_date = st.date_input(
        "Date",
        value=None,
        format="DD/MM/YYYY",
        help="Select a date from the calendar."
    )

    item_name = st.text_input(
        "Item Name"
    )

    st.markdown(
        "**Select fields to show in table**"
    )

    show_model = st.checkbox("Model No.")
    show_serial = st.checkbox("Serial No.")
    show_make = st.checkbox("Make")
    show_mac = st.checkbox("MAC ID")
    show_win = st.checkbox("WIN KEY")


with right_col:

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

    mac = (
        st.text_input("MAC ID Value")
        if show_mac else ""
    )

    win = (
        st.text_input("WIN KEY Value")
        if show_win else ""
    )

    sold_to = st.text_input(
        "Sold To / Customer Name"
    )


# =========================
# SELECTED TABLE FIELDS
# =========================

selected = []

for enabled, key, label in [
    (show_model, "model_no", settings["model"]),
    (show_serial, "serial_no", settings["serial"]),
    (show_make, "make", settings["make"]),
    (show_mac, "mac", settings["mac"]),
    (show_win, "win", settings["win"]),
]:

    if enabled:
        selected.append(
            (key, label)
        )


# =========================
# GENERATE CERTIFICATE
# =========================

if st.button(
    "📄 GENERATE CERTIFICATE",
    type="primary",
    use_container_width=True
):

    required = {
        "Date": selected_date,
        "Item Name": item_name,
        "Sold To / Customer Name": sold_to,
    }

    if show_model:
        required["Model No."] = model_no

    if show_serial:
        required["Serial No."] = serial_no

    if show_make:
        required["Make"] = make

    if show_mac:
        required["MAC ID"] = mac

    if show_win:
        required["WIN KEY"] = win

    missing = [
        label
        for label, value in required.items()
        if value is None or not str(value).strip()
    ]

    if missing:

        st.error(
            "Please fill in: "
            + ", ".join(missing)
        )

    else:

        # FIXED: Include every field value in data
        pdf_data = {
            "date": selected_date,
            "item": item_name.strip(),
            "sold_to": sold_to.strip(),
            "selected": selected,

            "model_no": model_no,
            "serial_no": serial_no,
            "make": make,
            "mac": mac,
            "win": win,
        }

        pdf_bytes = create_pdf(
            pdf_data,
            settings
        )

        st.session_state.certificate_pdf = (
            pdf_bytes
        )

        st.session_state.certificate_filename = (
            safe_filename(item_name)
        )

        st.success(
            "Certificate generated successfully!"
        )


# =========================
# PDF PREVIEW AND DOWNLOAD
# =========================

if "certificate_pdf" in st.session_state:

    st.subheader("Certificate Preview")

    try:

        preview_image = pdf_to_image(
            st.session_state.certificate_pdf
        )

        st.image(
            preview_image,
            use_container_width=True
        )

    except Exception as exc:

        st.warning(
            f"Preview could not be displayed: {exc}"
        )

    st.download_button(
        "⬇️ DOWNLOAD CERTIFICATE PDF",

        data=st.session_state.certificate_pdf,

        file_name=st.session_state.get(
            "certificate_filename",
            "test_certificate.pdf"
        ),

        mime="application/pdf",

        use_container_width=True
    )

    st.caption(
        "To print, download the PDF and open it "
        "in a PDF viewer."
    )


# =========================
# WEBSITE FOOTER ONLY
# =========================

st.markdown(
    "<p style='text-align:center; "
    "font-size:12px; padding-top:18px'>"
    "Made by MEET MEVADA</p>",
    unsafe_allow_html=True
)
