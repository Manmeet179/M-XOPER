import streamlit as st
import io
import base64
from datetime import date
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
)

# --------------------------------------------------
# PAGE SETTINGS
# --------------------------------------------------

st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="📄",
    layout="wide"
)

DEFAULT_HEADINGS = {
    "item": "ITEM",
    "model_no": "MODEL NO.",
    "serial_no": "SERIAL NO.",
    "make": "MAKE",
    "mac": "MAC ID",
    "win_key": "WIN KEY",
}

DATE_FORMATS = {
    "23-Jun-26": "%d-%b-%y",
    "23-JUN-26": "%d-%b-%y",
    "23-06-2026": "%d-%m-%Y",
    "23/06/2026": "%d/%m/%Y",
    "23-Jun-2026": "%d-%b-%Y",
    "2026-06-23": "%Y-%m-%d",
}

# --------------------------------------------------
# CSS
# --------------------------------------------------

st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        font-size: 42px;
        font-weight: 800;
        margin-top: 10px;
        margin-bottom: 25px;
    }

    .made-by {
        text-align: center;
        color: #777777;
        font-size: 13px;
        padding-top: 15px;
    }

    div[data-testid="stFormSubmitButton"] button {
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

if "certificate_headings" not in st.session_state:
    st.session_state.certificate_headings = DEFAULT_HEADINGS.copy()

if "certificate_company" not in st.session_state:
    st.session_state.certificate_company = "OptoPixel Pvt. Ltd."

if "certificate_date_format" not in st.session_state:
    st.session_state.certificate_date_format = "23-Jun-26"

if "certificate_pdf" not in st.session_state:
    st.session_state.certificate_pdf = None

# --------------------------------------------------
# PDF GENERATION
# --------------------------------------------------

def safe_paragraph(value):
    """Escape user-entered text so special characters display safely."""
    value = "" if value is None else str(value)
    value = escape(value)
    value = value.replace("\n", "<br/>")
    return value if value.strip() else "&#160;"


def create_pdf(
    selected_date,
    date_format,
    company_name,
    item_name,
    model_no,
    serial_no,
    make_value,
    mac_value,
    win_key,
    headings,
    customer_statement,
):
    pdf_buffer = io.BytesIO()

    # A4 page, margins similar to the supplied reference image.
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=17 * mm,
        leftMargin=17 * mm,
        topMargin=32 * mm,
        bottomMargin=20 * mm,
        title="TEST CERTIFICATE",
        author="",
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CertificateTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=21,
        alignment=TA_CENTER,
        underline=True,
        spaceAfter=0,
    )

    date_style = ParagraphStyle(
        "CertificateDate",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=12,
        leading=16,
        alignment=TA_RIGHT,
    )

    concern_style = ParagraphStyle(
        "Concern",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=12,
        leading=16,
        alignment=TA_CENTER,
        underline=True,
    )

    body_style = ParagraphStyle(
        "CertificateBody",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=12,
        leading=17,
        alignment=TA_LEFT,
        wordWrap="CJK",
    )

    italic_style = ParagraphStyle(
        "CustomerStatement",
        parent=body_style,
        fontName="Helvetica-Oblique",
    )

    company_style = ParagraphStyle(
        "Company",
        parent=body_style,
        fontName="Helvetica",
    )

    header_style = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
        wordWrap="CJK",
    )

    cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        alignment=TA_CENTER,
        wordWrap="CJK",
    )

    story = []

    # 1. Title near the upper-middle area of the page
    story.append(Spacer(1, 28 * mm))
    story.append(
        Paragraph(
            "<u>TEST CERTIFICATE</u>",
            title_style,
        )
    )

    # 2. Date aligned to the right
    story.append(Spacer(1, 18 * mm))

    if selected_date is not None:
        formatted_date = selected_date.strftime(date_format)
        story.append(
            Paragraph(
                f"Date: {escape(formatted_date)}",
                date_style,
            )
        )
    else:
        story.append(Paragraph("Date:", date_style))

    # 3. Centered heading
    story.append(Spacer(1, 16 * mm))
    story.append(
        Paragraph(
            "<u>TO WHOM IT MAY CONCERN</u>",
            concern_style,
        )
    )

    # 4. Introductory text
    story.append(Spacer(1, 16 * mm))
    story.append(
        Paragraph(
            "This test certificate is for "
            f"{safe_paragraph(item_name)} having Model No. and "
            "Serial No. as per below table:",
            body_style,
        )
    )

    # 5. Table - show only columns with headings
    all_columns = [
        ("item", item_name),
        ("model_no", model_no),
        ("serial_no", serial_no),
        ("make", make_value),
        ("mac", mac_value),
        ("win_key", win_key),
    ]

    visible_columns = [
        (key, value)
        for key, value in all_columns
        if str(headings.get(key, "")).strip()
    ]

    if not visible_columns:
        visible_columns = [("item", item_name)]

    table_data = [
        [
            Paragraph(
                safe_paragraph(headings.get(key, "")),
                header_style,
            )
            for key, _ in visible_columns
        ],
        [
            Paragraph(
                safe_paragraph(value),
                cell_style,
            )
            for _, value in visible_columns
        ],
    ]

    available_width = A4[0] - 34 * mm
    column_count = len(visible_columns)

    # Equal widths keep the table inside the page margins.
    col_widths = [available_width / column_count] * column_count

    certificate_table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1,
        hAlign="CENTER",
    )

    certificate_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )

    story.append(Spacer(1, 12 * mm))
    story.append(certificate_table)

    # 6. Test result statement
    story.append(Spacer(1, 11 * mm))
    story.append(
        Paragraph(
            "Specified component is tested for all its functionalities "
            "and is performing satisfactorily and found OK.",
            body_style,
        )
    )

    # 7. Customer statement
    story.append(Spacer(1, 12 * mm))
    story.append(
        Paragraph(
            safe_paragraph(customer_statement),
            italic_style,
        )
    )

    # 8. Company and authorization
    story.append(Spacer(1, 15 * mm))
    story.append(
        Paragraph(
            f"For {safe_paragraph(company_name)},",
            company_style,
        )
    )

    story.append(Spacer(1, 13 * mm))
    story.append(Paragraph("Authorized", company_style))

    # IMPORTANT:
    # No PDF footer is added here.
    # "Made by MEET MEVADA" will NOT appear in the PDF.
    doc.build(story)

    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()


# --------------------------------------------------
# SIDEBAR
# --------------------------------------------------

with st.sidebar:
    st.title("Settings")

    st.subheader("Date Format")

    selected_date_format_label = st.selectbox(
        "Select date format",
        options=list(DATE_FORMATS.keys()),
        index=list(DATE_FORMATS.keys()).index(
            st.session_state.certificate_date_format
        ),
    )

    st.session_state.certificate_date_format = (
        DATE_FORMATS[selected_date_format_label]
    )

    st.divider()
    st.subheader("Table Headings")

    current_headings = st.session_state.certificate_headings

    for key, default_value in DEFAULT_HEADINGS.items():
        current_value = current_headings.get(key, default_value)
        current_headings[key] = st.text_input(
            f"{default_value} heading",
            value=current_value,
            key=f"heading_{key}",
        )

    st.divider()
    st.subheader("Company Details")

    st.session_state.certificate_company = st.text_input(
        "Company name in PDF",
        value=st.session_state.certificate_company,
    )

    if st.button("Restore Default Settings", use_container_width=True):
        st.session_state.certificate_headings = DEFAULT_HEADINGS.copy()
        st.session_state.certificate_company = "OptoPixel Pvt. Ltd."
        st.session_state.certificate_date_format = "%d-%b-%y"

        for key, value in DEFAULT_HEADINGS.items():
            st.session_state[f"heading_{key}"] = value

        st.rerun()

    st.divider()
    st.markdown(
        '<div class="made-by">Made by MEET MEVADA</div>',
        unsafe_allow_html=True,
    )

# --------------------------------------------------
# MAIN PAGE
# --------------------------------------------------

st.markdown(
    '<div class="main-title">M-CERTIFICATE</div>',
    unsafe_allow_html=True,
)

st.subheader("Certificate Details")

with st.form("certificate_form"):
    col1, col2 = st.columns(2)

    with col1:
        selected_date = st.date_input(
            "Certificate Date",
            value=None,
            format="DD/MM/YYYY",
        )

        item_name = st.text_input(
            "Item",
            value="CAMERA",
        )

        model_no = st.text_input("Model No.")

        serial_no = st.text_input("Serial No.")

    with col2:
        make_value = st.text_input("Make")

        mac_value = st.text_input("MAC ID")

        win_key = st.text_input("WIN KEY")

        customer_statement = st.text_input(
            "Customer statement",
            value="The above said item is sold to Amneal Ireland Ltd., Ireland.",
        )

    generate = st.form_submit_button(
        "Generate PDF",
        use_container_width=True,
    )

if generate:
    try:
        pdf_bytes = create_pdf(
            selected_date=selected_date,
            date_format=st.session_state.certificate_date_format,
            company_name=st.session_state.certificate_company,
            item_name=item_name.strip(),
            model_no=model_no.strip(),
            serial_no=serial_no.strip(),
            make_value=make_value.strip(),
            mac_value=mac_value.strip(),
            win_key=win_key.strip(),
            headings=st.session_state.certificate_headings.copy(),
            customer_statement=customer_statement.strip(),
        )

        st.session_state.certificate_pdf = pdf_bytes

    except Exception as error:
        st.error(f"PDF generation failed: {error}")

# --------------------------------------------------
# PDF PREVIEW + DOWNLOAD
# --------------------------------------------------

if st.session_state.certificate_pdf:
    st.divider()
    st.subheader("Generated Certificate Preview")

    pdf_bytes = st.session_state.certificate_pdf

    base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

    pdf_display = f"""
        <iframe
            src="data:application/pdf;base64,{base64_pdf}"
            width="100%"
            height="850"
            type="application/pdf"
            style="border: 1px solid #cccccc; border-radius: 6px;"
        ></iframe>
    """

    st.markdown(pdf_display, unsafe_allow_html=True)

    st.download_button(
        label="Download Certificate PDF",
        data=pdf_bytes,
        file_name="test_certificate.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

# Footer only on the Streamlit page, never inside the PDF.
st.markdown(
    '<div class="made-by">Made by MEET MEVADA</div>',
    unsafe_allow_html=True,
    )
