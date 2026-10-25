
import io
import re
from datetime import datetime

import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth


st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="📄",
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
st.caption("Certificate Generator • Scan Pages • PDF Download")


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


def make_certificate_pdf(data):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4

    c.setFillColorRGB(0.08, 0.08, 0.08)

    date_text = "Date: " + data["date"].strftime("%d-%b-%y")

    fit_text(
        c, date_text, 428, H - 204, 100,
        size=12, min_size=10
    )

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

    c.setFont("Helvetica-Bold", 11)

    for i, (_, label, _) in enumerate(columns):
        cx = (x_positions[i] + x_positions[i + 1]) / 2
        c.drawCentredString(
            cx, table_top - 20, label
        )

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

    c.setFont("Helvetica", 11.8)
    c.drawString(72, H - 585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H - 640, "Authorized")

    c.showPage()
    c.save()

    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------
# IMAGE / SCANNED PDF HELPERS
# --------------------------------------------------

def normalize_image(image_bytes):
    image = Image.open(io.BytesIO(image_bytes))
    image = ImageOps.exif_transpose(image)
    return image.convert("RGB")


def make_scanned_pdf(pages):
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=A4)

    page_w, page_h = A4
    margin = 24

    for page_bytes in pages:
        image = normalize_image(page_bytes)
        image_w, image_h = image.size

        scale = min(
            (page_w - 2 * margin) / image_w,
            (page_h - 2 * margin) / image_h
        )

        draw_w = image_w * scale
        draw_h = image_h * scale

        x = (page_w - draw_w) / 2
        y = (page_h - draw_h) / 2

        c.drawImage(
            ImageReader(image),
            x, y,
            width=draw_w,
            height=draw_h,
            preserveAspectRatio=True
        )

        c.showPage()

    c.save()
    output.seek(0)

    return output.getvalue()


# --------------------------------------------------
# CERTIFICATE INPUTS
# --------------------------------------------------

st.header("1. Certificate Details")

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

selected_columns = []

if show_model:
    selected_columns.append(("model_no", "Model No."))

if show_serial:
    selected_columns.append(("serial_no", "Serial No."))

if show_make:
    selected_columns.append(("make", "Make"))


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
        if not str(value).strip()
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
            pdf_bytes = make_certificate_pdf({
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

            st.success("Certificate PDF generated successfully!")

        except Exception as exc:
            st.error(f"PDF generation failed: {exc}")


if "mcert_pdf" in st.session_state:
    st.download_button(
        "⬇️ DOWNLOAD CERTIFICATE PDF",
        data=st.session_state["mcert_pdf"],
        file_name=st.session_state.get(
            "mcert_filename",
            "test_certificate.pdf"
        ),
        mime="application/pdf",
        use_container_width=True
    )


# --------------------------------------------------
# SCANNING SECTION
# --------------------------------------------------

st.divider()
st.header("2. Scan Documents")

st.info(
    "Use Scan Now to capture a page with your phone camera, "
    "or upload a scan/image. This section does not activate "
    "the HP printer's physical scanner from Streamlit Cloud."
)

if "scan_pages" not in st.session_state:
    st.session_state["scan_pages"] = []

if "camera_slot" not in st.session_state:
    st.session_state["camera_slot"] = 0


st.subheader("Scan Now — Phone Camera")

camera_image = st.camera_input(
    "Aim your phone camera at the document",
    key=f"camera_scan_{st.session_state['camera_slot']}"
)

if st.button(
    "📷 SCAN NOW / ADD PAGE",
    type="primary",
    use_container_width=True
):
    if camera_image is None:
        st.warning(
            "First take a photo using the camera above, "
            "then press Scan Now / Add Page."
        )
    else:
        st.session_state["scan_pages"].append(
            camera_image.getvalue()
        )
        st.session_state["camera_slot"] += 1
        st.rerun()


st.subheader("Add Scanned Images")

uploaded_images = st.file_uploader(
    "Upload JPG or PNG scan images",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="scan_image_upload"
)

if st.button(
    "➕ ADD UPLOADED PAGES",
    use_container_width=True
):
    if not uploaded_images:
        st.warning("Please select one or more images first.")
    else:
        added = 0

        for uploaded in uploaded_images:
            try:
                image = normalize_image(uploaded.getvalue())
                buffer = io.BytesIO()
                image.save(buffer, format="JPEG", quality=92)

                st.session_state["scan_pages"].append(
                    buffer.getvalue()
                )

                added += 1

            except Exception as exc:
                st.warning(
                    f"Could not read {uploaded.name}: {exc}"
                )

        if added:
            st.success(
                f"Added {added} page(s). "
                f"Total pages: {len(st.session_state['scan_pages'])}"
            )
            st.rerun()


# --------------------------------------------------
# PREVIEW / DELETE / PDF
# --------------------------------------------------

st.divider()

pages = st.session_state["scan_pages"]

st.subheader(f"Scanned Images — {len(pages)} Page(s)")

if not pages:
    st.caption(
        "No pages yet. Capture a photo or upload scanned images."
    )

else:
    for index, page_bytes in enumerate(pages):
        with st.container(border=True):
            image_col, action_col = st.columns([4, 1])

            with image_col:
                st.markdown(f"**Page {index + 1}**")
                st.image(
                    page_bytes,
                    caption=f"Scanned Page {index + 1}",
                    use_container_width=True
                )

            with action_col:
                st.write("")
                st.write("")

                if st.button(
                    "🗑️ DELETE",
                    key=f"delete_page_{index}",
                    use_container_width=True
                ):
                    st.session_state["scan_pages"].pop(index)
                    st.rerun()

    try:
        scanned_pdf = make_scanned_pdf(
            st.session_state["scan_pages"]
        )

        st.download_button(
            "⬇️ DOWNLOAD ALL PAGES AS PDF",
            data=scanned_pdf,
            file_name=(
                "scanned_document_"
                + datetime.now().strftime("%Y%m%d_%H%M%S")
                + ".pdf"
            ),
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )

    except Exception as exc:
        st.error(f"Could not create scanned PDF: {exc}")


if st.session_state["scan_pages"]:
    if st.button(
        "🗑️ DELETE ALL PAGES",
        use_container_width=True
    ):
        st.session_state["scan_pages"] = []
        st.rerun()
