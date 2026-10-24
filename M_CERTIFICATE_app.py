import io
import os
import re
from datetime import datetime

import streamlit as st
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth


st.set_page_config(
    page_title="M-CERTIFICATE",
    page_icon="📄",
    layout="wide",
)

st.markdown("""
<style>
.stApp { background:#0b1220; color:#eef2ff; }
.block-container { max-width:1200px; padding-top:1rem; }
h1,h2,h3,label,p { color:#eef2ff !important; }
.scan-card {
    border:1px solid #334155; border-radius:12px;
    padding:14px; margin-bottom:10px; background:#111c2e;
}
</style>
""", unsafe_allow_html=True)

st.title("M-CERTIFICATE")
st.caption("Certificate + multi-page document scanning • Cloud/browser compatible")


# ============================================================
# HELPERS
# ============================================================
def safe_filename(name, default="document"):
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name).strip())
    name = name.strip("._-")
    return name or default


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


def fit_text(c, text, x, y, max_width, font="Helvetica",
             size=10, min_size=7):
    current = size
    while current > min_size and stringWidth(text, font, current) > max_width:
        current -= 0.25
    c.setFont(font, current)
    c.drawString(x, y, text)


def parse_date(value):
    try:
        return datetime.strptime(value.strip(), "%d-%b-%y").date()
    except ValueError:
        return None


# ============================================================
# CERTIFICATE PDF
# ============================================================
def make_certificate_pdf(data):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    W, H = A4
    c.setFillColorRGB(0.08, 0.08, 0.08)

    fit_text(c, "Date: " + data["date"].strftime("%d-%b-%y"),
             428, H - 204, 100, size=12, min_size=10)

    title = "TEST CERTIFICATE"
    title_y = H - 142
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)
    tw = stringWidth(title, "Helvetica-Bold", 16)
    c.setLineWidth(0.8)
    c.line(W/2-tw/2, title_y-3, W/2+tw/2, title_y-3)

    whom = "TO WHOM IT MAY CONCERN"
    whom_y = H - 262
    c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, whom_y, whom)
    ww = stringWidth(whom, "Helvetica", 12)
    c.setLineWidth(0.55)
    c.line(W/2-ww/2, whom_y-2.5, W/2+ww/2, whom_y-2.5)

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
    c.setFont("Helvetica", 12)
    for i, line in enumerate(wrap_words(intro, "Helvetica", 12, 470)[:2]):
        c.drawString(72, intro_y - i * 17, line)

    x_left, x_right = 72, 550
    table_top = H - 359
    header_h, row_h = 31, 31
    table_bottom = table_top - header_h - row_h

    columns = [("item_name", "ITEM", data["item_name"])]
    for key, label in selected:
        columns.append((key, label, data[key]))

    col_width = (x_right - x_left) / len(columns)
    xp = [x_left + i * col_width for i in range(len(columns) + 1)]

    c.setLineWidth(0.55)
    for x in xp:
        c.line(x, table_top, x, table_bottom)
    c.line(x_left, table_top, x_right, table_top)
    c.line(x_left, table_top-header_h, x_right, table_top-header_h)
    c.line(x_left, table_bottom, x_right, table_bottom)

    c.setFont("Helvetica-Bold", 11)
    for i, (_, label, _) in enumerate(columns):
        c.drawCentredString((xp[i]+xp[i+1])/2, table_top-20, label)

    for i, (_, _, value) in enumerate(columns):
        left, right = xp[i], xp[i+1]
        txt = str(value)
        size = 10.5
        while size > 7 and stringWidth(txt, "Helvetica", size) > right-left-10:
            size -= 0.25
        c.setFont("Helvetica", size)
        c.drawCentredString((left+right)/2, table_bottom+9.5, txt)

    body_y = H - 452
    c.setFont("Helvetica", 11.8)
    c.drawString(72, body_y, "Specified component is tested for all its functionalities and is performing")
    c.drawString(72, body_y-18, "satisfactorily and found OK.")

    sold = f'The above said item is sold to {data["sold_to"]}.'
    c.setFont("Helvetica-Oblique", 11.8)
    for i, line in enumerate(wrap_words(sold, "Helvetica-Oblique", 11.8, 470)[:2]):
        c.drawString(72, H-528-i*17, line)

    c.setFont("Helvetica", 11.8)
    c.drawString(72, H-585, "For OptoPixel Pvt. Ltd.,")
    c.drawString(72, H-640, "Authorized")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# SCANNED IMAGE/PDF -> ONE PDF
# ============================================================
def image_to_pdf_page(image_bytes):
    image = Image.open(io.BytesIO(image_bytes))
    image = ImageOps.exif_transpose(image).convert("RGB")

    # A4 portrait at 150 DPI
    page_w, page_h = 1240, 1754
    image.thumbnail((page_w, page_h), Image.Resampling.LANCZOS)

    page = Image.new("RGB", (page_w, page_h), "white")
    x = (page_w - image.width) // 2
    y = (page_h - image.height) // 2
    page.paste(image, (x, y))

    out = io.BytesIO()
    page.save(out, format="PDF", resolution=150.0)
    return out.getvalue()


def combine_scan_files(files):
    writer = PdfWriter()

    for file in files:
        data = file["data"]
        kind = file["kind"]

        if kind == "pdf":
            reader = PdfReader(io.BytesIO(data))
            for page in reader.pages:
                writer.add_page(page)
        else:
            pdf_data = image_to_pdf_page(data)
            reader = PdfReader(io.BytesIO(pdf_data))
            writer.add_page(reader.pages[0])

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ============================================================
# SESSION STATE
# ============================================================
if "scan_pages" not in st.session_state:
    st.session_state.scan_pages = []

if "camera_counter" not in st.session_state:
    st.session_state.camera_counter = 0


# ============================================================
# MAIN UI
# ============================================================
cert_col, scan_col = st.columns([1, 1], gap="large")

with cert_col:
    st.subheader("📄 Certificate")

    date_text = st.text_input(
        "Date",
        placeholder="DD-MMM-YY (e.g. 23-Jun-26)",
        key="cert_date",
    )
    item_name = st.text_input(
        "Item Name",
        placeholder="Enter item name",
        key="cert_item",
    )

    st.markdown("**Select fields to show in table**")
    c1, c2, c3 = st.columns(3)
    with c1:
        show_model = st.checkbox("Model No.", key="show_model")
    with c2:
        show_serial = st.checkbox("Serial No.", key="show_serial")
    with c3:
        show_make = st.checkbox("Make", key="show_make")

    model_no = st.text_input("Model No. Value", key="model_value") if show_model else ""
    serial_no = st.text_input("Serial No. Value", key="serial_value") if show_serial else ""
    make = st.text_input("Make Value", key="make_value") if show_make else ""

    sold_to = st.text_input(
        "Sold To",
        placeholder="Enter customer name",
        key="sold_to",
    )

    selected_columns = []
    if show_model:
        selected_columns.append(("model_no", "Model No."))
    if show_serial:
        selected_columns.append(("serial_no", "Serial No."))
    if show_make:
        selected_columns.append(("make", "Make"))

    if st.button("📄 GENERATE CERTIFICATE", type="primary",
                 use_container_width=True):
        parsed = parse_date(date_text)
        required = {"Date": date_text, "Item Name": item_name, "Sold To": sold_to}
        if show_model:
            required["Model No."] = model_no
        if show_serial:
            required["Serial No."] = serial_no
        if show_make:
            required["Make"] = make

        missing = [k for k, v in required.items() if not str(v).strip()]

        if parsed is None:
            st.error("Date must be DD-MMM-YY, e.g. 23-Jun-26.")
        elif missing:
            st.error("Please fill in: " + ", ".join(missing))
        else:
            pdf = make_certificate_pdf({
                "date": parsed,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected_columns,
            })
            st.session_state.cert_pdf = pdf
            st.session_state.cert_filename = (
                f"test_certificate_{safe_filename(item_name, 'item')}.pdf"
            )
            st.success("Certificate generated.")

    if "cert_pdf" in st.session_state:
        st.download_button(
            "⬇️ DOWNLOAD CERTIFICATE PDF",
            st.session_state.cert_pdf,
            file_name=st.session_state.cert_filename,
            mime="application/pdf",
            use_container_width=True,
        )


with scan_col:
    st.subheader("📷 Scanner / Document Scan")
    st.info(
        "Cloud-safe scanning: use your PC/phone camera or upload pages from "
        "your scanner software. Use ➕ SCAN NEXT for multiple pages."
    )

    scan_name = st.text_input(
        "PDF File Name",
        placeholder="Enter file name, e.g. customer_document",
        key="scan_filename",
    )

    scan_tab, upload_tab = st.tabs(["📷 Camera Scan", "📁 Scanner Upload"])

    with scan_tab:
        camera = st.camera_input(
            "Take / scan page",
            key=f"camera_{st.session_state.camera_counter}",
        )

        if camera is not None:
            raw = camera.getvalue()
            signature = str(len(raw)) + "_" + str(hash(raw))

            if not st.session_state.scan_pages or \
               st.session_state.scan_pages[-1].get("signature") != signature:
                st.session_state.scan_pages.append({
                    "data": raw,
                    "kind": "image",
                    "signature": signature,
                })
                st.success(
                    f"Page {len(st.session_state.scan_pages)} added."
                )

        if st.button("➕ SCAN NEXT", use_container_width=True):
            st.session_state.camera_counter += 1
            st.rerun()

    with upload_tab:
        uploads = st.file_uploader(
            "Upload scanned pages (JPG/PNG/PDF)",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
            key="scan_uploads",
        )

        if uploads and st.button("➕ ADD UPLOADED PAGES",
                                 use_container_width=True):
            for f in uploads:
                kind = "pdf" if f.type == "application/pdf" else "image"
                st.session_state.scan_pages.append({
                    "data": f.getvalue(),
                    "kind": kind,
                    "signature": f"{f.name}_{len(f.getvalue())}",
                })
            st.success(f"{len(uploads)} page/file(s) added.")

    st.markdown("### Scanned Pages")

    if not st.session_state.scan_pages:
        st.caption("No pages scanned yet.")
    else:
        for i, page in enumerate(st.session_state.scan_pages, 1):
            a, b = st.columns([5, 1])
            with a:
                st.write(
                    f"**Page {i}** — "
                    f"{'PDF' if page['kind']=='pdf' else 'Image'}"
                )
            with b:
                if st.button("🗑️", key=f"del_{i}"):
                    st.session_state.scan_pages.pop(i-1)
                    st.rerun()

        b1, b2 = st.columns(2)

        with b1:
            if st.button("📑 CREATE ONE PDF", type="primary",
                         use_container_width=True):
                if not scan_name.strip():
                    st.error("Please enter a PDF file name.")
                else:
                    try:
                        combined = combine_scan_files(
                            st.session_state.scan_pages
                        )
                        st.session_state.scan_pdf = combined
                        st.session_state.scan_pdf_name = (
                            safe_filename(scan_name, "scanned_document")
                            + ".pdf"
                        )
                        st.success(
                            f"{len(st.session_state.scan_pages)} page(s) "
                            "combined into one PDF."
                        )
                    except Exception as e:
                        st.error(f"PDF creation failed: {e}")

        with b2:
            if st.button("🗑️ CLEAR ALL SCANS",
                         use_container_width=True):
                st.session_state.scan_pages = []
                st.session_state.pop("scan_pdf", None)
                st.rerun()

        if "scan_pdf" in st.session_state:
            st.download_button(
                "⬇️ DOWNLOAD SCANNED PDF",
                st.session_state.scan_pdf,
                file_name=st.session_state.scan_pdf_name,
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )
            st.caption(
                "For HP Wi-Fi printing from cloud, open the downloaded PDF "
                "and use your device/browser Print dialog."
            )
