import io
import re
import base64
from datetime import datetime

import streamlit as st
from PIL import Image
import fitz
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="M-CERTIFICATE", page_icon="📄", layout="wide")

DEFAULTS = {
    "app_heading": "M-CERTIFICATE",
    "app_subheading": "Certificate Generator",
    "certificate_heading": "TEST CERTIFICATE",
    "concern_heading": "TO WHOM IT MAY CONCERN",
    "date_label": "Date",
    "item_label": "ITEM",
    "model_label": "Model No.",
    "serial_label": "Serial No.",
    "make_label": "Make",
    "mac_label": "MAC ID",
    "win_label": "WIN KEY",
    "sold_to_label": "Sold To / Customer Name",
    "company_name": "OptoPixel Pvt. Ltd.",
    "authorized_text": "Authorized",
    "intro_prefix": "This test certificate is for",
    "intro_suffix": "as per below table",
    "test_line_1": "Specified component is tested for all its functionalities and is performing",
    "test_line_2": "satisfactorily and found OK.",
}

if "certificate_settings" not in st.session_state:
    st.session_state["certificate_settings"] = DEFAULTS.copy()
cfg = st.session_state["certificate_settings"]

st.markdown("""
<style>
.stApp { background:#0b1220; color:#eef2ff; }
.block-container { max-width:1150px; padding-top:1.5rem; }
h1,h2,h3,p,label { color:#eef2ff; }
</style>
""", unsafe_allow_html=True)

# Settings sidebar
with st.sidebar:
    st.header("⚙️ Certificate Settings")
    st.caption("Edit headings and company details.")
    with st.form("settings_form"):
        edited = {}
        groups = [
            ("App heading", "app_heading"),
            ("App subtitle", "app_subheading"),
            ("PDF certificate heading", "certificate_heading"),
            ("To whom heading", "concern_heading"),
            ("Date label", "date_label"),
            ("Item column heading", "item_label"),
            ("Model No. heading", "model_label"),
            ("Serial No. heading", "serial_label"),
            ("Make heading", "make_label"),
            ("MAC ID heading", "mac_label"),
            ("WIN KEY heading", "win_label"),
            ("Sold To label", "sold_to_label"),
            ("Company name", "company_name"),
            ("Signature text", "authorized_text"),
            ("Introduction start", "intro_prefix"),
            ("Introduction ending", "intro_suffix"),
            ("Test statement line 1", "test_line_1"),
            ("Test statement line 2", "test_line_2"),
        ]
        for label, key in groups:
            edited[key] = st.text_input(label, value=cfg[key], key="setting_" + key)
        save = st.form_submit_button("💾 Save Settings", use_container_width=True)
        if save:
            st.session_state["certificate_settings"] = {
                k: v.strip() for k, v in edited.items()
            }
            st.rerun()
    if st.button("↩️ Restore Defaults", use_container_width=True):
        st.session_state["certificate_settings"] = DEFAULTS.copy()
        st.rerun()

cfg = st.session_state["certificate_settings"]

def wrap_words(text, font, size, max_width):
    words = str(text).split()
    lines, line = [], ""
    for word in words:
        candidate = word if not line else line + " " + word
        if stringWidth(candidate, font, size) <= max_width:
            line = candidate
        else:
            if line:
                lines.append(line)
            if stringWidth(word, font, size) <= max_width:
                line = word
            else:
                chunk = ""
                for ch in word:
                    if stringWidth(chunk + ch, font, size) <= max_width:
                        chunk += ch
                    else:
                        if chunk:
                            lines.append(chunk)
                        chunk = ch
                line = chunk
    if line:
        lines.append(line)
    return lines or [""]

def safe_filename(name):
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", name.strip()).strip("_")
    return f"test_certificate_{cleaned or 'item'}.pdf"

def make_pdf(data, cfg):
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    W, H = A4
    left, right = 54, W - 54
    content_w = right - left
    c.setFillColorRGB(0.08, 0.08, 0.08)

    date_text = f'{cfg["date_label"]}: {data["date"].strftime("%d-%b-%y").upper()}'
    c.setFont("Helvetica", 11)
    c.drawRightString(right, H - 145, date_text)

    title = cfg["certificate_heading"] or "TEST CERTIFICATE"
    title_y = H - 95
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, title_y, title)
    tw = stringWidth(title, "Helvetica-Bold", 16)
    c.line(W / 2 - tw / 2, title_y - 3, W / 2 + tw / 2, title_y - 3)

    whom = cfg["concern_heading"]
    c.setFont("Helvetica", 12)
    whom_y = H - 190
    c.drawCentredString(W / 2, whom_y, whom)
    ww = stringWidth(whom, "Helvetica", 12)
    c.line(W / 2 - ww / 2, whom_y - 2, W / 2 + ww / 2, whom_y - 2)

    labels = [label for _, label in data["selected_columns"]]
    intro = f'{cfg["intro_prefix"]} {data["item_name"]}'
    if labels:
        intro += " having " + ", ".join(labels)
    intro += " " + cfg["intro_suffix"]
    intro_lines = wrap_words(intro, "Helvetica", 11, content_w)
    intro_y = H - 245
    c.setFont("Helvetica", 11)
    for i, line in enumerate(intro_lines):
        c.drawString(left, intro_y - i * 15, line)

    columns = [(cfg["item_label"], data["item_name"])]
    columns += [(label, data[key]) for key, label in data["selected_columns"]]
    n = len(columns)
    widths = [content_w] if n == 1 else [content_w * 0.27] + [content_w * 0.73 / (n - 1)] * (n - 1)
    pad, fs, gap = 6, 9.5, 12
    wrapped = []
    max_head_lines = max_value_lines = 1
    for i, (head, value) in enumerate(columns):
        avail = widths[i] - 2 * pad
        hl = wrap_words(head, "Helvetica-Bold", 9, avail)
        vl = wrap_words(value, "Helvetica", fs, avail)
        wrapped.append((hl, vl))
        max_head_lines = max(max_head_lines, len(hl))
        max_value_lines = max(max_value_lines, len(vl))

    head_h = max(30, max_head_lines * gap + 12)
    row_h = max(31, max_value_lines * gap + 12)
    table_top = intro_y - len(intro_lines) * 15 - 25
    table_bottom = table_top - head_h - row_h
    xs = [left]
    for width in widths:
        xs.append(xs[-1] + width)
    c.setLineWidth(0.6)
    c.line(left, table_top, right, table_top)
    c.line(left, table_top - head_h, right, table_top - head_h)
    c.line(left, table_bottom, right, table_bottom)
    for x in xs:
        c.line(x, table_top, x, table_bottom)

    for i, (hl, vl) in enumerate(wrapped):
        cx = (xs[i] + xs[i + 1]) / 2
        y = table_top - (head_h - len(hl) * gap) / 2 - 9
        c.setFont("Helvetica-Bold", 9)
        for line in hl:
            c.drawCentredString(cx, y, line)
            y -= gap
        y = table_top - head_h - (row_h - len(vl) * gap) / 2 - 8
        c.setFont("Helvetica", fs)
        for line in vl:
            c.drawCentredString(cx, y, line)
            y -= gap

    y = table_bottom - 35
    c.setFont("Helvetica", 10.5)
    for line in wrap_words(cfg["test_line_1"], "Helvetica", 10.5, content_w):
        c.drawString(left, y, line)
        y -= 14
    for line in wrap_words(cfg["test_line_2"], "Helvetica", 10.5, content_w):
        c.drawString(left, y, line)
        y -= 14
    y -= 28
    sold = f'The above said item is sold to {data["sold_to"]}.'
    c.setFont("Helvetica-Oblique", 10.5)
    for line in wrap_words(sold, "Helvetica-Oblique", 10.5, content_w):
        c.drawString(left, y, line)
        y -= 14
    y -= 30
    c.setFont("Helvetica", 10.5)
    c.drawString(left, y, f'For {cfg["company_name"]},')
    c.drawString(left, y - 48, cfg["authorized_text"])
    c.save()
    out.seek(0)
    return out.getvalue()

def pdf_preview(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
    return Image.open(io.BytesIO(pix.tobytes("png"))).copy()

st.title(cfg["app_heading"])
st.caption(cfg["app_subheading"])
st.header("Certificate Details")

left_col, right_col = st.columns(2)
with left_col:
    chosen_date = st.date_input(
        cfg["date_label"], value=None, format="DD/MM/YYYY",
        help="Choose a date. Initially, no date is selected."
    )
    item_name = st.text_input(cfg["item_label"] + " Name")
    st.markdown("**Select fields to show in table**")
    show_model = st.checkbox(cfg["model_label"])
    show_serial = st.checkbox(cfg["serial_label"])
    show_make = st.checkbox(cfg["make_label"])
    show_mac = st.checkbox(cfg["mac_label"])
    show_win = st.checkbox(cfg["win_label"])

with right_col:
    model_no = st.text_input(cfg["model_label"] + " Value") if show_model else ""
    serial_no = st.text_input(cfg["serial_label"] + " Value") if show_serial else ""
    make = st.text_input(cfg["make_label"] + " Value") if show_make else ""
    mac_id = st.text_input(cfg["mac_label"] + " Value") if show_mac else ""
    win_key = st.text_input(cfg["win_label"] + " Value") if show_win else ""
    sold_to = st.text_input(cfg["sold_to_label"])

selected = []
for enabled, key, label in [
    (show_model, "model_no", cfg["model_label"]),
    (show_serial, "serial_no", cfg["serial_label"]),
    (show_make, "make", cfg["make_label"]),
    (show_mac, "mac_id", cfg["mac_label"]),
    (show_win, "win_key", cfg["win_label"]),
]:
    if enabled:
        selected.append((key, label))

if st.button("📄 GENERATE CERTIFICATE", type="primary", use_container_width=True):
    required = {"Date": chosen_date, "Item Name": item_name, cfg["sold_to_label"]: sold_to}
    if show_model: required[cfg["model_label"]] = model_no
    if show_serial: required[cfg["serial_label"]] = serial_no
    if show_make: required[cfg["make_label"]] = make
    if show_mac: required[cfg["mac_label"]] = mac_id
    if show_win: required[cfg["win_label"]] = win_key
    missing = [label for label, value in required.items() if value is None or not str(value).strip()]
    if missing:
        st.error("Please fill in: " + ", ".join(missing))
    else:
        try:
            pdf = make_pdf({
                "date": chosen_date,
                "item_name": item_name.strip(),
                "model_no": model_no.strip(),
                "serial_no": serial_no.strip(),
                "make": make.strip(),
                "mac_id": mac_id.strip(),
                "win_key": win_key.strip(),
                "sold_to": sold_to.strip(),
                "selected_columns": selected,
            }, cfg)
            st.session_state["certificate_pdf"] = pdf
            st.session_state["certificate_filename"] = safe_filename(item_name)
            st.success("Certificate generated successfully!")
        except Exception as exc:
            st.error(f"Certificate generation failed: {exc}")

if "certificate_pdf" in st.session_state:
    st.subheader("Certificate Preview")
    try:
        st.image(pdf_preview(st.session_state["certificate_pdf"]), use_container_width=True)
    except Exception as exc:
        st.warning(f"Could not show preview: {exc}")
    st.download_button(
        "⬇️ DOWNLOAD CERTIFICATE PDF",
        data=st.session_state["certificate_pdf"],
        file_name=st.session_state.get("certificate_filename", "test_certificate.pdf"),
        mime="application/pdf",
        use_container_width=True,
    )
    st.caption("To print, download the PDF and choose Print from your PDF viewer.")
