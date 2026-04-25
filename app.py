import streamlit as st
import tempfile
import re
from pdf2image import convert_from_path
import pytesseract


# -----------------------------
# SMART CLEANING (IMPORTANT)
# -----------------------------
def smart_clean(lines):
    cleaned = []

    for line in lines:
        line = line.strip()

        if not line:
            continue

        # remove pure garbage lines (numbers only / single chars)
        if re.fullmatch(r'[\d\s\-\.:,()%]+', line):
            continue

        # remove OCR noise fragments
        if len(line) < 3:
            continue

        # keep Arabic + English + numbers only
        line = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-\+]', ' ', line)

        # normalize spaces
        line = re.sub(r'\s+', ' ', line).strip()

        cleaned.append(line)

    return cleaned


# -----------------------------
# OCR FUNCTION
# -----------------------------
def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)

    all_pages = []

    config = r'--oem 3 --psm 6 -l ara+eng'

    for img in images:
        text = pytesseract.image_to_string(img, config=config)

        lines = text.split("\n")

        clean_lines = smart_clean(lines)

        if clean_lines:
            all_pages.append("\n".join(clean_lines))

    return "\n\n".join(all_pages)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Smart Invoice OCR", layout="wide")

st.title("📄 Smart Invoice OCR (Arabic + English Clean)")

uploaded_file = st.file_uploader("Upload PDF Invoice", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        path = tmp.name

    with st.spinner("Reading invoice..."):

        text = pdf_to_text(path)

    if not text.strip():
        st.error("No readable text found 😢")
    else:
        st.success("Clean invoice extracted ✅")

        st.text_area("📜 Clean Invoice Text", text, height=500)

        st.download_button(
            "📥 Download Clean Text",
            text,
            file_name="clean_invoice.txt",
            mime="text/plain"
        )
        
