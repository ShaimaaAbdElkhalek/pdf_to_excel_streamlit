import streamlit as st
import tempfile
import re
from pdf2image import convert_from_path
import pytesseract


# -----------------------------
# CLEAN TEXT
# -----------------------------
def clean_text(text):
    # keep Arabic + English + numbers only
    text = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-\n]', ' ', text)

    # normalize spaces
    text = re.sub(r'[ \t]+', ' ', text)

    # remove empty junk lines
    lines = [l.strip() for l in text.split("\n") if len(l.strip()) > 1]

    return "\n".join(lines).strip()


# -----------------------------
# OCR FUNCTION (IMPROVED)
# -----------------------------
def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)

    full_text = []

    config = r'--oem 3 --psm 6 -l ara+eng'

    for img in images:
        text = pytesseract.image_to_string(img, config=config)

        text = clean_text(text)

        if text:
            full_text.append(text)

    return "\n\n".join(full_text)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="OCR Arabic + English", layout="wide")

st.title("📄 PDF OCR Cleaner (Arabic + English)")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    with st.spinner("Processing OCR..."):
        text = pdf_to_text(pdf_path)

    if not text.strip():
        st.error("No readable text found 😢")
    else:
        st.success("Done!")

        st.text_area("📜 Clean OCR Output", text, height=500)

        st.download_button(
            "📥 Download Text",
            text,
            file_name="ocr_output.txt",
            mime="text/plain"
        )
