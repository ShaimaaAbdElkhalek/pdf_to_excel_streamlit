import streamlit as st
import tempfile
from pdf2image import convert_from_path
import pytesseract
import re


# -----------------------------
# CLEAN TEXT FUNCTION
# -----------------------------
def clean_text(text):
    # remove weird OCR artifacts
    text = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-]', ' ', text)

    # fix multiple spaces
    text = re.sub(r'\s+', ' ', text)

    return text.strip()


# -----------------------------
# OCR FUNCTION (AR + EN)
# -----------------------------
def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=350)

    full_text = []

    for img in images:
        # VERY IMPORTANT OCR CONFIG
        config = r'--oem 3 --psm 6 -l ara+eng'

        text = pytesseract.image_to_string(img, config=config)

        text = clean_text(text)

        if text:
            full_text.append(text)

    return "\n\n".join(full_text)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Arabic + English OCR", layout="wide")

st.title("📄 Clean Arabic + English PDF OCR")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        path = tmp.name

    with st.spinner("Reading Arabic + English text..."):
        text = pdf_to_text(path)

    if not text.strip():
        st.error("No readable text found 😢")
    else:
        st.success("Done!")

        st.text_area("📜 Clean OCR Output", text, height=500)

        st.download_button(
            "📥 Download Text",
            text,
            file_name="clean_ocr.txt",
            mime="text/plain"
        )

