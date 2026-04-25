import streamlit as st
import tempfile
from pdf2image import convert_from_path
import pytesseract


# -----------------------------
# OCR FUNCTION (CLEAN OUTPUT)
# -----------------------------
def pdf_to_clean_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)

    full_text = []

    for img in images:
        # better OCR config for readable text
        custom_config = r'--oem 3 --psm 6'

        text = pytesseract.image_to_string(img, config=custom_config)

        # CLEANING STEP
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        clean_page = "\n".join(lines)

        full_text.append(clean_page)

    return "\n\n".join(full_text)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="PDF OCR Text Cleaner", layout="wide")

st.title("📄 PDF → Clean Readable Text (OCR)")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        path = tmp.name

    with st.spinner("Extracting clean text with OCR..."):
        text = pdf_to_clean_text(path)

    if not text.strip():
        st.error("No readable text found 😢")
    else:
        st.success("Done!")

        st.text_area("📜 Clean OCR Text", text, height=500)

        st.download_button(
            "📥 Download Text",
            text,
            file_name="ocr_text.txt",
            mime="text/plain"
        )
