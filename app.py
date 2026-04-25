import streamlit as st
import tempfile
from pdf2image import convert_from_path
import pytesseract
import re

# Arabic fixing
import arabic_reshaper
from bidi.algorithm import get_display


# -----------------------------
# CLEAN TEXT
# -----------------------------
def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# -----------------------------
# FIX ARABIC RTL PROPERLY
# -----------------------------
def fix_arabic(text):
    try:
        reshaped = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped)
        return bidi_text
    except:
        return text


# -----------------------------
# OCR FUNCTION
# -----------------------------
def pdf_to_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=350)

    result = []

    for img in images:
        config = r'--oem 3 --psm 6 -l ara+eng'

        text = pytesseract.image_to_string(img, config=config)

        text = clean_text(text)
        text = fix_arabic(text)

        if text:
            result.append(text)

    return "\n\n".join(result)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Arabic OCR Fixed", layout="wide")

st.title("📄 OCR عربي + إنجليزي (مُحسن)")

file = st.file_uploader("Upload PDF", type=["pdf"])

if file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        path = tmp.name

    with st.spinner("Processing Arabic text correctly..."):
        text = pdf_to_text(path)

    st.success("Done!")

    st.text_area("📜 OCR Output (Fixed Arabic)", text, height=500)

    st.download_button(
        "📥 Download Text",
        text,
        file_name="arabic_ocr.txt",
        mime="text/plain"
    )
