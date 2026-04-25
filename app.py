import streamlit as st
import pdfplumber
import tempfile
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from io import BytesIO


# -----------------------------
# OCR + TEXT EXTRACTION
# -----------------------------
def extract_text(pdf_path):
    text = ""

    # 1. Try normal text extraction
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except:
        pass

    # 2. If empty → OCR
    if not text.strip():
        images = convert_from_path(pdf_path)
        for img in images:
            text += pytesseract.image_to_string(img) + "\n"

    return text


# -----------------------------
# PROCESS PDF
# -----------------------------
def process_pdf(file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        pdf_path = tmp.name

    text = extract_text(pdf_path)

    if not text.strip():
        return pd.DataFrame()

    lines = [l.strip() for l in text.split("\n") if l.strip()]
    data = [l.split() for l in lines]

    if not data:
        return pd.DataFrame()

    max_len = max(len(r) for r in data)
    data = [r + [""] * (max_len - len(r)) for r in data]

    return pd.DataFrame(data)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="PDF OCR to Excel", layout="wide")

st.title("📄 PDF → Excel Converter (OCR + Smart Mode)")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing..."):
        df = process_pdf(uploaded_file)

    if df.empty:
        st.error("No data found 😢 (PDF might be encrypted or image-only with poor OCR)")
    else:
        st.success("Done!")

        st.dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            "📥 Download Excel",
            data=output,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
