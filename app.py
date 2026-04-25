import streamlit as st
import tempfile
import re
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from io import BytesIO


# -----------------------------
# OCR
# -----------------------------
def extract_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)
    config = r'--oem 3 --psm 6 -l ara+eng'

    text = ""

    for img in images:
        text += pytesseract.image_to_string(img, config=config) + "\n"

    return text


# -----------------------------
# SMART FIELD EXTRACTION
# -----------------------------
def extract_fields(text):

    def find(pattern):
        match = re.search(pattern, text)
        return match.group(1).strip() if match else ""

    data = {
        "Invoice Number": find(r"الفاتورة\s*[:\-]?\s*(\d+)"),
        "Customer Name": find(r"اسم العميل\s*[:\-]?\s*([^\n]+)"),
        "VAT Number": find(r"الرقم الضريبي\s*[:\-]?\s*([0-9]+)"),
        "CR Number": find(r"رقم السجل\s*[:\-]?\s*([0-9\- ]+)"),
        "Phone": find(r"رقم الجوال\s*[:\-]?\s*([0-9]+)"),
        "Total": find(r"المجموع\s*([0-9,\.]+)"),
        "VAT Value": find(r"القيمة المضافة\s*([0-9,\.]+)"),
        "Grand Total": find(r"الإحمالي\s*([0-9,\.]+)"),
        "IBAN": find(r"SA[0-9A-Z]+"),
    }

    return pd.DataFrame([data])


# -----------------------------
# CLEAN TEXT VIEW
# -----------------------------
def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Invoice AI Parser", layout="wide")

st.title("📄 Invoice AI Parser (Fix for messy OCR)")

file = st.file_uploader("Upload PDF", type=["pdf"])

if file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        path = tmp.name

    with st.spinner("Reading invoice..."):

        raw_text = extract_text(path)
        clean = clean_text(raw_text)
        df = extract_fields(raw_text)

    # ---------------- TEXT ----------------
    st.subheader("📜 OCR Text")
    st.text_area("Raw Text", clean, height=300)

    # ---------------- STRUCTURED DATA ----------------
    st.subheader("📊 Extracted Invoice Data")
    st.dataframe(df)

    # ---------------- DOWNLOAD ----------------
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        "📥 Download Excel",
        data=output,
        file_name="invoice.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
