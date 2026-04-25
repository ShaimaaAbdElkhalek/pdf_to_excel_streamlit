import streamlit as st
import tempfile
import re
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from io import BytesIO


# -----------------------------
# OCR FUNCTION
# -----------------------------
def extract_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)
    config = r'--oem 3 --psm 6 -l ara+eng'

    text = ""

    for img in images:
        text += pytesseract.image_to_string(img, config=config) + "\n"

    return text


# -----------------------------
# SAFE FIND FUNCTION (FIXED)
# -----------------------------
def find(pattern, text):
    match = re.search(pattern, text)
    if not match:
        return ""

    # safe: handle groups or full match
    try:
        return match.group(1).strip()
    except:
        return match.group(0).strip()


# -----------------------------
# FIELD EXTRACTION (FIXED)
# -----------------------------
def extract_fields(text):

    data = {
        "Invoice Number": find(r"الفاتورة\s*[:\-]?\s*(\d+)", text),
        "Customer Name": find(r"اسم العميل\s*[:\-]?\s*([^\n]+)", text),
        "VAT Number": find(r"الرقم الضريبي\s*[:\-]?\s*([0-9]+)", text),
        "CR Number": find(r"رقم السجل\s*[:\-]?\s*([0-9\- ]+)", text),
        "Phone": find(r"رقم الجوال\s*[:\-]?\s*([0-9]+)", text),

        "Total": find(r"المجموع\s*([0-9,\.]+)", text),
        "VAT Value": find(r"القيمة المضافة\s*([0-9,\.]+)", text),
        "Grand Total": find(r"الإحمالي\s*([0-9,\.]+)", text),

        # FIXED IBAN (no crash)
        "IBAN": find(r"(SA[0-9A-Z]+)", text),
    }

    return pd.DataFrame([data])


# -----------------------------
# CLEAN TEXT
# -----------------------------
def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Invoice AI FIXED", layout="wide")

st.title("📄 Invoice AI Parser (FIXED VERSION)")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    with st.spinner("Processing OCR..."):

        raw_text = extract_text(pdf_path)
        cleaned_text = clean_text(raw_text)
        df = extract_fields(raw_text)

    # ---------------- TEXT ----------------
    st.subheader("📜 OCR Text")
    st.text_area("Raw OCR Output", cleaned_text, height=300)

    # ---------------- TABLE ----------------
    st.subheader("📊 Extracted Data")
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
