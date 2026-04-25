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
# SMART KEY-VALUE PARSER (FIXED PROPERLY)
# -----------------------------
def smart_extract(text):

    def find_value(keyword_patterns):
        for pattern in keyword_patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1).strip()
        return ""

    data = {
        "Invoice Number": find_value([
            r"الفاتورة\s*[:\-]?\s*(\d+)",
            r"رقم الفاتورة\s*(\d+)"
        ]),

        "Customer Name": find_value([
            r"اسم العميل\s*[:\-]?\s*([^\n]+)"
        ]),

        "VAT Number": find_value([
            r"الرقم الضريبي\s*[:\-]?\s*([0-9]+)"
        ]),

        "CR Number": find_value([
            r"رقم السجل\s*[:\-]?\s*([0-9\- ]+)"
        ]),

        "Phone": find_value([
            r"رقم الجوال\s*[:\-]?\s*([0-9]+)"
        ]),

        "Total": find_value([
            r"المجموع\s*([0-9,\.]+)"
        ]),

        "VAT Value": find_value([
            r"القيمة المضافة\s*([0-9,\.]+)"
        ]),

        "Grand Total": find_value([
            r"الإحمالي\s*[:\-]?\s*([0-9,\.]+)"
        ]),

        "IBAN": find_value([
            r"(SA[0-9A-Z]{20,})"
        ])
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
st.set_page_config(page_title="Robust Invoice AI", layout="wide")

st.title("📄 Robust Invoice AI Parser (Handles all formats)")

file = st.file_uploader("Upload PDF", type=["pdf"])

if file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        path = tmp.name

    with st.spinner("Processing invoice..."):

        raw_text = extract_text(path)
        cleaned_text = clean_text(raw_text)
        df = smart_extract(raw_text)

    # ---------------- TEXT ----------------
    st.subheader("📜 OCR Text")
    st.text_area("Raw OCR", cleaned_text, height=300)

    # ---------------- STRUCTURED DATA ----------------
    st.subheader("📊 Extracted Data (Robust)")
    st.dataframe(df)

    # ---------------- DOWNLOAD ----------------
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        "📥 Download Excel",
        data=output,
        file_name="invoice_structured.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
