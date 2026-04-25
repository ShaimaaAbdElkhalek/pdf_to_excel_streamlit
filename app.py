# streamlit_app.py

import streamlit as st
import fitz
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

import pytesseract
from pdf2image import convert_from_path

# =========================
# OCR FUNCTION
# =========================

def ocr_pdf(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)
    text = ""

    for img in images:
        text += pytesseract.image_to_string(
            img,
            lang="ara",
            config="--oem 3 --psm 6"
        ) + "\n"

    return text

# =========================
# CLEAN TEXT
# =========================

def clean_text(text):
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# SMART EXTRACTION 🔥
# =========================

def extract_data(text):

    data = {}

    # 🔥 Invoice Number
    inv = re.search(r"رقم الفاتورة\s*[:\-]?\s*(\d+)", text)
    if not inv:
        inv = re.search(r"\b0?\d{4,5}\b", text)
    data["Invoice Number"] = inv.group(1) if inv else ""

    # 🔥 Date
    date = re.search(r"\d{2}/\d{2}/\d{4}", text)
    data["Invoice Date"] = date.group(0) if date else ""

    # 🔥 Customer Name
    cust = re.search(r"مؤسسة\s+[^\d]+", text)
    if not cust:
        cust = re.search(r"شركة\s+[^\d]+", text)
    data["Customer Name"] = cust.group(0).strip() if cust else ""

    # 🔥 Address
    addr = re.search(r"(حي\s+[^\d]+جدة)", text)
    data["Address"] = addr.group(0) if addr else ""

    # 🔥 Tax Number (very reliable)
    tax = re.search(r"\b3\d{14}\b", text)
    data["Tax Number"] = tax.group(0) if tax else ""

    # 🔥 Paid / Balance
    paid = re.search(r"مدفوع\s*(\d+)", text)
    data["Paid"] = paid.group(1) if paid else "0"

    balance = re.search(r"الرصيد المستحق\s*(\d+)", text)
    data["Balance"] = balance.group(1) if balance else "0"

    return data

# =========================
# PROCESS
# =========================

def process_pdf(pdf_path):
    text = ocr_pdf(pdf_path)
    text = clean_text(text)

    # 👇 DEBUG (important)
    st.text_area("🔍 OCR TEXT", text[:1500], height=200)

    data = extract_data(text)
    data["Source File"] = pdf_path.name

    return pd.DataFrame([data])

# =========================
# UI
# =========================

st.set_page_config(layout="wide")
st.title("📄 Arabic Invoice Extractor (Accurate Version)")

files = st.file_uploader("Upload PDF or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

if files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdfs = []

        for f in files:
            path = temp_dir / f.name
            with open(path, "wb") as file:
                file.write(f.read())

            if f.name.endswith(".zip"):
                with zipfile.ZipFile(path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                pdfs.extend(temp_dir.glob("*.pdf"))
            else:
                pdfs.append(path)

        all_data = []

        for pdf in pdfs:
            st.write(f"📄 Processing: {pdf.name}")
            df = process_pdf(pdf)
            all_data.append(df)

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Extraction complete")
        st.dataframe(final_df)

        output = BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button("📥 Download Excel", data=output, file_name="invoices.xlsx")
