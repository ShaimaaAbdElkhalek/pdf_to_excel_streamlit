# streamlit_app.py

import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

# OCR
import pytesseract
from pdf2image import convert_from_path

# =========================
# OCR FUNCTION (FIXED 🔥)
# =========================

def ocr_pdf(pdf_path):
    text_pages = []
    images = convert_from_path(pdf_path, dpi=400)  # higher quality

    for img in images:
        text = pytesseract.image_to_string(
            img,
            lang="ara",  # 🔥 IMPORTANT: Arabic only
            config="--oem 3 --psm 4"
        )
        text_pages.append(text)

    return "\n".join(text_pages)

# =========================
# TEXT CLEANING
# =========================

def normalize_text(text):
    text = text.replace("ـ", "")
    text = text.replace("٬", "").replace("٫", ".")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# =========================
# SMART FIELD EXTRACTION
# =========================

def extract_invoice_number(text):
    # Try normal Arabic label
    match = re.search(r"(رقم.*?)(\d{4,})", text)
    if match:
        return match.group(2)

    # fallback: first long number
    match = re.search(r"\b\d{4,}\b", text)
    return match.group(0) if match else ""


def extract_customer(text):
    # Look for مؤسسة or شركة
    match = re.search(r"(مؤسسة\s+[^\n]+)", text)
    if match:
        return match.group(1)

    match = re.search(r"(شركة\s+[^\n]+)", text)
    return match.group(1) if match else ""


def extract_address(text):
    match = re.search(r"(حي\s+[^\n]+)", text)
    return match.group(1) if match else ""


# =========================
# METADATA EXTRACTION
# =========================

def extract_metadata(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            text = "\n".join([page.get_text() for page in doc])

        # OCR fallback
        if not text.strip() or len(text.strip()) < 50:
            text = ocr_pdf(pdf_path)

        text = normalize_text(text)

        # 🔥 DEBUG (see OCR result)
        st.text_area("🔍 OCR TEXT", text[:2000], height=200)

        metadata = {
            "Invoice Number": extract_invoice_number(text),
            "Customer Name": extract_customer(text),
            "Address": extract_address(text),
            "Invoice Date": "",
            "Paid": "",
            "Balance": "",
            "Not Paid": "",
            "Source File": pdf_path.name
        }

        return metadata

    except Exception as e:
        st.error(f"❌ Metadata error: {e}")
        return {}

# =========================
# TABLE EXTRACTION (OCR BASED)
# =========================

def extract_tables(pdf_path):
    try:
        text = ocr_pdf(pdf_path)
        lines = text.split("\n")

        data = []

        for line in lines:
            if re.search(r"\d", line):
                parts = re.split(r"\s{2,}", line)
                if len(parts) >= 3:
                    data.append(parts)

        return pd.DataFrame(data) if data else pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Table error: {e}")
        return pd.DataFrame()

# =========================
# MAIN PROCESS
# =========================

def process_pdf(pdf_path):
    metadata = extract_metadata(pdf_path)
    table_data = extract_tables(pdf_path)

    if not table_data.empty:
        for key, value in metadata.items():
            table_data[key] = value
        return table_data
    else:
        return pd.DataFrame([metadata])

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Arabic Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor (Stable Version)")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name

            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                pdf_paths.extend(temp_dir.glob("*.pdf"))
            else:
                pdf_paths.append(file_path)

        all_data = []

        for pdf_path in pdf_paths:
            st.write(f"📄 Processing: {pdf_path.name}")
            df = process_pdf(pdf_path)

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ Extraction complete!")
            st.dataframe(final_df)

            output = BytesIO()
            final_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "📥 Download Excel",
                data=output,
                file_name="Invoices.xlsx"
            )

        else:
            st.warning("⚠️ No data extracted.")
