# streamlit_app.py

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import arabic_reshaper
from bidi.algorithm import get_display

# OCR
import pytesseract
from pdf2image import convert_from_path

# =========================
# Arabic Helpers
# =========================

def reshape_arabic_text(text):
    try:
        reshaped = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped)
        return bidi_text
    except:
        return text

# =========================
# Normalize OCR Text 🔥
# =========================

def normalize_text(text):
    text = text.replace("ـ", "")
    text = text.replace("٬", "").replace("٫", ".")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# =========================
# OCR FUNCTION
# =========================

def ocr_pdf(pdf_path):
    text_pages = []
    images = convert_from_path(pdf_path, dpi=300)

    for img in images:
        text = pytesseract.image_to_string(
            img,
            lang="ara+eng",
            config="--oem 3 --psm 6"
        )
        text_pages.append(text)

    return "\n".join(text_pages)

# =========================
# Smart Field Extraction 🔥
# =========================

def find_field(text, keyword_pattern):
    text = normalize_text(text)

    pattern = rf"{keyword_pattern}\s*[:\-]?\s*(.{{0,60}})"
    match = re.search(pattern, text)

    if match:
        value = match.group(1)

        # stop at next keyword-like word
        value = re.split(r"(رقم|تاريخ|الإجمالي|الضريبة|المبلغ)", value)[0]

        return value.strip()

    return ""

# =========================
# Metadata Extraction
# =========================

def extract_metadata(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        # OCR fallback
        if not full_text.strip() or len(full_text.strip()) < 50:
            full_text = ocr_pdf(pdf_path)

        metadata = {
            "Invoice Number": find_field(full_text, r"(رقم الفاتورة|فاتورة رقم|Invoice\s*No)"),
            "Invoice Date": find_field(full_text, r"(تاريخ الفاتورة|التاريخ|Date)"),
            "Customer Name": find_field(full_text, r"(اسم العميل|العميل|Customer)"),
            "Address": find_field(full_text, r"(العنوان)"),
            "Paid": find_field(full_text, r"(مدفوع|Paid)"),
            "Balance": find_field(full_text, r"(الإجمالي|Total)"),
            "Not Paid": find_field(full_text, r"(الرصيد المستحق|Remaining)"),
            "Source File": pdf_path.name
        }

        return metadata

    except Exception as e:
        st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
        return {}

# =========================
# Table Extraction
# =========================

def extract_tables(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
            found_tables = False

            for page in pdf.pages:
                tables = page.extract_tables()

                if tables:
                    found_tables = True

                for table in tables:
                    if table:
                        df = pd.DataFrame(table)
                        df = df.dropna(how="all").reset_index(drop=True)

                        if df.empty:
                            continue

                        all_data.append(df)

            # OCR fallback for tables
            if not found_tables:
                text = ocr_pdf(pdf_path)
                lines = text.split("\n")
                data = []

                for line in lines:
                    if re.search(r"\d", line):
                        parts = re.split(r"\s{2,}", line)
                        if len(parts) >= 3:
                            data.append(parts)

                if data:
                    return pd.DataFrame(data)

            return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Error extracting table from {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Main Process
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
# Streamlit UI
# =========================

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor (OCR Ready)")

uploaded_files = st.file_uploader("Upload PDF or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

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
                for pdf in temp_dir.glob("*.pdf"):
                    pdf_paths.append(pdf)
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

            # Clean numeric fields
            for col in ["Paid", "Balance", "Not Paid"]:
                if col in final_df.columns:
                    final_df[col] = (
                        final_df[col].astype(str)
                        .str.replace(r"[^\d.]", "", regex=True)
                        .replace("", None)
                        .astype(float)
                    )

            if "Total before tax" in final_df.columns:
                final_df["Total before tax"] = pd.to_numeric(final_df["Total before tax"], errors="coerce")
                final_df["VAT 15%"] = final_df["Total before tax"] * 0.15
                final_df["Total after tax"] = final_df["Total before tax"] * 1.15

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            st.success("✅ Done!")
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
