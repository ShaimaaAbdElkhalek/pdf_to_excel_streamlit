import streamlit as st
import os
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path

# =========================
# Helper Functions
# =========================

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def extract_fields(text):
    fields = {
        "رقم الفاتورة": "",
        "التاريخ": "",
        "اسم العميل": "",
        "العنوان": "",
    }
    for field in fields:
        match = re.search(fr"{field}[:\s]*([^\n\r]+)", text)
        if match:
            fields[field] = match.group(1).strip()
    return fields

def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def merge_multiline_rows(table):
    cleaned = []
    buffer = None
    for row in table:
        if is_data_row(row):
            if buffer:
                cleaned.append(buffer)
                buffer = None
            cleaned.append(row)
        else:
            if buffer:
                buffer = [
                    f"{buffer[i]} {row[i]}" if i < len(row) else buffer[i] for i in range(len(buffer))
                ]
            else:
                buffer = row
    if buffer:
        cleaned.append(buffer)
    return cleaned

def extract_tables_from_pdf(pdf_path):
    all_dataframes = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                merged_table = merge_multiline_rows(table)
                df = pd.DataFrame(merged_table)
                df = df.dropna(how="all").dropna(axis=1, how="all")
                if 4 <= len(df.columns) <= 8:
                    all_dataframes.append(df)
    return all_dataframes

def process_pdf(pdf_path):
    metadata = extract_fields(extract_text_from_pdf(pdf_path))
    tables = extract_tables_from_pdf(pdf_path)
    final_data = []
    for df in tables:
        df.columns = [f"عمود {i+1}" for i in range(len(df.columns))]
        for key, value in metadata.items():
            df[key] = value
        final_data.append(df)
    return pd.concat(final_data, ignore_index=True) if final_data else pd.DataFrame()

# =========================
# Streamlit UI
# =========================

st.title("📄 Arabic PDF Invoice Extractor")

uploaded_files = st.file_uploader("Upload PDF invoices", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        results = []
        for uploaded_file in uploaded_files:
            file_path = os.path.join(tmpdir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())
            try:
                df = process_pdf(file_path)
                if not df.empty:
                    results.append((uploaded_file.name, df))
                else:
                    st.warning(f"⚠️ No valid tables found in {uploaded_file.name}")
            except Exception as e:
                st.error(f"❌ Error in {uploaded_file.name}: {str(e)}")

        if results:
            output_path = os.path.join(tmpdir, "invoices.xlsx")
            with pd.ExcelWriter(output_path) as writer:
                for name, df in results:
                    sheet_name = Path(name).stem[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Extracted Excel",
                    data=f,
                    file_name="invoices_cleaned.xlsx"
                )
