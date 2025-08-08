# streamlit_app.py

import streamlit as st
import os
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
import arabic_reshaper
from bidi.algorithm import get_display
import shutil

# =========================
# Helper Functions
# =========================

def reshape_arabic_text(text):
    try:
        reshaped = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped)
        return bidi_text
    except:
        return text

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def fix_shifted_rows(row):
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def extract_fields_from_text(text):
    fields = {}
    patterns = {
        "رقم الفاتورة": r"رقم\s*الفاتورة[:\-]?\s*([\w\-\/]+)",
        "اسم العميل": r"اسم\s*العميل[:\-]?\s*(.+)",
        "العنوان": r"العنوان[:\-]?\s*(.+)",
        "رقم السجل التجاري": r"السجل\s*التجاري[:\-]?\s*([\d\-\/]+)",
        "الرقم الضريبي": r"الرقم\s*الضريبي[:\-]?\s*([\d\-\/]+)",
        "التاريخ": r"التاريخ[:\-]?\s*([\d\/\-]+)",
        "مدفوع": r"مدفوع[:\-]?\s*([\d.,]+)",
        "الرصيد المستحق": r"الرصيد\s*المستحق[:\-]?\s*([\d.,]+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            value = match.group(1).strip()
            fields[key] = reshape_arabic_text(value)
    return fields

def process_single_pdf(pdf_path, safe_folder):
    line_items = []
    with fitz.open(pdf_path) as doc:
        full_text = "\n".join([page.get_text() for page in doc])

    fields = extract_fields_from_text(full_text)
    source_file = pdf_path.name
    fields["Source File"] = source_file

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table).dropna(how="all").reset_index(drop=True)
                if df.empty:
                    continue
                merged_rows = []
                temp_row = []
                for _, row in df.iterrows():
                    row_values = row.fillna("").astype(str).tolist()
                    row_values = [reshape_arabic_text(cell) for cell in row_values]
                    row_values = fix_shifted_rows(row_values)

                    if is_data_row(row_values):
                        if temp_row:
                            combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
                            merged_rows.append(combined)
                            temp_row = []
                        else:
                            merged_rows.append(row_values)
                    else:
                        temp_row = row_values

                if merged_rows:
                    headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند", "إضافي"]
                    num_cols = len(merged_rows[0])
                    df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                    for key, val in fields.items():
                        df_cleaned[key] = val
                    line_items.append(df_cleaned)

    if line_items:
        return pd.concat(line_items, ignore_index=True)
    else:
        # If no tables found, return fields as 1 row DataFrame
        return pd.DataFrame([fields])

# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="📄 Arabic Invoice Processor", layout="wide")
st.title("📄 Arabic Invoice Table + Metadata Extractor")

uploaded_files = st.file_uploader("Upload PDFs or ZIPs", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        pdf_paths = []

        for uploaded_file in uploaded_files:
            file_path = tmp_path / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            if file_path.suffix == ".zip":
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(tmp_path)
                pdf_paths.extend(list(tmp_path.glob("*.pdf")))
            elif file_path.suffix == ".pdf":
                pdf_paths.append(file_path)

        safe_folder = tmp_path / "safe"
        safe_folder.mkdir(exist_ok=True)

        all_data = []
        error_files = []

        for pdf_path in pdf_paths:
            st.write(f"🔍 Processing: {pdf_path.name}")
            try:
                df = process_single_pdf(pdf_path, safe_folder)
                if not df.empty:
                    all_data.append(df)
                else:
                    error_files.append(pdf_path.name)
            except Exception as e:
                st.error(f"❌ Error in {pdf_path.name}: {e}")
                error_files.append(pdf_path.name)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.success("✅ All files processed successfully!")
            st.dataframe(final_df)

            # Download
            output_excel = tmp_path / "Merged_Extracted_Invoices.xlsx"
            final_df.to_excel(output_excel, index=False)
            st.download_button("📥 Download Merged Excel", output_excel.read_bytes(), file_name="Invoices_Combined.xlsx")

        if error_files:
            st.warning("⚠️ Issues occurred with:")
            for ef in error_files:
                st.markdown(f"- {ef}")
