import streamlit as st
import os
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path

# ======================
# Helper Functions
# ======================

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def extract_fields_from_text(text):
    fields = {}
    patterns = {
        "رقم الفاتورة": r"رقم\s*الفاتورة[:\-]?\s*([\w\-\/]+)",
        "اسم العميل": r"اسم\s*العميل[:\-]?\s*(.+)",
        "العنوان": r"العنوان[:\-]?\s*(.+)",
        "رقم السجل التجاري": r"السجل\s*التجاري[:\-]?\s*([\d\-\/]+)",
        "الرقم الضريبي": r"الرقم\s*الضريبي[:\-]?\s*([\d\-\/]+)",
        "التاريخ": r"التاريخ[:\-]?\s*([\d\/\-]+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            fields[key] = match.group(1).strip()
    return fields

def fix_shifted_rows(row):
    # Fix case when "العدد" is empty but next column contains data
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        # Merge columns 4 and 5 into 4, shift left
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def process_pdf(pdf_path):
    all_tables = []
    all_fields = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
        fields = extract_fields_from_text(full_text)

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    df = pd.DataFrame(table)
                    df = df.dropna(how="all").reset_index(drop=True)
                    if df.empty:
                        continue

                    merged_rows = []
                    temp_row = []

                    for _, row in df.iterrows():
                        row_values = row.fillna("").astype(str).tolist()
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
                        num_cols = len(merged_rows[0])
                        headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند", "إضافي"]
                        df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                        for key, value in fields.items():
                            df_cleaned[key] = value
                        all_tables.append(df_cleaned)
                        all_fields.append(fields)

    return pd.concat(all_tables, ignore_index=True) if all_tables else pd.DataFrame()

def save_to_excel(dataframes, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframes.to_excel(writer, index=False, sheet_name="Extracted Data")

# ======================
# Streamlit Interface
# ======================

st.set_page_config(page_title="Arabic PDF Invoice Extractor", layout="wide")
st.title("📄 Arabic PDF Invoice Extractor")

uploaded_files = st.file_uploader("Upload Arabic PDF invoices", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        all_data = []
        error_files = []

        for uploaded_file in uploaded_files:
            try:
                file_path = os.path.join(tmpdir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.read())

                st.write(f"✅ Processing: {uploaded_file.name}")
                extracted_df = process_pdf(file_path)
                if not extracted_df.empty:
                    all_data.append(extracted_df)
                else:
                    st.warning(f"⚠️ No valid tables found in: {uploaded_file.name}")
                    error_files.append(uploaded_file.name)

            except Exception as e:
                st.error(f"❌ Error in {uploaded_file.name}: {str(e)}")
                error_files.append(uploaded_file.name)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.success("✅ Extraction complete.")
            st.dataframe(final_df)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                save_to_excel(final_df, tmp.name)
                st.download_button("📥 Download Excel File", tmp.name, file_name="cleaned_invoices.xlsx")

        if error_files:
            st.warning("⚠️ Some files had issues:")
            for ef in error_files:
                st.markdown(f"- {ef}")
