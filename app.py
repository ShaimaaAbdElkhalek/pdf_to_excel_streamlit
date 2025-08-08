import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
import os
from pathlib import Path
from datetime import datetime

# =========================
# Helper Functions
# =========================

def extract_tables_from_pdf(pdf_path):
    all_tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df = pd.DataFrame(table)
                        if len(set(df.columns)) == 1 and df.columns[0] is None:
                            continue
                        if is_data_row(df.iloc[0]):
                            all_tables.append(df)
        return all_tables, None
    except Exception as e:
        return None, str(e)

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def clean_table(df):
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = [f"Col{i+1}" for i in range(len(df.columns))]
    return df

def display_error(filename, msg):
    st.error(f"❌ Error in {filename}: {msg}")

# =========================
# Streamlit App
# =========================

st.title("📄 Arabic Invoice Table Extractor (No Java Needed)")
st.markdown("Upload Arabic PDF invoices (single or zipped folder), and this app will extract tables using `pdfplumber` and export them to Excel.")

uploaded_file = st.file_uploader("Upload a PDF or ZIP file", type=["pdf", "zip"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)

        pdf_paths = []
        extracted_dir = temp_dir / "extracted"
        extracted_dir.mkdir(exist_ok=True)

        # Handle uploaded file
        if uploaded_file.name.endswith(".zip"):
            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                zip_ref.extractall(extracted_dir)
            for file in extracted_dir.rglob("*.pdf"):
                pdf_paths.append(file)
        else:
            file_path = extracted_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())
            pdf_paths.append(file_path)

        combined_data = []
        for pdf_path in pdf_paths:
            st.markdown(f"📄 Processing: {pdf_path.name}")
            tables, error = extract_tables_from_pdf(pdf_path)

            if error:
                display_error(pdf_path.name, error)
                continue

            if not tables:
                display_error(pdf_path.name, "⚠️ No valid tables found.")
                continue

            for df in tables:
                try:
                    cleaned = clean_table(df)
                    # Simulate 6-column validation
                    if len(cleaned.columns) != 6:
                        raise ValueError(f"6 columns passed, passed data had {len(cleaned.columns)} columns")
                    cleaned.insert(0, "Source File", pdf_path.name)
                    combined_data.append(cleaned)
                except Exception as e:
                    display_error(pdf_path.name, str(e))
                    break  # Skip to next file

        if combined_data:
            result_df = pd.concat(combined_data, ignore_index=True)
            output_path = temp_dir / "Cleaned_Tables.xlsx"
            result_df.to_excel(output_path, index=False)
            st.success("✅ Extraction completed successfully!")
            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Excel File", f, "Cleaned_Tables.xlsx")
        else:
            st.warning("⚠️ No valid data extracted from uploaded PDFs.")
