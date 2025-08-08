import streamlit as st
import os
import shutil
import pdfplumber
import pandas as pd
import tempfile
import zipfile
from pathlib import Path

# =========================
# Helper Functions
# =========================

def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

def extract_tables(pdf_path):
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                for table in page_tables:
                    if table:
                        df = pd.DataFrame(table)
                        if not df.empty:
                            tables.append(df)
    except Exception as e:
        st.error(f"❌ Error reading tables from {pdf_path.name}: {e}")
    return tables

def process_pdf(pdf_path, safe_folder):
    all_rows = []
    try:
        ascii_name = f"bill_{pdf_path.stem.encode('ascii', errors='ignore').decode()}.pdf"
        safe_pdf_path = safe_folder / ascii_name
        shutil.copy(pdf_path, safe_pdf_path)

        tables = extract_tables(safe_pdf_path)

        for table in tables:
            if not table.empty:
                merged_rows = []
                temp_row = []

                for _, row in table.iterrows():
                    row_values = row.fillna("").astype(str).tolist()

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
                    default_headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند", "إضافي"]
                    df_merged = pd.DataFrame(merged_rows, columns=default_headers[:num_cols])
                    df_merged["Source File"] = pdf_path.name
                    all_rows.append(df_merged)

    except Exception as e:
        st.error(f"❌ Error in {pdf_path.name}: {e}")
    return all_rows

# =========================
# Streamlit App UI
# =========================

st.title("📄 Arabic Invoice Table Extractor (Tables Only)")

uploaded_files = st.file_uploader("Upload PDF files or a ZIP of PDFs", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        # Unpack and handle ZIP or single/multiple PDFs
        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                for pdf in temp_dir.glob("*.pdf"):
                    pdf_paths.append(pdf)
            else:
                pdf_paths.append(file_path)

        # Safe temp folder
        safe_folder = temp_dir / "safe"
        safe_folder.mkdir(exist_ok=True)

        all_rows = []
        for pdf_path in pdf_paths:
            st.write(f"📄 Processing: {pdf_path.name}")
            extracted = process_pdf(pdf_path, safe_folder)
            all_rows.extend(extracted)

        if all_rows:
            final_df = pd.concat(all_rows, ignore_index=True)
            output_excel = temp_dir / "Extracted_Tables_Only.xlsx"
            final_df.to_excel(output_excel, index=False)

            st.success("✅ Table extraction complete!")
            st.download_button("📥 Download Tables Excel", output_excel.read_bytes(), file_name="Extracted_Tables.xlsx")
        else:
            st.warning("⚠️ No valid tables found.")
