import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
from pathlib import Path
import zipfile

# =========================
# Helper Functions
# =========================

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def extract_text_fields(text, fields_regex):
    extracted = {}
    for label, pattern in fields_regex.items():
        match = re.search(pattern, text)
        if match:
            extracted[label] = match.group(1).strip()
    return extracted

def process_pdf(pdf_path, fields_regex):
    all_tables = []
    extracted_fields = {}

    # Extract tables with pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                cleaned = [row for row in table if is_data_row(row)]
                if cleaned:
                    df = pd.DataFrame(cleaned)
                    all_tables.append(df)

    # Extract fields with fitz (PyMuPDF)
    with fitz.open(pdf_path) as doc:
        text = ""
        for page in doc:
            text += page.get_text()
        extracted_fields = extract_text_fields(text, fields_regex)

    return all_tables, extracted_fields

# =========================
# Streamlit App
# =========================

st.title("🧾 Arabic Invoice Extractor")

uploaded_files = st.file_uploader(
    "Upload one or more PDF files", type="pdf", accept_multiple_files=True
)

if uploaded_files:
    # Arabic field patterns (add more as needed)
    fields_regex = {
        "رقم الفاتورة": r"رقم الفاتورة[:\s\-]*([^\n\r]+)",
        "اسم العميل": r"اسم العميل[:\s\-]*([^\n\r]+)",
        "العنوان": r"العنوان[:\s\-]*([^\n\r]+)",
        "تاريخ الفاتورة": r"تاريخ الفاتورة[:\s\-]*([^\n\r]+)"
    }

    with tempfile.TemporaryDirectory() as tmpdirname:
        result_excel_path = Path(tmpdirname) / "cleaned_results.xlsx"
        writer = pd.ExcelWriter(result_excel_path, engine="openpyxl")

        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            st.write(f"📄 Processing: {file_name}")

            safe_pdf_path = Path(tmpdirname) / file_name
            with open(safe_pdf_path, "wb") as f:
                f.write(uploaded_file.read())

            try:
                tables, fields = process_pdf(safe_pdf_path, fields_regex)

                if tables:
                    for i, table in enumerate(tables):
                        sheet_name = f"{file_name[:25]}_T{i+1}"
                        for key, value in fields.items():
                            table[key] = value
                        table.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    st.warning(f"⚠️ No valid tables found in: {file_name}")
            except Exception as e:
                st.error(f"❌ Failed to process {file_name}: {e}")

        writer.close()

        # Create download link
        with open(result_excel_path, "rb") as f:
            st.download_button(
                label="📥 Download Cleaned Excel",
                data=f,
                file_name="Cleaned_Invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
