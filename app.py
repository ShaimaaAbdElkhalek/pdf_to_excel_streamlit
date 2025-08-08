import streamlit as st
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

# Keywords to search for Arabic fields
FIELD_KEYWORDS = {
    "رقم الفاتورة": "Invoice Number",
    "اسم العميل": "Customer Name",
    "التاريخ": "Date",
    "العنوان": "Address"
}

def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def extract_fields_from_text(text):
    fields = {}
    for key_ar, key_en in FIELD_KEYWORDS.items():
        match = re.search(fr"{key_ar}\s*[:\-]?\s*(.+)", text)
        if match:
            fields[key_en] = match.group(1).strip()
        else:
            fields[key_en] = ""
    return fields

def extract_text_fields(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join(page.get_text() for page in doc)
        return extract_fields_from_text(text)
    except Exception as e:
        return {v: "" for v in FIELD_KEYWORDS.values()}

def extract_tables_from_pdf(pdf_path):
    tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df = df[df.apply(is_data_row, axis=1)]
                    tables.append(df)
    except Exception as e:
        return []
    return tables

def process_pdf(pdf_path):
    fields = extract_text_fields(pdf_path)
    tables = extract_tables_from_pdf(pdf_path)

    if not tables:
        return None

    combined = pd.concat(tables, ignore_index=True)
    for field, value in fields.items():
        combined[field] = value
    combined["Source File"] = Path(pdf_path).name
    return combined

# =========================
# Streamlit UI
# =========================

st.title("📄 Arabic PDF Invoice Extractor")

uploaded_file = st.file_uploader("Upload a PDF or ZIP of PDFs", type=["pdf", "zip"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir = Path(tmpdir)

        # Save uploaded file
        file_path = temp_dir / uploaded_file.name
        file_path.write_bytes(uploaded_file.read())

        # Handle zip or single PDF
        pdf_files = []
        if uploaded_file.name.endswith(".zip"):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            pdf_files = list(temp_dir.rglob("*.pdf"))
        elif uploaded_file.name.endswith(".pdf"):
            pdf_files = [file_path]

        all_data = []

        for pdf_file in pdf_files:
            st.write(f"📄 Processing: {pdf_file.name}")
            result = process_pdf(pdf_file)
            if result is not None:
                all_data.append(result)
            else:
                st.warning(f"⚠️ No valid tables found in {pdf_file.name}")

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            output_path = temp_dir / "Cleaned_Invoices.xlsx"
            final_df.to_excel(output_path, index=False)

            st.success("✅ Extraction complete!")
            st.download_button("📥 Download Excel", data=output_path.read_bytes(), file_name="Cleaned_Invoices.xlsx")
        else:
            st.error("❌ No tables were extracted from any PDFs.")
