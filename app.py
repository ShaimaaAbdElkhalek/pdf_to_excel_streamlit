import streamlit as st
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
    """Check if row contains at least one number (Arabic or English)"""
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def find_field(text, keyword):
    """Find field value following a keyword (e.g. 'رقم الفاتورة: 123')"""
    pattern = rf"{keyword}\s*[:：]?\s*(.+)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else None

def extract_text_fields(pdf_path):
    """Extract key Arabic metadata fields from PDF"""
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc)
        return {
            "اسم العميل": find_field(text, "اسم العميل"),
            "رقم الفاتورة": find_field(text, "رقم الفاتورة"),
            "العنوان": find_field(text, "العنوان"),
        }

def extract_tables_from_pdf(pdf_path):
    """Extract tables using pdfplumber and filter numeric rows"""
    combined_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                cleaned = [row for row in table if is_data_row(row)]
                if cleaned:
                    combined_rows.extend(cleaned)
    return pd.DataFrame(combined_rows) if combined_rows else None

def save_temp_file(uploaded_file):
    """Save uploaded file to a temporary file and return path"""
    suffix = Path(uploaded_file.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.read())
        return tmp.name

# ======================
# Streamlit UI
# ======================
st.title("📄 Arabic PDF Invoice Extractor")

uploaded_files = st.file_uploader("Upload PDF files or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Processing..."):
        temp_folder = Path(tempfile.mkdtemp())
        all_dfs = []

        for uploaded_file in uploaded_files:
            if uploaded_file.name.endswith(".zip"):
                zip_path = save_temp_file(uploaded_file)
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_folder)
                pdf_paths = list(temp_folder.rglob("*.pdf"))
            else:
                pdf_path = save_temp_file(uploaded_file)
                pdf_paths = [Path(pdf_path)]

            for pdf_path in pdf_paths:
                try:
                    fields = extract_text_fields(str(pdf_path))
                    table_df = extract_tables_from_pdf(str(pdf_path))
                    if table_df is not None:
                        for k, v in fields.items():
                            table_df[k] = v
                        table_df["اسم الملف"] = pdf_path.name
                        all_dfs.append(table_df)
                    else:
                        st.warning(f"⚠️ No valid tables found in: {pdf_path.name}")
                except Exception as e:
                    st.error(f"❌ Error in {pdf_path.name}: {e}")

        if all_dfs:
            result_df = pd.concat(all_dfs, ignore_index=True)
            st.success(f"✅ Extracted {len(result_df)} rows from {len(all_dfs)} files.")
            st.dataframe(result_df)

            output_path = temp_folder / "cleaned_output.xlsx"
            result_df.to_excel(output_path, index=False)

            with open(output_path, "rb") as f:
                st.download_button("⬇️ Download Excel File", f, file_name="cleaned_output.xlsx")
