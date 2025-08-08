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
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def extract_fields(text, patterns):
    result = {}
    for label, pattern in patterns.items():
        match = re.search(pattern, text)
        result[label] = match.group(1).strip() if match else None
    return result

def extract_text_fields(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    fields = extract_fields(full_text, {
        "رقم الفاتورة": r"رقم الفاتورة[:\s\-]*([\d\w/\\\-]+)",
        "اسم العميل": r"اسم العميل[:\s\-]*([\u0600-\u06FF\s\w]+)",
        "العنوان": r"العنوان[:\s\-]*([\u0600-\u06FF\s\w\d,.-]+)",
        "التاريخ": r"التاريخ[:\s\-]*([\d/\-]+)"
    })
    return fields

def extract_table(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    cleaned = [row for row in table if is_data_row(row)]
                    if cleaned:
                        df = pd.DataFrame(cleaned)
                        all_tables.append(df)
            if all_tables:
                return pd.concat(all_tables, ignore_index=True)
            else:
                return None
    except Exception as e:
        st.warning(f"❌ Failed to extract table from {pdf_path.name}: {e}")
        return None

def process_pdf(pdf_path):
    try:
        fields = extract_text_fields(pdf_path)
        table = extract_table(pdf_path)
        if table is not None:
            for key, val in fields.items():
                table[key] = val
            return table
        else:
            st.warning(f"⚠️ No valid tables found in {pdf_path.name}")
            return None
    except Exception as e:
        st.error(f"❌ Error in {pdf_path.name}: {e}")
        return None

def process_uploaded_file(uploaded_file):
    temp_dir = Path(tempfile.mkdtemp())
    extracted_dfs = []

    if uploaded_file.name.endswith(".zip"):
        with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        files = list(temp_dir.glob("*.pdf"))
    elif uploaded_file.name.endswith(".pdf"):
        temp_pdf = temp_dir / uploaded_file.name
        with open(temp_pdf, "wb") as f:
            f.write(uploaded_file.read())
        files = [temp_pdf]
    else:
        st.error("Please upload a PDF or ZIP file.")
        return []

    for file in files:
        st.write(f"📄 Processing: {file.name}")
        df = process_pdf(file)
        if df is not None:
            extracted_dfs.append(df)

    return extracted_dfs

# ======================
# Streamlit UI
# ======================

st.set_page_config(page_title="📄 Arabic Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Table + Field Extractor (No Java)")

uploaded_file = st.file_uploader("Upload PDF or ZIP file with invoices", type=["pdf", "zip"])

if uploaded_file:
    with st.spinner("Processing..."):
        dfs = process_uploaded_file(uploaded_file)
        if dfs:
            combined_df = pd.concat(dfs, ignore_index=True)
            st.success("✅ Extraction complete!")
            st.dataframe(combined_df)

            # Download Excel
            temp_xlsx = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            combined_df.to_excel(temp_xlsx.name, index=False)
            st.download_button("📥 Download Excel", data=open(temp_xlsx.name, 'rb'), file_name="extracted_data.xlsx")
        else:
            st.error("❌ No valid data extracted.")
