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
from io import BytesIO
import arabic_reshaper
from bidi.algorithm import get_display

# ========== Functions ==========

def reshape_arabic(text):
    """Reshape Arabic text for proper display in Excel/Streamlit."""
    if not text:
        return ""
    return get_display(arabic_reshaper.reshape(text))

def extract_metadata_fields(pdf_path):
    """Extract Arabic metadata fields using PyMuPDF (fitz)."""
    fields = {
        "رقم الفاتورة": "",
        "تاريخ الفاتورة": "",
        "اسم العميل": "",
        "العنوان": "",
    }

    patterns = {
        "رقم الفاتورة": r"رقم الفاتورة[:\- ]+(\d+)",
        "تاريخ الفاتورة": r"تاريخ الفاتورة[:\- ]+([\d\/\-]+)",
        "اسم العميل": r"اسم العميل[:\- ]+([\u0600-\u06FF\s]+)",
        "العنوان": r"العنوان[:\- ]+([\u0600-\u06FF\s\d\-]+)",
    }

    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text")

    for field, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            fields[field] = reshape_arabic(match.group(1).strip())

    return fields


def extract_tables(pdf_path):
    """Extract tables using pdfplumber (no fixed column count)."""
    tables_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    # Create DataFrame directly from table
                    df = pd.DataFrame(table)

                    # First row is usually header
                    df.columns = df.iloc[0]
                    df = df.drop(0).reset_index(drop=True)

                    # Reshape Arabic text in all cells
                    df = df.applymap(lambda x: reshape_arabic(str(x)) if x else "")

                    tables_data.append(df)

    return tables_data


def process_pdf(pdf_path):
    """Combine field + table extraction for one PDF."""
    metadata = extract_metadata_fields(pdf_path)
    tables = extract_tables(pdf_path)

    if tables:
        final_table = tables[0].copy()
        # Attach metadata to each row
        for key, value in metadata.items():
            final_table[key] = value
        return final_table
    else:
        return pd.DataFrame([metadata])


def process_files(uploaded_files):
    """Process uploaded PDFs or ZIPs and return merged DataFrame."""
    all_data = []

    for uploaded_file in uploaded_files:
        # If ZIP -> extract and process PDFs inside
        if uploaded_file.name.endswith(".zip"):
            with tempfile.TemporaryDirectory() as tmpdir:
                zip_path = os.path.join(tmpdir, uploaded_file.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(tmpdir)

                for file in Path(tmpdir).rglob("*.pdf"):
                    try:
                        df = process_pdf(str(file))
                        all_data.append(df)
                    except Exception as e:
                        st.error(f"❌ Error extracting from {file.name}: {e}")

        else:  # Single PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                tmpfile.write(uploaded_file.getbuffer())
                tmp_path = tmpfile.name

            try:
                df = process_pdf(tmp_path)
                all_data.append(df)
            except Exception as e:
                st.error(f"❌ Error extracting from {uploaded_file.name}: {e}")

    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()


# ========== Streamlit UI ==========

st.title("📑 Arabic Invoice Extractor")
st.write("Upload Arabic PDF invoices (or a ZIP of PDFs). The app will extract fields + tables.")

uploaded_files = st.file_uploader("Upload PDF or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    df = process_files(uploaded_files)

    if not df.empty:
        st.success("✅ Extraction complete!")
        st.dataframe(df)

        # Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Invoices")

        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name="extracted_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No data extracted.")
