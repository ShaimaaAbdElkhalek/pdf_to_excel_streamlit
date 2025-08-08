# streamlit_app.py

import streamlit as st
import os
import shutil
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

def find_field(text, keyword):
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""

def extract_metadata(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")
        customer_name = find_field(full_text, "فاتورة ضريبية")
        address_part2 = find_field(full_text, "العنوان")
        address_part1 = find_field(full_text, "رقم السجل")
        address = f"{address_part1} {address_part2}".strip()
        paid_value = find_field(full_text, "مدفوع")
        balance_value = find_field(full_text, "الرصيد المستحق")

        return pd.DataFrame([{
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": address,
            "Paid": paid_value,
            "Balance": balance_value,
            "Source File": pdf_path.name
        }])
    except Exception as e:
        st.error(f"❌ Error in metadata extraction for {pdf_path.name}: {e}")
        return pd.DataFrame()

def extract_tables(pdf_path):
    all_rows = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df["Source File"] = pdf_path.name
                    all_rows.append(df)
        return pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Error in table extraction for {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Streamlit App UI
# =========================

st.title("📄 Arabic Invoice Extractor (Metadata + Tables)")

uploaded_files = st.file_uploader("Upload PDF files or a ZIP of PDFs", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

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

        metadata_rows = []
        table_rows = []

        for pdf_path in pdf_paths:
            st.write(f"🔍 Processing: {pdf_path.name}")
            metadata_df = extract_metadata(pdf_path)
            table_df = extract_tables(pdf_path)

            if not metadata_df.empty:
                metadata_rows.append(metadata_df)
            if not table_df.empty:
                table_rows.append(table_df)

        if metadata_rows or table_rows:
            output_excel = temp_dir / "Extracted_Invoices.xlsx"
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                if metadata_rows:
                    meta_df = pd.concat(metadata_rows, ignore_index=True)
                    meta_df["Customer Name"] = meta_df["Customer Name"].astype(str).str.replace(r"اسم العميل\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")
                    meta_df["Address"] = meta_df["Address"].astype(str).str.replace(r"العنوان\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")

                    for col in ["Paid", "Balance"]:
                        meta_df[col] = meta_df[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "", regex=False)
                        meta_df[col] = pd.to_numeric(meta_df[col], errors="coerce")

                    meta_df.to_excel(writer, index=False, sheet_name="Metadata")

                if table_rows:
                    tables_df = pd.concat(table_rows, ignore_index=True)
                    tables_df.to_excel(writer, index=False, sheet_name="Tables")

            st.success("✅ Extraction complete!")
            st.download_button("📥 Download Excel", output_excel.read_bytes(), file_name="Extracted_Invoices.xlsx")
        else:
            st.warning("⚠️ No valid metadata or tables found.")
