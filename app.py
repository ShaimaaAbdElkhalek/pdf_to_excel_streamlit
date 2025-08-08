# streamlit_app.py

import streamlit as st
import os
import shutil
import fitz  # PyMuPDF
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

def process_pdf(pdf_path, safe_folder):
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

        ascii_name = f"bill_{pdf_path.stem.encode('ascii', errors='ignore').decode()}.pdf"
        safe_pdf_path = safe_folder / ascii_name
        shutil.copy(pdf_path, safe_pdf_path)

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
        st.error(f"❌ Error in {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Streamlit App UI
# =========================

st.title("📄 Arabic Invoice Field Extractor")

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
            extracted_df = process_pdf(pdf_path, safe_folder)
            if not extracted_df.empty:
                all_rows.append(extracted_df)

        if all_rows:
            final_df = pd.concat(all_rows, ignore_index=True)

            # Cleaning
            final_df["Customer Name"] = final_df["Customer Name"].astype(str).str.replace(r"اسم العميل\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")
            final_df["Address"] = final_df["Address"].astype(str).str.replace(r"العنوان\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")

            for col in ["Paid", "Balance"]:
                final_df[col] = final_df[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "", regex=False)
                final_df[col] = pd.to_numeric(final_df[col], errors="coerce")

            final_df = final_df[[
                "Invoice Number", "Invoice Date", "Customer Name", "Address", "Paid", "Balance", "Source File"
            ]]

            # Export Excel
            output_excel = temp_dir / "Cleaned_Invoice_Metadata.xlsx"
            final_df.to_excel(output_excel, index=False)

            st.success("✅ Field extraction complete!")
            st.download_button("📥 Download Extracted Excel", output_excel.read_bytes(), file_name="Cleaned_Invoices_Metadata.xlsx")
        else:
            st.warning("⚠️ No valid invoice fields found.")
