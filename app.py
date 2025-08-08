import streamlit as st
import os
import shutil
import fitz  # PyMuPDF
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
import pdfplumber

# =========================
# Helper Functions
# =========================

def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

def find_field(text, keyword):
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""

def process_pdf(pdf_path, safe_folder):
    all_rows = []
    try:
        # Extract full text using PyMuPDF
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")
        customer_name = find_field(full_text, "فاتورة ضريبية")
        address_part2 = find_field(full_text, "العنوان")
        address_part1 = find_field(full_text, "رقم السجل")
        address = f"{address_part1} {address_part2}" if address_part1 or address_part2 else ""
        paid_value = find_field(full_text, "مدفوع")
        balance_value = find_field(full_text, "الرصيد المستحق")

        # Safe file copy
        ascii_name = f"bill_{pdf_path.stem.encode('ascii', errors='ignore').decode()}.pdf"
        safe_pdf_path = safe_folder / ascii_name
        shutil.copy(pdf_path, safe_pdf_path)

        # Extract tables using pdfplumber
        with pdfplumber.open(safe_pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    cleaned = [row for row in table if is_data_row(row)]
                    if cleaned:
                        df = pd.DataFrame(cleaned)
                        all_tables.append(df)

        if all_tables:
            combined_df = pd.concat(all_tables, ignore_index=True)

            # Try assigning column names
            if len(combined_df.columns) == 6:
                combined_df.columns = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند"]
            else:
                combined_df.columns = [f"col_{i}" for i in range(len(combined_df.columns))]

            # Add metadata columns
            combined_df["Invoice Number"] = invoice_number
            combined_df["Invoice Date"] = invoice_date
            combined_df["Customer Name"] = customer_name
            combined_df["Address"] = address
            combined_df["Paid"] = paid_value
            combined_df["Balance"] = balance_value
            combined_df["Source File"] = pdf_path.name

            all_rows.append(combined_df)
        else:
            st.warning(f"⚠️ No valid tables found in {pdf_path.name}")
    except Exception as e:
        st.error(f"❌ Error in {pdf_path.name}: {e}")
    return all_rows

# =========================
# Streamlit App UI
# =========================

st.title("📄 Arabic Invoice Table Extractor")

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
                pdf_paths += list(temp_dir.glob("*.pdf"))
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

            # Cleaning
            final_df["Customer Name"] = final_df["Customer Name"].astype(str).str.replace(r"اسم العميل\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")
            final_df["Address"] = final_df["Address"].astype(str).str.replace(r"العنوان\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")

            for col in ["Paid", "Balance"]:
                final_df[col] = final_df[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "", regex=False).astype(float)

            if "العدد" in final_df.columns:
                final_df["العدد"] = pd.to_numeric(final_df["العدد"], errors="coerce")
            if "المجموع" in final_df.columns:
                final_df["المجموع"] = final_df["المجموع"].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "", regex=False).astype(float)
                final_df["VAT 15% Calc"] = (final_df["المجموع"] * 0.15).round(2)
                final_df = final_df.rename(columns={
                    "المجموع": "Total before tax",
                    "سعر الوحدة": "Unit price",
                    "العدد": "Quantity",
                    "الوصف": "Description",
                    "البند": "SKU"
                })
                final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15% Calc"]).round(2)

            final_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Address", "Paid", "Balance",
                "Total before tax", "VAT 15% Calc", "Total after tax",
                "Unit price", "Quantity", "Description", "SKU", "Source File"
            ]

            # Keep only available columns
            final_df = final_df[[col for col in final_columns if col in final_df.columns]]

            # Export Excel
            output_excel = temp_dir / "Cleaned_Combined_Tables.xlsx"
            final_df.to_excel(output_excel, index=False)

            st.success("✅ Extraction complete!")
            st.download_button("📥 Download Cleaned Excel", output_excel.read_bytes(), file_name="Cleaned_Invoices.xlsx")
        else:
            st.warning("⚠️ No valid tables found.")
