import os
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st
from pathlib import Path
import re
import tabula

# ========== Streamlit App Setup ==========
st.set_page_config(page_title="Arabic PDF Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor")
st.markdown("Upload one or more **Arabic invoices (PDF)** and download the cleaned Excel file.")

# ========== Field Extraction Helpers ==========
def find_field(text, label):
    try:
        pattern = rf"{label}[\s:|\n]*([\u0600-\u06FF0-9a-zA-Z,.\-\/ ]+)"
        match = re.search(pattern, text)
        return match.group(1).strip() if match else ""
    except:
        return ""

# ========== Data Row Checker ==========
def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

# ========== Process PDF Function ==========
def process_pdf(pdf_path):
    all_rows = []
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")
        customer_name = find_field(full_text, "فاتورة ضريبية")
        address_part1 = find_field(full_text, "رقم السجل")
        address_part2 = find_field(full_text, "العنوان")
        address = f"{address_part1} {address_part2}" if address_part1 or address_part2 else ""
        paid_value = find_field(full_text, "مدفوع")
        balance_value = find_field(full_text, "الرصيد المستحق")

        # Extract tables using tabula
        tables = tabula.read_pdf(str(pdf_path), pages='all', multiple_tables=True, stream=False)

        for table in tables:
            if not table.empty:
                df = table.dropna(how="all")
                df["Invoice Number"] = invoice_number
                df["Invoice Date"] = invoice_date
                df["Customer Name"] = customer_name
                df["Address"] = address
                df["Paid"] = paid_value
                df["Balance"] = balance_value
                df["Source File"] = pdf_path.name
                all_rows.append(df)

        if all_rows:
            return pd.concat(all_rows, ignore_index=True)

    except Exception as e:
        st.warning(f"❌ Failed to process {pdf_path.name}: {e}")

    return pd.DataFrame()

# ========== Upload PDFs ==========
uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

# ========== Main Processing ==========
if uploaded_files:
    all_dataframes = []

    with st.spinner("🔄 Processing files..."):
        for uploaded_file in uploaded_files:
            file_path = Path(f"temp_{uploaded_file.name}")
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            df = process_pdf(file_path)
            if not df.empty:
                all_dataframes.append(df)
            os.remove(file_path)

    if all_dataframes:
        cleaned_df = pd.concat(all_dataframes, ignore_index=True)

        # Display preview
        st.success(f"✅ Extracted {len(cleaned_df)} rows from {len(uploaded_files)} file(s)")
        st.dataframe(cleaned_df.head(50), use_container_width=True)

        # Download Excel
        excel_path = "Cleaned_Combined_Tables.xlsx"
        cleaned_df.to_excel(excel_path, index=False, engine="openpyxl")
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 Download Excel",
                data=f,
                file_name=excel_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ No valid data found in uploaded PDFs.")
else:
    st.info("📤 Please upload one or more PDF files.")
