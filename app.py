import streamlit as st
import os
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
from pathlib import Path

# ========== CLEANING FUNCTIONS ==========
def clean_total_before_tax(value):
    """Remove non-numeric characters and specific Arabic words from Total before tax."""
    if pd.isna(value):
        return value
    value = str(value)
    value = value.replace("المجموع", "")  # remove the Arabic word "المجموع"
    value = re.sub(r"[^\d.,]", "", value)  # keep only numbers and decimal separators
    return value.strip()

def clean_address_or_name(value):
    """Remove only the exact words 'اسم العميل' or 'العنوان' (and colons after them), keep the rest."""
    if pd.isna(value):
        return value
    value = str(value)
    # Remove the words and an optional colon after them
    value = re.sub(r"(اسم العميل|العنوان)\s*:?", "", value)
    return value.strip()

# ========== PDF EXTRACTION ==========
def extract_tables_from_pdf(pdf_path):
    """Extract tables from a PDF using pdfplumber."""
    tables_list = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table)
                tables_list.append(df)
    return tables_list

# ========== MAIN PROCESS ==========
def process_pdfs(source_folder):
    """Process all PDFs in a folder and return cleaned combined DataFrame."""
    all_tables = []

    for file in os.listdir(source_folder):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(source_folder, file)
            pdf_tables = extract_tables_from_pdf(pdf_path)
            for table in pdf_tables:
                all_tables.append(table)

    if not all_tables:
        return pd.DataFrame()

    # Combine all extracted tables
    combined_df = pd.concat(all_tables, ignore_index=True)

    # Rename columns if needed
    combined_df.columns = [str(col).strip() for col in combined_df.columns]

    # Apply cleaning functions
    if "Total before tax" in combined_df.columns:
        combined_df["Total before tax"] = combined_df["Total before tax"].apply(clean_total_before_tax)

    if "Address" in combined_df.columns:
        combined_df["Address"] = combined_df["Address"].apply(clean_address_or_name)

    if "Customer Name" in combined_df.columns:
        combined_df["Customer Name"] = combined_df["Customer Name"].apply(clean_address_or_name)

    return combined_df

# ========== STREAMLIT UI ==========
st.title("📄 PDF Table Extractor & Cleaner")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    temp_folder = "temp_pdfs"
    os.makedirs(temp_folder, exist_ok=True)

    for uploaded_file in uploaded_files:
        with open(os.path.join(temp_folder, uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())

    df_cleaned = process_pdfs(temp_folder)

    if not df_cleaned.empty:
        st.subheader("Cleaned Data")
        st.dataframe(df_cleaned)

        # Download link
        output_path = "Cleaned_Combined_Tables.xlsx"
        df_cleaned.to_excel(output_path, index=False)
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Cleaned Excel",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("No tables found in uploaded PDFs.")
