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
import arabic_reshaper
from bidi.algorithm import get_display

# ========== Arabic Text Fix ==========
def reshape_arabic_text(text):
    try:
        reshaped = arabic_reshaper.reshape(text)
        bidi_text = get_display(reshaped)
        return bidi_text
    except:
        return text

# ========== Extract Fields ==========
def find_field(text, keyword):
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    return reshape_arabic_text(match.group(1).strip()) if match else ""

def extract_metadata(pdf_path, safe_folder):
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")
        customer_name = find_field(full_text, "فاتورة ضريبية") or find_field(full_text, "اسم العميل")
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
        st.error(f"❌ Error in {pdf_path.name} (metadata): {e}")
        return pd.DataFrame()

# ========== Extract Tables ==========
def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
               for cell in row)

def fix_shifted_rows(row):
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def extract_tables(pdf_path, fields_dict):
    all_tables = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df = pd.DataFrame(table)
                        df = df.dropna(how="all").reset_index(drop=True)
                        if df.empty:
                            continue

                        merged_rows = []
                        temp_row = []

                        for _, row in df.iterrows():
                            row_values = row.fillna("").astype(str).tolist()
                            row_values = [reshape_arabic_text(cell) for cell in row_values]
                            row_values = fix_shifted_rows(row_values)

                            if is_data_row(row_values):
                                if temp_row:
                                    combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
                                    merged_rows.append(combined)
                                    temp_row = []
                                else:
                                    merged_rows.append(row_values)
                            else:
                                temp_row = row_values

                        if merged_rows:
                            num_cols = len(merged_rows[0])
                            headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند", "إضافي"]
                            df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                            for key, value in fields_dict.items():
                                df_cleaned[key] = value
                            df_cleaned["Source File"] = pdf_path.name
                            all_tables.append(df_cleaned)
    except Exception as e:
        st.error(f"❌ Error in {pdf_path.name} (tables): {e}")
    return pd.concat(all_tables, ignore_index=True) if all_tables else pd.DataFrame()

# ========== Streamlit UI ==========
st.set_page_config(page_title="📄 Arabic Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor (Metadata + Tables)")

uploaded_files = st.file_uploader("Upload PDFs or a ZIP of PDFs", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                for pdf in temp_dir.glob("*.pdf"):
                    pdf_paths.append(pdf)
            else:
                pdf_paths.append(file_path)

        safe_folder = temp_dir / "safe"
        safe_folder.mkdir(exist_ok=True)

        all_metadata = []
        all_tables = []

        for pdf_path in pdf_paths:
            st.write(f"📄 Processing: {pdf_path.name}")
            metadata_df = extract_metadata(pdf_path, safe_folder)
            if not metadata_df.empty:
                all_metadata.append(metadata_df)
                table_df = extract_tables(pdf_path, metadata_df.iloc[0].to_dict())
                if not table_df.empty:
                    all_tables.append(table_df)

        if all_metadata:
            final_metadata = pd.concat(all_metadata, ignore_index=True)
            final_metadata["Customer Name"] = final_metadata["Customer Name"].astype(str).str.replace(r"اسم العميل\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")
            final_metadata["Address"] = final_metadata["Address"].astype(str).str.replace(r"العنوان\s*[:：]?\s*", "", regex=True).str.strip(" :：﹕")

            for col in ["Paid", "Balance"]:
                final_metadata[col] = final_metadata[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "", regex=False)
                final_metadata[col] = pd.to_numeric(final_metadata[col], errors="coerce")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    final_metadata.to_excel(writer, index=False, sheet_name="Metadata")
                    if all_tables:
                        final_tables = pd.concat(all_tables, ignore_index=True)
                        final_tables.to_excel(writer, index=False, sheet_name="Tables")

                st.success("✅ Extraction complete!")
                st.download_button("📥 Download Excel", tmp.name, file_name="Extracted_Invoices.xlsx")

        else:
            st.warning("⚠️ No valid invoice fields found.")
