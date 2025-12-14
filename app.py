# streamlit_app.py

import streamlit as st
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
import unicodedata

# =========================
# Arabic Helpers
# =========================
def reshape_arabic_text(text):
    try:
        if text is None:
            return ""
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text) if text is not None else ""

def normalize_arabic_unicode(text: str) -> str:
    """
    Convert Arabic presentation forms (ﻓﺎﺗﻮرة / اﻟﻔﺎﺗﻮرة ...) to normal Arabic (فاتورة / الفاتورة ...)
    so regex keywords match in BOTH old and new PDFs.
    """
    if text is None:
        return ""
    text = str(text)
    # Convert compatibility forms to canonical where possible
    text = unicodedata.normalize("NFKC", text)
    # normalize spaces
    text = re.sub(r"\s+", " ", text).strip()
    return text

def clean_money(x: str) -> str:
    if x is None:
        return ""
    x = str(x)
    x = x.replace("٫", ".").replace("٬", ",")
    x = re.sub(r"[^\d\.,]", "", x)
    x = x.replace(",", "")
    return x.strip()

def to_float_safe(x):
    s = clean_money(x)
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None

# =========================
# Metadata Extraction (PyMuPDF + fallback)
# =========================
def extract_metadata(pdf_path: Path):
    try:
        # 1) Read text
        full_text = ""
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text("text") for page in doc])

        # fallback if needed
        if not full_text or len(full_text) < 50:
            with pdfplumber.open(pdf_path) as pdf:
                full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

        # 2) Normalize Arabic unicode (CRITICAL FIX)
        full_text = normalize_arabic_unicode(full_text)

        # 3) Robust field extraction (same line OR next line)
        def find_field(text, key_variants):
            """
            key_variants: list of possible labels to support both old/new layouts
            """
            for key in key_variants:
                # key : value (same line)
                pat1 = rf"{re.escape(key)}\s*[:：]?\s*([^\n]+)"
                m = re.search(pat1, text)
                if m and m.group(1).strip():
                    return m.group(1).strip()

                # key\nvalue (next line)
                pat2 = rf"{re.escape(key)}\s*[:：]?\s*\n\s*([^\n]+)"
                m2 = re.search(pat2, text)
                if m2 and m2.group(1).strip():
                    return m2.group(1).strip()

            return ""

        # 4) Use multiple label variants (old vs new)
        invoice_number = find_field(full_text, ["رقم الفاتورة", "الفاتورة رقم"])
        invoice_date   = find_field(full_text, ["تاريخ الفاتورة", "الفاتورة تاريخ"])

        customer_name  = find_field(full_text, ["اسم العميل", "العميل اسم"])
        # customer may wrap into next words; cut before tax/cr/address if present
        customer_name = re.split(r"الرقم الضريبي|رقم السجل|العنوان", customer_name)[0].strip()

        address_line   = find_field(full_text, ["العنوان"])
        cr_number      = find_field(full_text, ["رقم السجل", "السجل رقم"])

        paid           = clean_money(find_field(full_text, ["مدفوع"]))
        balance        = clean_money(find_field(full_text, ["الرصيد المستحق", "المستحق الرصيد"]))

        metadata = {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": f"{cr_number} {address_line}".strip(),
            "Paid": paid,
            "Balance": balance,
            "Source File": pdf_path.name
        }
        return metadata

    except Exception as e:
        st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
        return {}

# =========================
# Table Extraction (pdfplumber)
# =========================
def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def fix_shifted_rows(row):
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def extract_tables(pdf_path: Path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
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
                            headers = ["Total before tax", "الكمية", "Unit price", "Quantity", "Description", "SKU", "إضافي"]
                            df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                            all_data.append(df_cleaned)

            return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Error extracting table from {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Main Process Function
# =========================
def process_pdf(pdf_path: Path):
    metadata = extract_metadata(pdf_path)
    table_data = extract_tables(pdf_path)

    if not table_data.empty:
        for key, value in metadata.items():
            table_data[key] = value
        return table_data
    else:
        return pd.DataFrame([metadata])

# =========================
# Streamlit App UI
# =========================
st.set_page_config(page_title="Merged Arabic Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor Pdf to Excel")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if uploaded_file.name.lower().endswith(".zip"):
                with zipfile.ZipFile(file_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                for pdf in temp_dir.glob("*.pdf"):
                    pdf_paths.append(pdf)
            else:
                pdf_paths.append(file_path)

        all_data = []
        for pdf_path in pdf_paths:
            st.write(f"📄 Processing: {pdf_path.name}")
            df = process_pdf(pdf_path)
            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            # ======== Cleaning Steps ========
            if "Total before tax" in final_df.columns:
                final_df["Total before tax"] = final_df["Total before tax"].apply(to_float_safe)
                final_df["VAT 15%"] = (final_df["Total before tax"].fillna(0) * 0.15).round(2)
                final_df["Total after tax"] = (final_df["Total before tax"].fillna(0) + final_df["VAT 15%"]).round(2)

            for col in ["Paid", "Balance"]:
                if col in final_df.columns:
                    final_df[col] = final_df[col].apply(to_float_safe)

            # ======== Fix Invoice Date to MM/DD/YYYY ========
            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            # ======== Keep only required columns in order ========
            required_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Balance", "Paid", "Address",
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Description", "SKU",
                "Source File"
            ]

            # Ensure all required columns exist
            for c in required_columns:
                if c not in final_df.columns:
                    final_df[c] = None

            final_df = final_df.reindex(columns=required_columns)

            st.success("✅ Extraction & cleaning complete!")
            st.dataframe(final_df, use_container_width=True)

            output = BytesIO()
            final_df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name="Merged_Invoice_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No data extracted from the uploaded files.")
