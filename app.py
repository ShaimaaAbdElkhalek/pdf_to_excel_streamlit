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

# =========================
# Arabic Helpers
# =========================
def reshape_arabic_text(text: str) -> str:
    try:
        if text is None:
            return ""
        text = str(text)
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    except Exception:
        return str(text) if text is not None else ""

# =========================
# Text Helpers
# =========================
def normalize_digits_punct(s: str) -> str:
    """Normalize Arabic punctuation variants and remove weird spaces."""
    if s is None:
        return ""
    s = str(s)
    # Arabic decimal separator variants
    s = s.replace("٫", ".").replace("٬", ",")
    # collapse spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s

def clean_money(x: str) -> str:
    """Keep digits and dot/comma then remove commas."""
    x = normalize_digits_punct(x)
    x = re.sub(r"[^\d\.,]", "", x)
    x = x.replace(",", "")
    return x.strip()

def to_float_safe(x):
    if x is None:
        return None
    s = clean_money(str(x))
    if s == "":
        return None
    try:
        return float(s)
    except Exception:
        return None

def get_full_text_fitz(pdf_path: Path) -> str:
    try:
        with fitz.open(pdf_path) as doc:
            return "\n".join([page.get_text("text") for page in doc]).strip()
    except Exception:
        return ""

def get_full_text_pdfplumber(pdf_path: Path) -> str:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join([(p.extract_text() or "") for p in pdf.pages]).strip()
    except Exception:
        return ""

def find_field(text: str, key: str) -> str:
    """
    Robust field extraction:
    - key: value
    - key value
    - key\nvalue
    """
    text = text or ""
    # Same line
    pat1 = rf"{re.escape(key)}\s*[:：]?\s*([^\n]+)"
    m = re.search(pat1, text)
    if m and m.group(1).strip():
        return m.group(1).strip()

    # Next line
    pat2 = rf"{re.escape(key)}\s*[:：]?\s*\n\s*([^\n]+)"
    m2 = re.search(pat2, text)
    if m2 and m2.group(1).strip():
        return m2.group(1).strip()

    return ""

# =========================
# Metadata Extraction (PyMuPDF + fallback)
# =========================
def extract_metadata(pdf_path: Path) -> dict:
    try:
        full_text = get_full_text_fitz(pdf_path)
        if not full_text or len(full_text) < 80:
            full_text = get_full_text_pdfplumber(pdf_path)

        full_text = normalize_digits_punct(full_text)

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")

        customer_name = find_field(full_text, "اسم العميل")
        # Remove tail like "الرقم الضريبي..."
        customer_name = re.split(r"الرقم الضريبي|رقم السجل|العنوان", customer_name)[0].strip()

        address_line = find_field(full_text, "العنوان")
        address_line = re.split(r"البند|الوصف|الرياض|جدة", address_line)[0].strip() if address_line else address_line

        cr_number = find_field(full_text, "رقم السجل")
        cr_number = re.split(r"العنوان|البند|الوصف", cr_number)[0].strip()

        paid = clean_money(find_field(full_text, "مدفوع"))
        balance = clean_money(find_field(full_text, "الرصيد المستحق"))

        # Some invoices use a different wording occasionally (optional fallback)
        if not balance:
            balance = clean_money(find_field(full_text, "الرصید المستحق"))

        # Build address
        full_address = f"{cr_number} {address_line}".strip()

        return {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": full_address,
            "Paid": paid,
            "Balance": balance,
            "Source File": pdf_path.name
        }

    except Exception as e:
        st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
        return {}

# =========================
# Table Extraction (pdfplumber)
# =========================
def is_data_row(row):
    """
    Row is data if it contains at least 2 numeric-ish cells.
    This makes it more stable across old/new PDFs.
    """
    numeric_cells = 0
    for cell in row:
        s = normalize_digits_punct(cell)
        s = re.sub(r"[^\d\.,]", "", s).replace(",", "")
        if s and re.fullmatch(r"\d+(\.\d+)?", s):
            numeric_cells += 1
    return numeric_cells >= 2

def fix_shifted_rows(row):
    """
    Your original heuristic kept, but guarded and normalized.
    """
    row = [normalize_digits_punct(x) for x in row]
    if len(row) == 7:
        # Guard against None/empty
        c3 = row[3].strip() if row[3] else ""
        c4 = row[4].strip() if row[4] else ""
        if c3 == "" and c4 != "":
            row[3] = row[4]
            row[4] = row[5]
            row[5] = row[6]
            row = row[:6]
    return row

def extract_tables(pdf_path: Path) -> pd.DataFrame:
    """
    Extract line items table(s).
    Works for both layouts by:
    - keeping your merge logic
    - being more tolerant about headers/column counts
    """
    try:
        all_data = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Try tables
                tables = page.extract_tables() or []
                for table in tables:
                    if not table:
                        continue

                    df = pd.DataFrame(table)
                    df = df.dropna(how="all").reset_index(drop=True)
                    if df.empty:
                        continue

                    merged_rows = []
                    temp_row = None

                    for _, r in df.iterrows():
                        row_values = r.fillna("").astype(str).tolist()
                        # Optional reshaping for readability
                        row_values = [reshape_arabic_text(cell) for cell in row_values]
                        row_values = fix_shifted_rows(row_values)

                        if is_data_row(row_values):
                            if temp_row:
                                # Merge previous text row (usually SKU/Arabic desc) with current
                                combined = [f"{temp_row[0]} {row_values[0]}".strip()] + row_values[1:]
                                merged_rows.append(combined)
                                temp_row = None
                            else:
                                merged_rows.append(row_values)
                        else:
                            # Save potential description continuation / header leftovers
                            temp_row = row_values

                    if merged_rows:
                        # Create a flexible header map depending on detected column count
                        num_cols = len(merged_rows[0])

                        # Your intended semantic columns (best effort)
                        base_headers = [
                            "Total before tax",  # line total
                            "الكمية",            # weight/qty unit (sometimes)
                            "Unit price",
                            "Quantity",
                            "Description",
                            "SKU",
                            "Extra"
                        ]
                        headers = base_headers[:num_cols]
                        df_cleaned = pd.DataFrame(merged_rows, columns=headers)

                        # Normalize expected output columns
                        # Ensure these columns exist even if missing in some layouts
                        for col in ["Total before tax", "Unit price", "Quantity", "Description", "SKU"]:
                            if col not in df_cleaned.columns:
                                df_cleaned[col] = ""

                        all_data.append(df_cleaned)

        if all_data:
            return pd.concat(all_data, ignore_index=True)
        return pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Error extracting table from {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Main Process Function
# =========================
def process_pdf(pdf_path: Path) -> pd.DataFrame:
    metadata = extract_metadata(pdf_path)
    table_data = extract_tables(pdf_path)

    # If we have line items, attach metadata to each row
    if not table_data.empty:
        for key, value in metadata.items():
            table_data[key] = value
        return table_data

    # Otherwise return just invoice-level metadata
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

        # Save uploads
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

            # ========= Cleaning & Computations =========
            # Line total
            if "Total before tax" in final_df.columns:
                final_df["Total before tax"] = final_df["Total before tax"].apply(to_float_safe)

                # VAT and after-tax at line level (as you currently do)
                final_df["VAT 15%"] = (final_df["Total before tax"].fillna(0) * 0.15).round(2)
                final_df["Total after tax"] = (final_df["Total before tax"].fillna(0) + final_df["VAT 15%"]).round(2)

            # Paid & Balance
            for col in ["Paid", "Balance"]:
                if col in final_df.columns:
                    final_df[col] = final_df[col].apply(to_float_safe)

            # Invoice Date to MM/DD/YYYY
            if "Invoice Date" in final_df.columns:
                # Your invoices are dd/mm/yyyy; dayfirst=True keeps it correct
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"].astype(str),
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            # ========= Keep only required columns in order =========
            required_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Balance", "Paid", "Address",
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Description", "SKU",
                "Source File"
            ]

            # Ensure missing columns exist
            for c in required_columns:
                if c not in final_df.columns:
                    final_df[c] = None

            final_df = final_df.reindex(columns=required_columns)

            st.success("✅ Extraction & cleaning complete!")
            st.dataframe(final_df, use_container_width=True)

            # ========= Export =========
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
