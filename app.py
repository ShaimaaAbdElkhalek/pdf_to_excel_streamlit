# streamlit_app.py

import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import unicodedata

# =========================
# TEXT NORMALIZATION
# =========================
def normalize_text(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("٫", ".").replace("٬", ",")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def clean_money(x):
    if not x:
        return ""
    x = normalize_text(x)
    x = re.sub(r"[^\d\.,]", "", x)
    return x.replace(",", "")

def to_float_safe(x):
    try:
        return float(clean_money(x))
    except:
        return None

# =========================
# ORDER-INDEPENDENT FIND
# =========================
def find_by_tokens(text, tokens, value_pattern):
    """
    Finds a value near given tokens (any order, any line)
    """
    text = normalize_text(text)

    for m in re.finditer(value_pattern, text):
        start, end = m.span()
        window = text[max(0, start-50): min(len(text), end+50)]

        if all(tok in window for tok in tokens):
            return m.group(1)

    return ""

# =========================
# METADATA EXTRACTION
# =========================
def extract_metadata(pdf_path):
    try:
        # Read text
        with fitz.open(pdf_path) as doc:
            text = " ".join(p.get_text("text") for p in doc)

        if not text.strip():
            with pdfplumber.open(pdf_path) as pdf:
                text = " ".join(p.extract_text() or "" for p in pdf.pages)

        text = normalize_text(text)

        invoice_number = find_by_tokens(
            text,
            tokens=["فاتورة", "رقم"],
            value_pattern=r"(\d{3,})"
        )

        invoice_date = find_by_tokens(
            text,
            tokens=["فاتورة", "تاريخ"],
            value_pattern=r"(\d{2}/\d{2}/\d{4})"
        )

        customer = find_by_tokens(
            text,
            tokens=["اسم", "العميل"],
            value_pattern=r"(?:اسم العميل|العميل اسم)\s*([^\d]+?)\s*(?:الرقم|رقم|العنوان)"
        )

        address = find_by_tokens(
            text,
            tokens=["العنوان"],
            value_pattern=r"العنوان\s*[:：]?\s*([^\d]+?)\s*(?:البند|الوصف|الرياض|جدة)"
        )

        paid = find_by_tokens(
            text,
            tokens=["مدفوع"],
            value_pattern=r"مدفوع\s*([0-9.,]+)"
        )

        balance = find_by_tokens(
            text,
            tokens=["الرصيد", "المستحق"],
            value_pattern=r"الرصيد المستحق\s*([0-9.,]+)"
        )

        return {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer,
            "Address": address,
            "Paid": paid,
            "Balance": balance,
            "Source File": pdf_path.name
        }

    except Exception as e:
        st.error(f"Metadata error in {pdf_path.name}: {e}")
        return {}

# =========================
# TABLE EXTRACTION
# =========================
def is_data_row(row):
    return any(
        str(c).replace(",", "").replace(".", "").isdigit()
        for c in row
    )

def extract_tables(pdf_path):
    rows = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    for row in table:
                        if row and is_data_row(row):
                            rows.append(row)

        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        df.columns = ["Description", "Quantity", "Unit price", "Total before tax"][:len(df.columns)]
        return df

    except Exception as e:
        st.error(f"Table error in {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# PROCESS PDF
# =========================
def process_pdf(pdf_path):
    meta = extract_metadata(pdf_path)
    table = extract_tables(pdf_path)

    if not table.empty:
        for k, v in meta.items():
            table[k] = v
        return table

    return pd.DataFrame([meta])

# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="Arabic Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor (Order-Independent)")

files = st.file_uploader("Upload PDFs or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

if files:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdfs = []

        for f in files:
            path = tmp / f.name
            path.write_bytes(f.read())

            if path.suffix.lower() == ".zip":
                with zipfile.ZipFile(path) as z:
                    z.extractall(tmp)
                pdfs.extend(tmp.glob("*.pdf"))
            else:
                pdfs.append(path)

        all_data = []
        for pdf in pdfs:
            st.write("Processing:", pdf.name)
            df = process_pdf(pdf)
            if not df.empty:
                all_data.append(df)

        if all_data:
            final = pd.concat(all_data, ignore_index=True)

            final["Paid"] = final["Paid"].apply(to_float_safe)
            final["Balance"] = final["Balance"].apply(to_float_safe)

            final["Invoice Date"] = pd.to_datetime(
                final["Invoice Date"], errors="coerce", dayfirst=True
            ).dt.strftime("%m/%d/%Y")

            st.success("✅ Extraction complete")
            st.dataframe(final)

            output = BytesIO()
            final.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "📥 Download Excel",
                output,
                "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No data extracted")
