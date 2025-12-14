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
import unicodedata

# =========================
# Helpers
# =========================

def normalize_text(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("٫", ".").replace("٬", ",")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_money(x):
    if not x:
        return ""
    x = normalize_text(x)
    x = re.sub(r"[^\d\.,]", "", x).replace(",", "")
    return x


def to_float_safe(x):
    x = clean_money(x)
    try:
        return float(x)
    except:
        return None


# =========================
# FIND FIELD (this OR that)
# =========================

def find_field(text, keywords):
    """
    keywords: string OR list of strings
    """
    if isinstance(keywords, str):
        keywords = [keywords]

    for key in keywords:
        # same line
        p1 = rf"{re.escape(key)}\s*[:：]?\s*([^\n]+)"
        m = re.search(p1, text)
        if m and m.group(1).strip():
            return m.group(1).strip()

        # next line
        p2 = rf"{re.escape(key)}\s*[:：]?\s*\n\s*([^\n]+)"
        m2 = re.search(p2, text)
        if m2 and m2.group(1).strip():
            return m2.group(1).strip()

    return ""


# =========================
# METADATA EXTRACTION
# =========================

def extract_metadata(pdf_path):
    try:
        # Read text
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join(p.get_text("text") for p in doc)

        if not full_text.strip():
            with pdfplumber.open(pdf_path) as pdf:
                full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        full_text = normalize_text(full_text)

        invoice_number = find_field(
            full_text,
            ["رقم الفاتورة", "الفاتورة رقم"]
        )

        invoice_date = find_field(
            full_text,
            ["تاريخ الفاتورة", "الفاتورة تاريخ"]
        )

        customer_name = find_field(
            full_text,
            ["اسم العميل", "العميل اسم"]
        )
        customer_name = re.split(r"الرقم الضريبي|رقم السجل|العنوان", customer_name)[0].strip()

        address = find_field(full_text, "العنوان")
        cr = find_field(full_text, ["رقم السجل", "السجل رقم"])

        paid = clean_money(find_field(full_text, "مدفوع"))
        balance = clean_money(find_field(full_text, ["الرصيد المستحق", "المستحق الرصيد"]))

        return {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": f"{cr} {address}".strip(),
            "Paid": paid,
            "Balance": balance,
            "Source File": pdf_path.name
        }

    except Exception as e:
        st.error(f"Metadata error in {pdf_path.name}: {e}")
        return {}


# =========================
# TABLE EXTRACTION (unchanged)
# =========================

def is_data_row(row):
    return any(
        str(c).replace(",", "").replace(".", "").isdigit()
        for c in row
    )


def extract_tables(pdf_path):
    try:
        all_rows = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    for row in table:
                        if row and is_data_row(row):
                            all_rows.append(row)

        if not all_rows:
            return pd.DataFrame()

        df = pd.DataFrame(all_rows)
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

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Arabic Invoice Extractor (Old + New PDFs)")

files = st.file_uploader(
    "Upload PDF or ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

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

            st.success("Done ✅")
            st.dataframe(final)

            output = BytesIO()
            final.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "Download Excel",
                output,
                "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
