import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

import pytesseract
from pdf2image import convert_from_path

# =========================
# OCR
# =========================

def extract_text_ocr(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang='ara+eng') + "\n"
    return text

# =========================
# NORMALIZE TEXT
# =========================

def normalize_text(text):
    replacements = {
        "الغاتورة": "الفاتورة",
        "اإلجمالي": "الإجمالي",
        "إلجمالي": "الإجمالي"
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    return text

# =========================
# METADATA EXTRACTION
# =========================

def extract_metadata(text, filename):

    def find(pattern):
        m = re.search(rf"{pattern}\s*[:\-]?\s*(.+)", text)
        return m.group(1).strip() if m else ""

    return {
        "Invoice Number": find("رقم الفاتورة"),
        "Invoice Date": find("تاريخ الفاتورة"),
        "Customer Name": find("اسم العميل"),
        "Address": find("العنوان"),
        "Paid": find("مدفوع"),
        "Balance": find("الإجمالي"),
        "Source File": filename
    }

# =========================
# SKU EXTRACTION (FIXED CORE)
# =========================

def extract_items_from_text(text):

    lines = text.split("\n")
    rows = []

    skip_keywords = [
        "المجموع", "الإجمالي", "الرصيد", "رقم الفاتورة",
        "تاريخ", "العنوان", "الايبان", "رقم الحساب",
        "مدفوع", "فاتورة", "شركة", "إلى", "من"
    ]

    for line in lines:

        line = line.strip()
        if not line:
            continue

        # 🚨 skip invoice/meta lines
        if any(k in line for k in skip_keywords):
            continue

        nums = re.findall(r"\d+\.?\d*", line)

        # must look like product line
        if len(nums) < 2:
            continue

        # remove obvious non-product junk
        if len(line) < 8:
            continue

        quantity = nums[-2]
        price = nums[-1]

        # clean description
        desc = re.sub(r"\d+\.?\d*", "", line)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 4:
            continue

        rows.append({
            "Description": desc,
            "Quantity": quantity,
            "Unit Price": price
        })

    return pd.DataFrame(rows)

# =========================
# PROCESS PDF
# =========================

def process_pdf(pdf_path):

    text = ""

    # try digital text first
    with fitz.open(pdf_path) as doc:
        text = "\n".join([p.get_text() for p in doc])

    # OCR fallback
    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    text = normalize_text(text)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items_from_text(text)

    # attach metadata to each SKU row
    if not items.empty:
        for k, v in meta.items():
            items[k] = v
        return items

    return pd.DataFrame([meta])

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor (Clean SKU Extraction Fixed)")

files = st.file_uploader(
    "Upload PDF / ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

if files:

    with tempfile.TemporaryDirectory() as tmp_dir:

        tmp = Path(tmp_dir)
        pdfs = []

        # save uploads
        for f in files:
            path = tmp / f.name
            with open(path, "wb") as out:
                out.write(f.read())

            if f.name.endswith(".zip"):
                with zipfile.ZipFile(path, "r") as z:
                    z.extractall(tmp)
                pdfs += list(tmp.glob("*.pdf"))
            else:
                pdfs.append(path)

        all_data = []

        for pdf in pdfs:
            st.write(f"📄 Processing: {pdf.name}")

            df = process_pdf(pdf)

            if not df.empty:
                all_data.append(df)

        if all_data:

            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ Extraction Completed Successfully")

            st.dataframe(final_df)

            # download
            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="invoices_clean.xlsx"
            )

        else:
            st.warning("No valid data extracted")
