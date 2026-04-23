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
# OCR TEXT
# =========================

def extract_text_ocr(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang='ara+eng') + "\n"
    return text

# =========================
# NORMALIZE
# =========================

def normalize_text(text):
    fixes = {
        "الغاتورة": "الفاتورة",
        "اإلجمالي": "الإجمالي",
        "إلجمالي": "الإجمالي"
    }

    for k, v in fixes.items():
        text = text.replace(k, v)

    return text

# =========================
# METADATA
# =========================

def extract_metadata(text, filename):

    def find(key):
        m = re.search(rf"{key}\s*[:\-]?\s*(.+)", text)
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
# SKU EXTRACTION (FIXED)
# =========================

def extract_items_from_text(text):

    lines = text.split("\n")
    rows = []

    skip_keywords = [
        "المجموع", "الإجمالي", "الرصيد", "رقم الفاتورة",
        "تاريخ", "الايبان", "رقم الحساب", "مدفوع",
        "شركة", "فاتورة", "العنوان", "إلى", "من"
    ]

    for line in lines:

        line = line.strip()
        if not line:
            continue

        if any(k in line for k in skip_keywords):
            continue

        nums = re.findall(r"\d+\.?\d*", line)

        if len(nums) < 2:
            continue

        # split description safely
        description = re.split(r"\s{2,}", line)[0]

        description = re.sub(r"\d+\.?\d*", "", description)
        description = re.sub(r"[^\w\s\u0600-\u06FF]", " ", description)
        description = re.sub(r"\s+", " ", description).strip()

        if len(description) < 4:
            continue

        quantity = nums[0]
        price = nums[1]

        rows.append({
            "SKU / Description": description,
            "Quantity": quantity,
            "Unit Price": price
        })

    return pd.DataFrame(rows)

# =========================
# PROCESS PDF
# =========================

def process_pdf(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join([p.get_text() for p in doc])

    # OCR fallback
    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    text = normalize_text(text)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items_from_text(text)

    return text, meta, items

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor Pro", layout="wide")
st.title("📄 Invoice Extractor PRO (Debug + Clean SKU Extraction)")

files = st.file_uploader(
    "Upload PDF / ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

# store debug text
debug_texts = {}

if files:

    with tempfile.TemporaryDirectory() as tmp_dir:

        tmp = Path(tmp_dir)
        pdfs = []

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

            text, meta, items = process_pdf(pdf)

            # store for debug
            debug_texts[pdf.name] = text

            if not items.empty:
                for k, v in meta.items():
                    items[k] = v
                all_data.append(items)

        # =========================
        # SHOW EXTRACTED TEXT (DEBUG BUTTON)
        # =========================

        with st.expander("🔍 Show Extracted Raw Text (DEBUG)"):
            for name, text in debug_texts.items():
                st.subheader(name)
                st.text_area("Extracted Text", text, height=300)

        # =========================
        # FINAL OUTPUT
        # =========================

        if all_data:

            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ Extraction Done (Clean SKU Mode)")

            st.dataframe(final_df)

            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="clean_invoices.xlsx"
            )

        else:
            st.warning("No valid data extracted")
