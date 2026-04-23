import streamlit as st
import fitz
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
        text += pytesseract.image_to_string(img, lang="ara+eng") + "\n"
    return text

# =========================
# CLEAN TEXT
# =========================

def clean_text(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
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
        "Invoice Date": find("تاريخ"),
        "Customer Name": find("اسم العميل"),
        "Tax Number": find("الرقم الضريبي"),
        "Source File": filename
    }

# =========================
# 🔥 CORE ENGINE (ROBUST OCR EXTRACTION)
# =========================

def extract_items(text):

    text = clean_text(text)

    rows = []

    # STEP 1: find ALL numbers with positions
    matches = [(m.group(), m.start()) for m in re.finditer(r"\d+\.\d+|\d+", text)]

    # STEP 2: sliding window extraction
    for i in range(len(matches) - 1):

        num1, pos1 = matches[i]
        num2, pos2 = matches[i + 1]

        # skip huge numbers (totals / IBAN / noise)
        if len(num1) > 6 or len(num2) > 6:
            continue

        # STEP 3: context window around numbers
        start = max(0, pos1 - 120)
        end = min(len(text), pos2 + 120)

        window = text[start:end]

        # must contain letters (product check)
        if not any(c.isalpha() for c in window):
            continue

        # STEP 4: clean description
        desc = re.sub(r"\d+\.\d+|\d+", "", window)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 6:
            continue

        # filter noise
        if any(x in desc for x in ["المجموع", "الإحمالي", "الرصيد", "الايبان"]):
            continue

        quantity = num1
        price = num2

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": price
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.drop_duplicates()

    return df

# =========================
# PROCESS PDF
# =========================

def process_pdf(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join([page.get_text() for page in doc])

    # OCR fallback
    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items(text)

    return text, meta, items

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor PRO", layout="wide")
st.title("📄 Invoice Extractor PRO (Handles Fully Unstructured OCR)")

files = st.file_uploader("Upload PDF / ZIP", type=["pdf", "zip"], accept_multiple_files=True)

debug_store = {}

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

            debug_store[pdf.name] = text

            # force fallback if empty (never crash)
            if items.empty:
                items = pd.DataFrame([{
                    "SKU / Description": "⚠️ Needs manual review (OCR too noisy)",
                    "Quantity": "",
                    "Unit Price": ""
                }])

            for k, v in meta.items():
                items[k] = v

            all_data.append(items)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show Extracted Raw Text"):
            for name, txt in debug_store.items():
                st.subheader(name)
                st.text_area("text", txt, height=300)

        # =========================
        # OUTPUT
        # =========================

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Extraction Completed (Robust OCR Engine)")

        st.dataframe(final_df)

        buffer = BytesIO()
        final_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Download Excel",
            buffer,
            file_name="invoice_output.xlsx"
        )
