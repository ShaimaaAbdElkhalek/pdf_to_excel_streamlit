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
        text += pytesseract.image_to_string(img, lang='ara+eng') + "\n"
    return text

# =========================
# CLEAN TEXT
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# METADATA
# =========================

def extract_metadata(text, filename):

    def find(k):
        m = re.search(rf"{k}\s*[:\-]?\s*(.+)", text)
        return m.group(1).strip() if m else ""

    return {
        "Invoice Number": find("رقم الفاتورة"),
        "Date": find("تاريخ"),
        "Customer": find("اسم العميل"),
        "Source": filename
    }

# =========================
# 🔥 FINAL SAFE EXTRACTION ENGINE
# =========================

def extract_items(text):

    text = clean(text)

    rows = []

    # split loosely by sentences (NOT lines)
    chunks = re.split(r"[|\n]", text)

    for chunk in chunks:

        chunk = chunk.strip()
        if len(chunk) < 8:
            continue

        # must contain at least ONE number
        nums = re.findall(r"\d+\.\d+|\d+", chunk)

        if len(nums) == 0:
            continue

        # must contain some letters (product text)
        if not any(c.isalpha() for c in chunk):
            continue

        # remove pure invoice noise
        if any(x in chunk for x in [
            "المجموع", "الإحمالي", "الرصيد",
            "الايبان", "رقم الحساب", "شركة"
        ]):
            continue

        # clean description
        desc = re.sub(r"\d+\.\d+|\d+", "", chunk)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 4:
            continue

        # fallback logic:
        quantity = nums[-2] if len(nums) >= 2 else nums[0]
        price = nums[-1]

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": price
        })

    return pd.DataFrame(rows)

# =========================
# PROCESS
# =========================

def process(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join([p.get_text() for p in doc])

    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items(text)

    return text, meta, items

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice AI FINAL", layout="wide")
st.title("📄 Invoice Extractor AI (Works on Broken OCR Invoices)")

files = st.file_uploader("Upload PDF / ZIP", type=["pdf", "zip"], accept_multiple_files=True)

debug = {}

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

            st.write(f"📄 {pdf.name}")

            text, meta, items = process(pdf)

            debug[pdf.name] = text

            # 🚨 IMPORTANT: even if empty, we force fallback row
            if items.empty:
                items = pd.DataFrame([{
                    "SKU / Description": "⚠️ Manual Review Required",
                    "Quantity": "",
                    "Unit Price": ""
                }])

            for k, v in meta.items():
                items[k] = v

            all_data.append(items)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Raw Extracted Text"):
            for k, v in debug.items():
                st.subheader(k)
                st.text_area("text", v, height=300)

        # =========================
        # OUTPUT
        # =========================

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Extraction Completed (AI Robust Mode)")

        st.dataframe(final_df)

        buffer = BytesIO()
        final_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Download Excel",
            buffer,
            file_name="invoice_ai_final.xlsx"
        )
