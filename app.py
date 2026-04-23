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
# OCR fallback
# =========================

def ocr_pdf(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    return "\n".join(pytesseract.image_to_string(img, lang="ara+eng") for img in images)

# =========================
# clean text
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# 🔥 TABLE EXTRACTION (FIXED FOR YOUR FORMAT)
# =========================

def extract_table(text):

    text = clean(text)

    rows = []

    # 1. isolate ONLY product table section
    match = re.search(r"العدد.*?(المجموع|الإحمالي|الرصيد المستحق)", text, re.S)
    table_text = match.group(0) if match else text

    # 2. split by product boundary (English words = product marker)
    chunks = re.split(r"(?=[A-Z]{3,})", table_text)

    for chunk in chunks:

        chunk = chunk.strip()
        if len(chunk) < 10:
            continue

        # must contain numbers
        nums = re.findall(r"\d+\.\d+|\d+", chunk)

        if len(nums) < 2:
            continue

        # remove numbers → description
        desc = re.sub(r"\d+\.\d+|\d+", "", chunk)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        # ignore totals/noise
        if any(x in desc for x in ["المجموع", "الإحمالي", "الرصيد"]):
            continue

        # your structure: last 2 numbers = qty + price
        quantity = nums[-2]
        unit_price = nums[-1]

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": unit_price
        })

    return pd.DataFrame(rows)

# =========================
# process PDF
# =========================

def process_pdf(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc)

    if len(text.strip()) < 50:
        text = ocr_pdf(pdf_path)

    return text, extract_table(text)

# =========================
# STREAMLIT APP
# =========================

st.set_page_config(page_title="Invoice Table Extractor", layout="wide")
st.title("📄 Invoice Table Extractor (Fixed for Your OCR Format)")

files = st.file_uploader(
    "Upload PDF / ZIP",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

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

            st.write(f"📄 Processing: {pdf.name}")

            text, df = process_pdf(pdf)

            debug[pdf.name] = text

            if df.empty:
                df = pd.DataFrame([{
                    "SKU / Description": "⚠️ No table detected",
                    "Quantity": "",
                    "Unit Price": ""
                }])

            df["Source File"] = pdf.name

            all_data.append(df)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show Extracted Text"):
            for name, txt in debug.items():
                st.subheader(name)
                st.text_area("text", txt, height=300)

        # =========================
        # OUTPUT
        # =========================

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Table Extracted Successfully")

        st.dataframe(final_df)

        buffer = BytesIO()
        final_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Download Excel",
            buffer,
            file_name="invoice_table.xlsx"
        )
