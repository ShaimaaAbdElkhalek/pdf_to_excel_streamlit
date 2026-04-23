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

def ocr_extract(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang="ara+eng") + "\n"
    return text

# =========================
# CLEAN
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# 🔥 CORE TABLE RECONSTRUCTION
# =========================

def extract_table(text):

    text = clean(text)

    rows = []

    # STEP 1: extract all numbers with positions
    nums = [(m.group(), m.start()) for m in re.finditer(r"\d{1,3}(?:,\d{3})*\.?\d*|\d+", text)]

    # STEP 2: slide through numbers
    for i in range(len(nums) - 2):

        n1, p1 = nums[i]
        n2, p2 = nums[i + 1]
        n3, p3 = nums[i + 2]

        # skip large totals / IBAN-like noise
        if len(n1) > 6 or len(n2) > 6 or len(n3) > 6:
            continue

        # STEP 3: capture context window
        start = max(0, p1 - 150)
        end = min(len(text), p3 + 150)

        window = text[start:end]

        # must contain product text
        if not any(c.isalpha() for c in window):
            continue

        # STEP 4: clean description
        desc = re.sub(r"\d{1,3}(?:,\d{3})*\.?\d*|\d+", "", window)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        # STEP 5: assign columns
        quantity = n2
        unit_price = n3

        # skip totals noise
        if "المجموع" in desc or "الإحمالي" in desc or "الرصيد" in desc:
            continue

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": unit_price
        })

    return pd.DataFrame(rows).drop_duplicates()

# =========================
# PDF PROCESS
# =========================

def process(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc)

    if len(text.strip()) < 50:
        text = ocr_extract(pdf_path)

    return text, extract_table(text)

# =========================
# STREAMLIT APP
# =========================

st.set_page_config(page_title="Invoice Extractor FINAL", layout="wide")
st.title("📄 Invoice Extractor (FINAL WORKING VERSION)")

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

            text, df = process(pdf)

            debug_store[pdf.name] = text

            # fallback if empty
            if df.empty:
                df = pd.DataFrame([{
                    "SKU / Description": "⚠️ No table detected",
                    "Quantity": "",
                    "Unit Price": ""
                }])

            all_data.append(df)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show OCR Text"):
            for name, txt in debug_store.items():
                st.subheader(name)
                st.text_area("text", txt, height=300)

        # =========================
        # OUTPUT
        # =========================

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("✅ Extraction Completed")

        st.dataframe(final_df)

        buffer = BytesIO()
        final_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Download Excel",
            buffer,
            file_name="invoice_table.xlsx"
        )
