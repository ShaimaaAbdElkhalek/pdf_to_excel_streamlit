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

def ocr(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    return "\n".join(pytesseract.image_to_string(img, lang="ara+eng") for img in images)

# =========================
# CLEAN
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# 🔥 FINAL ROBUST EXTRACTION ENGINE
# =========================

def extract_items(text):

    text = clean(text)

    rows = []

    # STEP 1: find ALL numbers with positions
    numbers = [(m.group(), m.start()) for m in re.finditer(r"\d+\.\d+|\d+", text)]

    # STEP 2: cluster numbers into groups of 3 (price, qty, price/total mix)
    for i in range(len(numbers) - 2):

        n1, p1 = numbers[i]
        n2, p2 = numbers[i + 1]
        n3, p3 = numbers[i + 2]

        # filter noise (IBAN, totals)
        if len(n1) > 6 or len(n2) > 6 or len(n3) > 6:
            continue

        # STEP 3: extract surrounding context window
        start = max(0, p1 - 150)
        end = min(len(text), p3 + 150)

        window = text[start:end]

        # must contain letters
        if not any(c.isalpha() for c in window):
            continue

        # STEP 4: clean description
        desc = re.sub(r"\d+\.\d+|\d+", "", window)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        # STEP 5: remove invoice noise
        if any(x in desc for x in ["المجموع", "الإحمالي", "الرصيد", "الايبان"]):
            continue

        # STEP 6: assign values
        quantity = n2
        unit_price = n3

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": unit_price
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df = df.drop_duplicates()

    return df

# =========================
# PROCESS PDF
# =========================

def process(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc)

    if len(text.strip()) < 50:
        text = ocr(pdf_path)

    return text, extract_items(text)

# =========================
# STREAMLIT APP
# =========================

st.set_page_config(page_title="Invoice AI FINAL", layout="wide")
st.title("📄 Invoice Extractor (FINAL ROBUST ENGINE)")

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

            st.write(f"📄 {pdf.name}")

            text, df = process(pdf)

            debug_store[pdf.name] = text

            # FORCE fallback if empty (never fail)
            if df.empty:
                df = pd.DataFrame([{
                    "SKU / Description": "⚠️ Manual review needed",
                    "Quantity": "",
                    "Unit Price": ""
                }])

            df["Source File"] = pdf.name

            all_data.append(df)

        # =========================
        # DEBUG TEXT
        # =========================

        with st.expander("🔍 Raw OCR Text"):
            for k, v in debug_store.items():
                st.subheader(k)
                st.text_area("text", v, height=300)

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
            file_name="invoice_final.xlsx"
        )
