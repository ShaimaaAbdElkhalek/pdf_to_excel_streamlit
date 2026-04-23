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
# CLEAN LIGHTLY (NO MEANING LOSS)
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# RAW TABLE EXTRACTION (KEY PART)
# =========================

def extract_raw_table(text):

    text = clean(text)

    # split into chunks using line breaks OR pipe
    raw_rows = re.split(r"[\n]", text)

    table = []

    for row in raw_rows:

        row = row.strip()
        if len(row) < 5:
            continue

        # keep EVERYTHING (no semantic filtering)

        # detect if row contains any numbers
        if not re.search(r"\d", row):
            continue

        # normalize spacing
        row = re.sub(r"\s+", " ", row)

        table.append(row)

    return pd.DataFrame({"RAW_ROW": table})

# =========================
# PROCESS
# =========================

def process(pdf_path):

    text = ""

    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc)

    if len(text.strip()) < 50:
        text = ocr(pdf_path)

    return text, extract_raw_table(text)

# =========================
# STREAMLIT
# =========================

st.set_page_config(page_title="Raw Table Extractor", layout="wide")
st.title("📄 RAW TABLE Extractor (No Interpretation Mode)")

files = st.file_uploader("Upload PDF / ZIP", type=["pdf", "zip"], accept_multiple_files=True)

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

        all_tables = []

        for pdf in pdfs:

            st.write(f"📄 {pdf.name}")

            text, df = process(pdf)

            df["SOURCE"] = pdf.name

            all_tables.append(df)

        final_df = pd.concat(all_tables, ignore_index=True)

        st.success("✅ RAW TABLE EXTRACTED")

        st.dataframe(final_df)

        buffer = BytesIO()
        final_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 Download Raw Table",
            buffer,
            file_name="raw_table.xlsx"
        )
