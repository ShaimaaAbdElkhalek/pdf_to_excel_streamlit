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
# CLEAN
# =========================

def clean(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# EXTRACT METADATA
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
# 🔥 CORE FIX (PATTERN-BASED EXTRACTION)
# =========================

def extract_items(text):

    text = clean(text)

    rows = []

    # STEP 1: isolate product area
    try:
        block = re.search(r"العدد(.*?)المجموع", text).group(1)
    except:
        block = text

    # STEP 2: find product chunks using keyword anchor (English names)
    product_chunks = re.split(r"(?=[A-Z]{3,})", block)

    for chunk in product_chunks:

        chunk = chunk.strip()
        if len(chunk) < 10:
            continue

        # STEP 3: extract all numbers
        nums = re.findall(r"\d+\.\d+|\d+", chunk)

        if len(nums) < 2:
            continue

        # STEP 4: detect product name (keep Arabic + English)
        desc = re.sub(r"\d+\.\d+|\d+", "", chunk)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        # STEP 5: assign smart values
        quantity = nums[-2]
        price = nums[-1]

        # ignore totals
        if "المجموع" in desc or "الإحمالي" in desc:
            continue

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
# STREAMLIT
# =========================

st.set_page_config(page_title="Invoice AI Fix", layout="wide")
st.title("📄 Invoice Extractor (AI Pattern Fix - Works on Your Invoice)")

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

            if not items.empty:
                for k, v in meta.items():
                    items[k] = v
                all_data.append(items)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show Raw OCR Text"):
            for k, v in debug.items():
                st.subheader(k)
                st.text_area("text", v, height=300)

        # =========================
        # OUTPUT
        # =========================

        if all_data:

            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ Extracted Successfully")

            st.dataframe(final_df)

            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="invoice_ai_fixed.xlsx"
            )

        else:
            st.error("❌ Still no extraction — invoice is fully unstructured OCR")
