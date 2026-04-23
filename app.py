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

def clean_text(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# FIX BROKEN OCR NUMBERS
# =========================

def fix_numbers(text):
    # 58,584.42 or 58 584.42 → 58584.42
    text = re.sub(r"(\d)\s+(\d{3}\.\d+)", r"\1\2", text)
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
        "Tax Number": find("الرقم الضريبي"),
        "Source File": filename
    }

# =========================
# PRODUCT EXTRACTION (ROBUST)
# =========================

def extract_items(text):

    text = clean_text(text)
    text = fix_numbers(text)

    rows = []

    # isolate product section
    match = re.search(r"العدد(.*?)المجموع", text)
    product_block = match.group(1) if match else text

    # split by capitalized product names OR Arabic+English mix
    parts = re.split(r"(?=[A-Z][A-Z\s]{3,})", product_block)

    for p in parts:

        p = p.strip()
        if len(p) < 10:
            continue

        nums = re.findall(r"\d+\.\d+|\d+", p)

        if len(nums) < 2:
            continue

        # remove numbers for description
        desc = re.sub(r"\d+\.\d+|\d+", "", p)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        quantity = nums[-2]
        price = nums[-1]

        # filter obvious totals
        if "المجموع" in desc or "الإحمالي" in desc:
            continue

        rows.append({
            "SKU / Description": desc,
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
        text = "\n".join([page.get_text() for page in doc])

    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items(text)

    return text, meta, items

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor Pro", layout="wide")
st.title("📄 Invoice Extractor PRO (Final Stable Version)")

files = st.file_uploader("Upload PDF / ZIP", type=["pdf", "zip"], accept_multiple_files=True)

debug_data = {}

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

            debug_data[pdf.name] = text

            if not items.empty:
                for k, v in meta.items():
                    items[k] = v
                all_data.append(items)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show Extracted Text (DEBUG)"):
            for name, txt in debug_data.items():
                st.subheader(name)
                st.text_area("Raw Text", txt, height=300)

        # =========================
        # OUTPUT
        # =========================

        if all_data:

            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ Extraction Successful")

            st.dataframe(final_df)

            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="invoices_clean.xlsx"
            )

        else:
            st.error("❌ No products extracted — invoice format needs deeper layout parsing")
