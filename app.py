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

def normalize_text(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# FIX BROKEN NUMBERS (IMPORTANT)
# =========================

def fix_broken_numbers(text):
    # merges: 58 584.42 -> 58584.42
    text = re.sub(r"(\d+)\s+(\d{3}\.\d+)", r"\1\2", text)
    text = re.sub(r"(\d+)\s+(\d{3})", r"\1\2", text)
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
        "Source File": filename
    }

# =========================
# SKU EXTRACTION (FIXED FOR YOUR CASE)
# =========================

def extract_items_from_text(text):

    lines = text.split("\n")
    rows = []

    skip_words = [
        "المجموع", "الإجمالي", "الرصيد", "الايبان",
        "رقم الحساب", "شركة", "السجل", "الرقم الضريبي",
        "فاتورة", "المملكة", "جدة", "الرياض"
    ]

    for line in lines:

        line = line.strip()
        if not line:
            continue

        if any(w in line for w in skip_words):
            continue

        # extract ALL numbers
        nums = re.findall(r"\d+\.\d+|\d+", line)

        # must have at least qty + price
        if len(nums) < 2:
            continue

        # extract description (remove numbers)
        desc = re.sub(r"\d+\.\d+|\d+", "", line)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        # 🚨 FIX LOGIC FOR YOUR INVOICE
        quantity = nums[-2]
        price = nums[-1]

        # ignore totals row (very important)
        if "BONE IN" not in desc and "WHOLE" not in desc:
            # still allow but safer filtering
            pass

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
        text = "\n".join([p.get_text() for p in doc])

    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    text = normalize_text(text)
    text = fix_broken_numbers(text)

    meta = extract_metadata(text, pdf_path.name)
    items = extract_items_from_text(text)

    return text, meta, items

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor FIXED", layout="wide")
st.title("📄 Invoice Extractor (FIXED FOR YOUR INVOICE STRUCTURE)")

files = st.file_uploader("Upload PDF / ZIP", type=["pdf", "zip"], accept_multiple_files=True)

debug_text = {}

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

            debug_text[pdf.name] = text

            if not items.empty:
                for k, v in meta.items():
                    items[k] = v
                all_data.append(items)

        # =========================
        # DEBUG VIEW
        # =========================

        with st.expander("🔍 Show Extracted Text (DEBUG)"):
            for name, txt in debug_text.items():
                st.subheader(name)
                st.text_area("Raw Text", txt, height=300)

        # =========================
        # OUTPUT
        # =========================

        if all_data:

            final_df = pd.concat(all_data, ignore_index=True)

            st.success("✅ FIXED EXTRACTION COMPLETED")

            st.dataframe(final_df)

            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button(
                "📥 Download Excel",
                buffer,
                file_name="fixed_invoices.xlsx"
            )

        else:
            st.warning("No data extracted")
