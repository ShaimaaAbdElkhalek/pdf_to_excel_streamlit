import streamlit as st
import tempfile
import re
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from io import BytesIO


# -----------------------------
# OCR
# -----------------------------
def extract_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)
    config = r'--oem 3 --psm 6 -l ara+eng'

    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, config=config) + "\n"

    return text


# -----------------------------
# CLEAN NAME (FIX YOUR ISSUE)
# -----------------------------
def clean_name(text):
    text = re.sub(r'[^\u0600-\u06FFa-zA-Z\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # keep only first meaningful part
    return text.split("الفاتورة")[0].strip()


# -----------------------------
# HEADER EXTRACTION
# -----------------------------
def extract_header(text):

    def find(pattern):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else ""

    return {
        "Customer Name": clean_name(find(r"اسم العميل\s*[:\-]?\s*([^\n]+)")),
        "Invoice Number": find(r"الفاتورة\s*[:\-]?\s*(\d+)"),
        "VAT Number": find(r"الرقم الضريبي\s*[:\-]?\s*([0-9]+)"),
        "CR Number": find(r"رقم السجل\s*[:\-]?\s*([0-9\- ]+)"),
        "Phone": find(r"رقم الجوال\s*[:\-]?\s*([0-9]+)"),
        "Total": find(r"المجموع\s*([0-9,\.]+)"),
        "VAT Value": find(r"القيمة المضافة\s*([0-9,\.]+)"),
        "Grand Total": find(r"الإحمالي\s*([0-9,\.]+)"),
        "IBAN": find(r"(SA[0-9A-Z]{20,})")
    }


# -----------------------------
# LINE ITEMS EXTRACTOR (FIX FOR MULTI PRODUCTS)
# -----------------------------
def extract_items(text):

    items = []

    lines = text.split("\n")

    for line in lines:

        # detect product-like rows (heuristic)
        if re.search(r"\d{1,5}\.\d{1,2}|\d+\s+\d+", line):

            parts = re.split(r'\s{2,}|\t+', line)
            parts = [p.strip() for p in parts if p.strip()]

            if len(parts) >= 3:
                items.append(parts)

    if not items:
        return pd.DataFrame()

    max_len = max(len(r) for r in items)
    items = [r + [""] * (max_len - len(r)) for r in items]

    return pd.DataFrame(items)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Invoice AI PRO", layout="wide")

st.title("📄 Invoice AI PRO (Multi-invoice + Items support)")

file = st.file_uploader("Upload PDF Invoice", type=["pdf"])

if file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        path = tmp.name

    with st.spinner("Processing invoice..."):

        text = extract_text(path)

        header = extract_header(text)
        items_df = extract_items(text)

    # ---------------- HEADER ----------------
    st.subheader("📌 Invoice Header")
    st.json(header)

    # ---------------- ITEMS ----------------
    st.subheader("🛒 Items (Products)")

    if items_df.empty:
        st.warning("No items detected (invoice may be image-only or unstructured)")
    else:
        st.dataframe(items_df)

    # ---------------- DOWNLOAD ----------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([header]).to_excel(writer, sheet_name="Invoice", index=False)
        items_df.to_excel(writer, sheet_name="Items", index=False)

    output.seek(0)

    st.download_button(
        "📥 Download Excel (Structured Invoice)",
        data=output,
        file_name="invoice_pro.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
