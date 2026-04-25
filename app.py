import streamlit as st
import tempfile
import pandas as pd
import re
from pdf2image import convert_from_path
import pytesseract
from io import BytesIO


# -----------------------------
# CLEAN TEXT
# -----------------------------
def clean_text(text):
    text = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-\+\n]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# -----------------------------
# OCR FUNCTION
# -----------------------------
def extract_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)

    config = r'--oem 3 --psm 6 -l ara+eng'

    full_text = []

    for img in images:
        text = pytesseract.image_to_string(img, config=config)
        text = clean_text(text)
        full_text.append(text)

    return "\n\n".join(full_text)


# -----------------------------
# STRUCTURE TEXT INTO TABLE
# -----------------------------
def to_table(text):
    rows = []

    for line in text.split("\n"):
        line = line.strip()

        if len(line) < 2:
            continue

        # split smartly
        cols = re.split(r'\t+|\s{2,}', line)

        cols = [c.strip() for c in cols if c.strip()]

        if len(cols) >= 3:
            rows.append(cols)

    if not rows:
        return pd.DataFrame()

    max_len = max(len(r) for r in rows)
    rows = [r + [""] * (max_len - len(r)) for r in rows]

    return pd.DataFrame(rows)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="OCR Invoice Tool", layout="wide")

st.title("📄 OCR PDF → Text + Table + Excel")

uploaded_file = st.file_uploader("Upload OCR PDF", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        path = tmp.name

    with st.spinner("Processing OCR..."):
        text = extract_text(path)
        df = to_table(text)

    # ---------------- TEXT VIEW ----------------
    st.subheader("📜 Clean OCR Text")
    st.text_area("Text Output", text, height=300)

    # ---------------- TABLE VIEW ----------------
    st.subheader("📊 Structured Data")
    if df.empty:
        st.warning("Could not structure data properly 😢")
    else:
        st.dataframe(df)

    # ---------------- DOWNLOAD EXCEL ----------------
    if not df.empty:
        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            "📥 Download Excel",
            data=output,
            file_name="ocr_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
