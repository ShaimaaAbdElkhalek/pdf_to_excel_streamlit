import streamlit as st
import pandas as pd
import re


# -----------------------------
# CLEAN FUNCTION
# -----------------------------
def clean_cell(text):
    if not text:
        return ""

    # fix Arabic encoding noise
    text = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-\+\"]', ' ', text)

    # normalize spaces
    text = re.sub(r'\s+', ' ', text).strip()

    return text


# -----------------------------
# PARSE ROWS (TAB / OCR OUTPUT)
# -----------------------------
def parse_data(raw_text):
    rows = []

    for line in raw_text.split("\n"):
        line = line.strip()

        if not line:
            continue

        # split by tabs OR multiple spaces
        cols = re.split(r'\t+|\s{2,}', line)

        cols = [clean_cell(c) for c in cols if c.strip()]

        if len(cols) >= 5:  # filter valid rows only
            rows.append(cols)

    return rows


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Structured Invoice Cleaner", layout="wide")

st.title("📊 OCR → Structured Invoice Table")

input_text = st.text_area("Paste OCR raw data here", height=300)

if input_text:

    data = parse_data(input_text)

    if not data:
        st.error("No structured data found 😢")
    else:
        # normalize columns length
        max_len = max(len(r) for r in data)
        data = [r + [""] * (max_len - len(r)) for r in data]

        df = pd.DataFrame(data)

        st.success("Structured table created ✅")

        st.dataframe(df)

        # download excel
        excel = df.to_excel(index=False, engine="openpyxl")

        st.download_button(
            "📥 Download Excel",
            data=open,
            file_name="structured_invoice.xlsx"
        )
        
