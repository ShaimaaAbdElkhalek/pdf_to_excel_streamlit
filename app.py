import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
from io import BytesIO

# Optional Arabic support
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    def reshape(text):
        try:
            return get_display(arabic_reshaper.reshape(text))
        except:
            return text
except:
    def reshape(text):
        return text


# -----------------------------
# SAFE TEXT EXTRACTION
# -----------------------------
def extract_text_safe(page):
    try:
        text = page.extract_text()
        if text:
            return text
    except Exception:
        pass

    # fallback: try extracting words
    try:
        words = page.extract_words()
        if words:
            return " ".join([w["text"] for w in words])
    except Exception:
        pass

    return ""


# -----------------------------
# PROCESS PDF
# -----------------------------
def process_pdf(pdf_file):
    all_rows = []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    with pdfplumber.open(tmp_path) as pdf:
        for page in pdf.pages:
            text = extract_text_safe(page)

            if not text:
                continue

            text = reshape(text)

            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                cols = line.split()  # simple split (safe)
                all_rows.append(cols)

    if not all_rows:
        return pd.DataFrame()

    max_len = max(len(r) for r in all_rows)
    cleaned = [r + [""] * (max_len - len(r)) for r in all_rows]

    return pd.DataFrame(cleaned)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="PDF to Excel", layout="wide")

st.title("📄 PDF → Excel Converter (Stable Version)")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing..."):
        df = process_pdf(uploaded_file)

    if df.empty:
        st.error("No data found in PDF 😢")
    else:
        st.success("Done!")

        st.dataframe(df)

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            "📥 Download Excel",
            output,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
