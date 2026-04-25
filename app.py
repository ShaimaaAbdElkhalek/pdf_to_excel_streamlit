import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import tempfile
import re
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


# ----------------------------
# Extract text from PDF page
# ----------------------------
def extract_text_from_page(page):
    text = page.extract_text()
    if text:
        return text

    # OCR fallback using image rendering
    pix = page.to_pixmap()
    img_bytes = pix.tobytes("png")

    try:
        from PIL import Image
        import pytesseract

        img = Image.open(BytesIO(img_bytes))
        text = pytesseract.image_to_string(img)
        return text
    except:
        return ""


# ----------------------------
# Process PDF
# ----------------------------
def process_pdf(pdf_file):
    all_data = []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_file.read())
        tmp_path = tmp.name

    doc = fitz.open(tmp_path)

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = extract_text_from_page(page)

        if not text:
            continue

        text = reshape(text)

        # Clean lines
        lines = [line.strip() for line in text.split("\n") if line.strip()]

        for line in lines:
            # simple split logic (you can customize later)
            cols = re.split(r"\s{2,}", line)
            all_data.append(cols)

    # Normalize rows length
    max_len = max((len(row) for row in all_data), default=0)
    cleaned = [row + [""] * (max_len - len(row)) for row in all_data]

    df = pd.DataFrame(cleaned)
    return df


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="PDF to Excel OCR", layout="wide")

st.title("📄 PDF to Excel Converter (OCR + Arabic Support)")

uploaded_file = st.file_uploader("Upload your PDF", type=["pdf"])

if uploaded_file:
    with st.spinner("Processing PDF..."):
        df = process_pdf(uploaded_file)

    st.success("Done!")

    st.dataframe(df)

    # Download Excel
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=output,
        file_name="output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
