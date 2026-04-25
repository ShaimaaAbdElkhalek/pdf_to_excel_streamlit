import streamlit as st
import tempfile
import re
from pdf2image import convert_from_path
import pytesseract


# -----------------------------
# ADVANCED CLEANING ENGINE
# -----------------------------
def is_noise(line):
    line = line.strip()

    # remove tiny garbage
    if len(line) < 3:
        return True

    # only numbers / symbols
    if re.fullmatch(r'[\d\s\-\.:,%()+]+', line):
        return True

    # OCR garbage patterns
    garbage_words = ["ee", "ae", "ob", "fay", "cece", "rates", "ta", "crs"]

    if any(g in line.lower() for g in garbage_words):
        if len(line.split()) < 5:
            return True

    return False


def clean_line(line):
    line = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-\+]', ' ', line)
    line = re.sub(r'\s+', ' ', line).strip()
    return line


# -----------------------------
# SMART OCR
# -----------------------------
def extract_text(pdf_path):
    images = convert_from_path(pdf_path, dpi=400)
    config = r'--oem 3 --psm 6 -l ara+eng'

    all_clean = []

    for img in images:
        text = pytesseract.image_to_string(img, config=config)

        lines = text.split("\n")

        for line in lines:
            line = clean_line(line)

            if not line:
                continue

            if is_noise(line):
                continue

            all_clean.append(line)

    return "\n".join(all_clean)


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="AI Invoice Cleaner", layout="wide")

st.title("📄 AI Invoice OCR Cleaner (Production Level)")

file = st.file_uploader("Upload PDF Invoice", type=["pdf"])

if file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        path = tmp.name

    with st.spinner("Processing invoice..."):

        result = extract_text(path)

    if not result.strip():
        st.error("No usable text found 😢")
    else:
        st.success("Invoice cleaned successfully ✅")

        st.text_area("📜 Clean Invoice Output", result, height=500)

        st.download_button(
            "📥 Download Clean Invoice",
            result,
            file_name="clean_invoice.txt",
            mime="text/plain"
        )
