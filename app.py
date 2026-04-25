import streamlit as st
import re
import unicodedata

import arabic_reshaper
from bidi.algorithm import get_display


# -----------------------------
# FIX BROKEN ARABIC (IMPORTANT PART)
# -----------------------------
def fix_arabic_text(text):
    # 1. Normalize unicode
    text = unicodedata.normalize("NFKC", text)

    # 2. Remove weird OCR artifacts
    text = re.sub(r'[^\w\s\u0600-\u06FF%:.,()/\-+]', ' ', text)

    # 3. Fix Arabic shaping + direction
    try:
        text = arabic_reshaper.reshape(text)
        text = get_display(text)
    except:
        pass

    # 4. Clean spaces
    text = re.sub(r'\s+', ' ', text).strip()

    return text


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.set_page_config(page_title="Fix Arabic OCR Text", layout="wide")

st.title("🔧 إصلاح النص العربي المشوه")

input_text = st.text_area("Paste your broken OCR text here", height=300)

if input_text:
    fixed = fix_arabic_text(input_text)

    st.subheader("✅ Fixed Text")
    st.text_area("", fixed, height=400)

    st.download_button(
        "📥 Download Fixed Text",
        fixed,
        file_name="fixed_text.txt"
    )
