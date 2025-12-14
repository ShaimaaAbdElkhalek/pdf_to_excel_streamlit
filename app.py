# streamlit_app.py
# Arabic Invoice Extractor (Old + New PDFs) - Order Independent + No table length mismatch

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import unicodedata

# =========================
# TEXT NORMALIZATION
# =========================
def normalize_text(text):
    """
    - Converts Arabic presentation forms (ﻓﺎﺗﻮرة / اﻟﻔﺎﺗﻮرة...) -> normal Arabic (فاتورة / الفاتورة...)
    - Normalizes spaces
    - Normalizes Arabic decimal separators
    """
    if text is None:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("٫", ".").replace("٬", ",")
    text = re.sub(r"\s+", " ", text).strip()
    return text

def clean_money(x):
    x = normalize_text(x)
    x = re.sub(r"[^\d\.,]", "", x)
    x = x.replace(",", "")
    return x.strip()

def to_float_safe(x):
    s = clean_money(x)
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None

# =========================
# ORDER-INDEPENDENT FIND
# =========================
def find_by_tokens(text, tokens, value_pattern, window_size=60):
    """
    Find a value (matching value_pattern) that appears near the given tokens (any order).
    tokens: list[str]
    value_pattern: regex capturing group 1 as the value
    """
    text = normalize_text(text)

    for m in re.finditer(value_pattern, text):
        start, end = m.span()
        window = text[max(0, start - window_size): min(len(text), end + window_size)]
        if all(tok in window for tok in tokens):
            return m.group(1).strip()
    return ""

def find_field(text, keywords):
    """
    Fallback: match label then value (same line or next line).
    keywords can be a string or list[str].
    """
    text = text or ""
    if isinstance(keywords, str):
        keywords = [keywords]

    for key in keywords:
        p1 = rf"{re.escape(key)}\s*[:：]?\s*([^\n]+)"
        m = re.search(p1, text)
        if m and m.group(1).strip():
            return m.group(1).strip()

        p2 = rf"{re.escape(key)}\s*[:：]?\s*\n\s*([^\n]+)"
        m2 = re.search(p2, text)
        if m2 and m2.group(1).strip():
            return m2.group(1).strip()

    return ""

# =========================
# READ FULL TEXT (fitz + fallback)
# =========================
def read_full_text(pdf_path: Path) -> str:
    txt = ""
    try:
        with fitz.open(pdf_path) as doc:
            txt = "\n".join([p.get_text("text") for p in doc])
    except Exception:
        txt = ""

    if not txt or len(txt.strip()) < 50:
        try:
            with pdfplumber.open(pdf_path) as pdf:
                txt = "\n".join([(p.extract_text() or "") for p in pdf.pages])
        except Exception:
            txt = txt or ""

    return normalize_text(txt)

# =========================
# METADATA EXTRACTION
# =========================
def extract_metadata(pdf_path: Path) -> dict:
    try:
        text = read_full_text(pdf_path)

        # Invoice number/date: order-independent
        invoice_number = find_by_tokens(text, ["فاتورة", "رقم"], r"\b(\d{4,6})\b")
        invoice_date   = find_by_tokens(text, ["فاتورة", "تاريخ"], r"(\d{2}/\d{2}/\d{4})")

        # Customer name: try label-based (more reliable for names)
        customer_name = find_field(text, ["اسم العميل", "العميل اسم"])
        customer_name = re.split(r"الرقم الضريبي|رقم السجل|العنوان", customer_name)[0].strip()

        # Address
        address = find_field(text, ["العنوان"])
        address = address.strip()

        # Paid / Balance (some new PDFs reorder words)
        paid = find_by_tokens(text, ["مدفوع"], r"مدفوع\s*([0-9\.,]+)")
        balance = (
            find_by_tokens(text, ["الرصيد", "المستحق"], r"الرصيد\s*المستحق\s*([0-9\.,]+)")
            or find_by_tokens(text, ["المستحق", "الرصيد"], r"المستحق\s*الرصيد\s*([0-9\.,]+)")
        )

        return {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": address,
            "Paid": clean_money(paid),
            "Balance": clean_money(balance),
            "Source File": pdf_path.name
        }

    except Exception as e:
        st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
        return {}

# =========================
# TABLE EXTRACTION (NO LENGTH MISMATCH)
# =========================
def extract_tables(pdf_path: Path) -> pd.DataFrame:
    """
    Bulletproof table extraction:
    - Accepts tables with variable column counts (4..7..etc)
    - Never assigns fixed df.columns length
    - Renames columns safely by index
    """
    try:
        all_rows = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        row = [normalize_text(c) for c in row if c is not None]
                        # keep rows containing any digit (likely data row)
                        if any(re.search(r"\d", c) for c in row):
                            all_rows.append(row)

        if not all_rows:
            return pd.DataFrame()

        # Normalize row length
        max_len = max(len(r) for r in all_rows)
        normalized = []
        for r in all_rows:
            if len(r) < max_len:
                r = r + [""] * (max_len - len(r))
            elif len(r) > max_len:
                r = r[:max_len]
            normalized.append(r)

        df = pd.DataFrame(normalized)

        # Rename only if column index exists
        rename_map = {
            0: "Description",
            1: "Quantity",
            2: "Unit price",
            3: "Total before tax",
        }
        for idx, name in rename_map.items():
            if idx in df.columns:
                df.rename(columns={idx: name}, inplace=True)

        # Ensure required cols exist
        for c in ["Description", "Quantity", "Unit price", "Total before tax"]:
            if c not in df.columns:
                df[c] = ""

        return df

    except Exception as e:
        st.error(f"❌ Table error in {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# PROCESS ONE PDF
# =========================
def process_pdf(pdf_path: Path) -> pd.DataFrame:
    meta = extract_metadata(pdf_path)
    table = extract_tables(pdf_path)

    if not table.empty:
        for k, v in meta.items():
            table[k] = v
        return table

    return pd.DataFrame([meta])

# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="Arabic Invoice Extractor (Order-Independent)", layout="wide")
st.title("📄 Arabic Invoice Extractor (Order-Independent)")

uploaded_files = st.file_uploader("Upload PDFs or ZIP", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        # Save uploads
        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name
            file_path.write_bytes(uploaded_file.read())

            if file_path.suffix.lower() == ".zip":
                with zipfile.ZipFile(file_path, "r") as z:
                    z.extractall(temp_dir)
                pdf_paths.extend(temp_dir.glob("*.pdf"))
            else:
                pdf_paths.append(file_path)

        all_data = []
        for pdf_path in pdf_paths:
            st.write(f"📄 Processing: {pdf_path.name}")
            df = process_pdf(pdf_path)
            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            # ===== Calculations (optional, same as your style) =====
            if "Total before tax" in final_df.columns:
                final_df["Total before tax"] = final_df["Total before tax"].apply(to_float_safe)
                final_df["VAT 15%"] = (final_df["Total before tax"].fillna(0) * 0.15).round(2)
                final_df["Total after tax"] = (final_df["Total before tax"].fillna(0) + final_df["VAT 15%"]).round(2)

            for col in ["Paid", "Balance"]:
                if col in final_df.columns:
                    final_df[col] = final_df[col].apply(to_float_safe)

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            required_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Balance", "Paid", "Address",
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Description",
                "Source File"
            ]

            # Ensure cols exist
            for c in required_columns:
                if c not in final_df.columns:
                    final_df[c] = None

            final_df = final_df.reindex(columns=required_columns)

            st.success("✅ Extraction complete")
            st.dataframe(final_df, use_container_width=True)

            output = BytesIO()
            final_df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)

            st.download_button(
                label="📥 Download Excel",
                data=output,
                file_name="Merged_Invoice_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No data extracted from the uploaded files.")
