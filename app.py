# streamlit_app.py

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import arabic_reshaper
from bidi.algorithm import get_display
import unicodedata

# =========================
# Arabic Helpers
# =========================
def reshape_arabic_text(text):
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except:
        return str(text)

# remove bidi control chars that break matching
_BIDI_CHARS_RE = re.compile(r"[\u200e\u200f\u202a-\u202e\u2066-\u2069]")

def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u00a0", " ")          # NBSP -> space
    text = _BIDI_CHARS_RE.sub("", text)         # remove bidi controls
    text = re.sub(r"[ \t]+", " ", text)         # collapse spaces
    text = re.sub(r"\n{2,}", "\n", text).strip()
    return text

# =========================
# Robust extraction helpers
# =========================
_AMOUNT_RE = re.compile(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)")

def first_amount_in(text: str) -> str:
    if not text:
        return ""
    m = _AMOUNT_RE.search(text)
    return m.group(1) if m else ""

def find_amount_near_keywords(text: str, keywords, window=160) -> str:
    """
    Search around keyword(s) and return first numeric amount near them.
    Works with: الرصيد المستحق元1,153.74  and  المستحق الرصيد元 58,500.49
    """
    text = normalize_text(text)
    if isinstance(keywords, str):
        keywords = [keywords]
    # allow flexible spaces inside keyword (e.g. "العميل اسم" vs "العميل   اسم")
    keyword_patterns = [re.sub(r"\s+", r"\\s*", re.escape(normalize_text(k))) for k in keywords]

    for kw_pat in keyword_patterns:
        for m in re.finditer(kw_pat, text):
            start = max(0, m.start() - window)
            end = min(len(text), m.end() + window)
            chunk = text[start:end]
            num = first_amount_in(chunk)
            if num:
                return num
    return ""

def extract_customer_name(full_text: str) -> str:
    """
    Supports BOTH old + new:
    - "اسم العميل : X"
    - "X : اسم العميل"
    - "العميل اسم : X"
    - "X : العميل اسم"
    - Also appends next line if name is split (like "الغذائية")
    """
    text = normalize_text(full_text)
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    label = r"(?:اسم\s*العميل|العميل\s*اسم)"

    # 1) Try line-by-line to handle multi-line continuation
    for i, line in enumerate(lines):
        # key : value
        m = re.search(label + r"\s*[:：]\s*(.+)", line)
        if m:
            name = m.group(1).strip()
            name = re.split(r"الرقم\s*الضريبي|رقم\s*السجل|العنوان", name)[0].strip()
            return name

        # value : key  (this is your NEW pdf style)
        m = re.search(r"(.+?)\s*[:：]\s*" + label, line)
        if m:
            name_parts = [m.group(1).strip()]

            # append next line(s) if they look like continuation (e.g., الغذائية)
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = lines[j]
                if re.search(r"(الرقم\s*الضريبي|رقم\s*السجل|العنوان|\d{6,})", nxt):
                    break
                if len(nxt) <= 40:
                    name_parts.append(nxt)
                else:
                    break

            name = " ".join(name_parts).strip()
            name = re.split(r"الرقم\s*الضريبي|رقم\s*السجل|العنوان", name)[0].strip()
            return name

    # 2) Fallback: after "فاتورة ضريبية" pick next 1-2 meaningful lines
    for i, line in enumerate(lines):
        if re.search(r"فاتورة\s*ضريبية", line):
            picked = []
            for j in range(i + 1, min(i + 6, len(lines))):
                nxt = lines[j]
                if re.search(r"(الرقم\s*الضريبي|رقم\s*السجل|العنوان|\d{6,})", nxt):
                    break
                # avoid taking generic headings
                if re.search(r"^(البند|الوصف|العدد|سعر|الكمية|المجموع)$", nxt):
                    break
                picked.append(nxt)
                if len(picked) == 2:
                    break
            if picked:
                return " ".join(picked).strip()

    return ""

# =========================
# Metadata Extraction (PyMuPDF)
# =========================
def extract_metadata(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text("text") for page in doc])

        full_text = normalize_text(full_text)

        def find_field(text, keywords):
            """
            key : value OR value : key (single-line)
            """
            text = normalize_text(text)
            if isinstance(keywords, str):
                keywords = [keywords]

            # flexible keyword whitespace
            keyword_patterns = [re.sub(r"\s+", r"\\s*", re.escape(normalize_text(k))) for k in keywords]

            for kw_pat in keyword_patterns:
                # key : value
                p1 = rf"{kw_pat}\s*[:：]?\s*([^\n]+)"
                m1 = re.search(p1, text)
                if m1 and m1.group(1).strip():
                    return m1.group(1).strip()

                # value : key
                p2 = rf"([^\n:：]+)\s*[:：]\s*{kw_pat}"
                m2 = re.search(p2, text)
                if m2 and m2.group(1).strip():
                    return m2.group(1).strip()

            return ""

        address_part1 = find_field(full_text, ["رقم السجل", "السجل رقم"])
        address_part2 = find_field(full_text, ["العنوان"])
        full_address = f"{address_part1} {address_part2}".strip()

        customer_name = extract_customer_name(full_text)

        invoice_number = find_field(full_text, ["رقم الفاتورة", "الفاتورة رقم"])
        invoice_date = find_field(full_text, ["تاريخ الفاتورة", "الفاتورة تاريخ"])

        paid = find_amount_near_keywords(full_text, ["مدفوع"])
        balance = find_amount_near_keywords(full_text, ["الرصيد المستحق", "المستحق الرصيد", "الرصيد"])

        metadata = {
            "Invoice Number": invoice_number,
            "Invoice Date": invoice_date,
            "Customer Name": customer_name,
            "Address": full_address,
            "Paid": paid,
            "Balance": balance,
            "Source File": pdf_path.name
        }
        return metadata

    except Exception as e:
        st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
        return {}

# =========================
# Table Extraction (pdfplumber)
# =========================
def is_data_row(row):
    return any(
        str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit()
        for cell in row
    )

def fix_shifted_rows(row):
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def extract_tables(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_data = []
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for table in tables:
                    if not table:
                        continue

                    df = pd.DataFrame(table)
                    df = df.dropna(how="all").reset_index(drop=True)
                    if df.empty:
                        continue

                    merged_rows = []
                    temp_row = []

                    for _, row in df.iterrows():
                        row_values = row.fillna("").astype(str).tolist()
                        row_values = [reshape_arabic_text(cell) for cell in row_values]
                        row_values = fix_shifted_rows(row_values)

                        if is_data_row(row_values):
                            if temp_row:
                                combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
                                merged_rows.append(combined)
                                temp_row = []
                            else:
                                merged_rows.append(row_values)
                        else:
                            temp_row = row_values

                    if merged_rows:
                        num_cols = len(merged_rows[0])
                        headers = ["Total before tax", "الكمية", "Unit price", "Quantity", "Description", "SKU", "إضافي"]
                        df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                        all_data.append(df_cleaned)

            return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

    except Exception as e:
        st.error(f"❌ Error extracting table from {pdf_path.name}: {e}")
        return pd.DataFrame()

# =========================
# Main Process Function
# =========================
def process_pdf(pdf_path):
    metadata = extract_metadata(pdf_path)
    table_data = extract_tables(pdf_path)

    if not table_data.empty:
        for key, value in metadata.items():
            table_data[key] = value
        return table_data
    else:
        return pd.DataFrame([metadata])

# =========================
# Streamlit App UI
# =========================
st.set_page_config(page_title="Merged Arabic Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor Pdf to Excel")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf", "zip"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        pdf_paths = []

        for uploaded_file in uploaded_files:
            file_path = temp_dir / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())

            if uploaded_file.name.lower().endswith(".zip"):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                for pdf in temp_dir.glob("*.pdf"):
                    pdf_paths.append(pdf)
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

            # ======== Cleaning Steps ========
            if "Total before tax" in final_df.columns:
                final_df["Total before tax"] = (
                    final_df["Total before tax"].astype(str)
                    .str.replace(r"[^\d.,]", "", regex=True)
                    .str.replace(",", "", regex=False)
                    .replace("", None)
                )
                final_df["Total before tax"] = pd.to_numeric(final_df["Total before tax"], errors="coerce")
                final_df["VAT 15%"] = (final_df["Total before tax"] * 0.15).round(2)
                final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15%"]).round(2)

            for col in ["Paid", "Balance"]:
                if col in final_df.columns:
                    final_df[col] = (
                        final_df[col].astype(str)
                        .str.replace(r"[^\d.,]", "", regex=True)
                        .str.replace(",", "", regex=False)
                        .replace("", None)
                    )
                    final_df[col] = pd.to_numeric(final_df[col], errors="coerce")

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            required_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Balance", "Paid", "Address",
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Description", "SKU",
                "Source File"
            ]
            for c in required_columns:
                if c not in final_df.columns:
                    final_df[c] = None
            final_df = final_df.reindex(columns=required_columns)

            st.success("✅ Extraction & cleaning complete!")
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
