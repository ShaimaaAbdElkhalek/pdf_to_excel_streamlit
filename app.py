# streamlit_app.py
# Corrected to extract Customer Name for BOTH old + new PDFs:
# - Old: "مؤسسة طعام المحترف: اسم العميل"
# - New: "مؤسسة الركن الخليجي : العميل اسم" + sometimes next line "الغذائية"
# Also:
# - Invoice Number: "رقم الفاتورة" OR "الفاتورة رقم"
# - Normalizes Arabic presentation forms using NFKC
# - Keeps your table logic as-is

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
        bidi_text = get_display(reshaped)
        return bidi_text
    except:
        return str(text)

def normalize_text(text: str) -> str:
    if text is None:
        return ""
    return unicodedata.normalize("NFKC", str(text))

def clean_money_str(x: str) -> str:
    x = normalize_text(x)
    x = re.sub(r"[^\d.,]", "", x)
    x = x.replace(",", "")
    return x.strip()

def first_amount_near(text: str, keywords, window=80) -> str:
    text = normalize_text(text)
    if isinstance(keywords, str):
        keywords = [keywords]
    keywords = [normalize_text(k) for k in keywords]

    for kw in keywords:
        for m in re.finditer(re.escape(kw), text):
            start = max(0, m.start() - window)
            end = min(len(text), m.end() + window)
            chunk = text[start:end]
            num = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)", chunk)
            if num:
                return num.group(1)
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
            Supports BOTH:
              1) key : value
              2) value : key
            Returns first matched 'value string' (not numeric-cleaned).
            """
            text = normalize_text(text)

            if isinstance(keywords, str):
                keywords = [keywords]
            keywords = [normalize_text(k) for k in keywords]

            for keyword in keywords:
                # 1) key : value
                p1 = rf"{re.escape(keyword)}\s*[:：]?\s*([^\n]+)"
                m1 = re.search(p1, text)
                if m1 and m1.group(1).strip():
                    return m1.group(1).strip()

                # 2) value : key
                p2 = rf"([^\n:：]+)\s*[:：]\s*{re.escape(keyword)}"
                m2 = re.search(p2, text)
                if m2 and m2.group(1).strip():
                    return m2.group(1).strip()

            return ""

        # ---------- Address ----------
        address_part1 = find_field(full_text, ["رقم السجل", "السجل رقم"])
        address_part2 = find_field(full_text, ["العنوان"])
        full_address = f"{address_part1} {address_part2}".strip()

        # ---------- Customer Name (FIXED) ----------
        # We try to capture customer name from BOTH directions and BOTH labels:
        # "اسم العميل" / "العميل اسم"
        # Also handle the case where "الغذائية" comes in the next line.
        customer_name = ""

        # Try patterns on the WHOLE text (more reliable than depending on find_field only)
        patterns = [
            r"(?:اسم العميل|العميل اسم)\s*[:：]\s*([^\n]+)",          # key : value
            r"([^\n:：]+?)\s*[:：]\s*(?:اسم العميل|العميل اسم)",      # value : key
        ]
        for p in patterns:
            m = re.search(p, full_text)
            if m:
                customer_name = m.group(1).strip()
                break

        # If still empty, fallback to find_field (sometimes it is enough)
        if not customer_name:
            customer_line = find_field(full_text, ["اسم العميل", "العميل اسم"])
            customer_line = normalize_text(customer_line)

            m1 = re.search(r"(?:اسم العميل|العميل اسم)\s*[:：]\s*(.+)", customer_line)
            if m1:
                customer_name = m1.group(1).strip()
            else:
                m2 = re.search(r"(.+?)\s*[:：]\s*(?:اسم العميل|العميل اسم)", customer_line)
                customer_name = m2.group(1).strip() if m2 else customer_line.strip()

        # Append next line if the PDF splits like:
        # "مؤسسة الركن الخليجي : العميل اسم"
        # "الغذائية"
        if customer_name:
            # find where this customer chunk appears, then look at the next line
            idx = full_text.find(customer_name)
            if idx != -1:
                tail = full_text[idx: idx + 200]
                lines = [l.strip() for l in tail.splitlines() if l.strip()]
                if lines:
                    # if first line is customer_name and second line is a short continuation, append it
                    if len(lines) >= 2 and lines[0] == customer_name:
                        cont = lines[1]
                        # avoid appending if continuation looks like tax/registry/address
                        if not re.search(r"(الرقم الضريبي|رقم السجل|العنوان|\d{6,})", cont):
                            # often continuation is "الغذائية"
                            if len(cont) <= 40:
                                customer_name = f"{customer_name} {cont}".strip()

        # final cleanup
        customer_name = re.split(r"الرقم الضريبي|رقم السجل|العنوان", customer_name)[0].strip()

        # ---------- Invoice Number/Date ----------
        invoice_number = find_field(full_text, ["رقم الفاتورة", "الفاتورة رقم"])
        invoice_date = find_field(full_text, ["تاريخ الفاتورة", "الفاتورة تاريخ"])

        # ---------- Paid/Balance (more reliable numeric extraction) ----------
        paid = first_amount_near(full_text, ["مدفوع"])
        balance = first_amount_near(full_text, ["الرصيد المستحق", "المستحق الرصيد", "الرصيد"])

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
                    if table:
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

            # ======== Fix Invoice Date to MM/DD/YYYY ========
            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            # ======== Keep only required columns in order ========
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
