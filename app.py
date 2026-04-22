import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
import arabic_reshaper
from bidi.algorithm import get_display

# ── Arabic helpers ─────────────────────────────────────────────
def reshape(text):
    try:
        return get_display(arabic_reshaper.reshape(text))
    except:
        return text

def clean_number(val):
    return re.sub(r"[^\d.]", "", str(val).replace(",", "").replace("٬", "").replace("٫", "."))

# ── Raw text extractor (for debugging) ────────────────────────
def get_raw_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        return "\n".join(page.get_text() for page in doc)

# ── Flexible field finder ──────────────────────────────────────
def find_after(text, *keywords):
    """Try multiple keywords, return first match found after any of them."""
    for kw in keywords:
        pattern = rf"{re.escape(kw)}\s*[:\-]?\s*([^\n]+)"
        m = re.search(pattern, text)
        if m:
            val = m.group(1).strip()
            if val:
                return val
    return ""

def find_first_number(text, *keywords):
    """Find the first standalone number after a keyword."""
    for kw in keywords:
        pattern = rf"{re.escape(kw)}[^\d\n]*([\d,٬٫\.]+)"
        m = re.search(pattern, text)
        if m:
            return clean_number(m.group(1))
    return ""

# ── Metadata extraction ────────────────────────────────────────
def extract_metadata(pdf_path):
    try:
        raw = get_raw_text(pdf_path)

        # All 15-digit numbers → tax numbers (supplier first, customer second)
        tax_numbers = re.findall(r"\b\d{15}\b", raw)
        # All 10-digit numbers → CR numbers
        cr_numbers  = re.findall(r"\b\d{10}\b", raw)

        # Invoice number: usually a short number near رقم الفاتورة
        inv_num = find_after(raw, "رقم الفاتورة")
        if not inv_num:
            # fallback: 4-5 digit standalone number
            m = re.search(r"\b(\d{4,6})\b", raw)
            inv_num = m.group(1) if m else ""

        # Date: look for DD/MM/YYYY or similar
        date_match = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", raw)
        inv_date = date_match.group(1) if date_match else find_after(raw, "تاريخ الفاتورة")

        # Customer name: text after اسم العميل
        cname_match = re.search(r"اسم العميل\s*[:\s]+(.+?)(?:\n)", raw)
        customer_name = cname_match.group(1).strip() if cname_match else ""

        # Address
        addr_match = re.search(r"العنوان\s*[:\s]+(.+?)(?:\n\d|\Z)", raw, re.DOTALL)
        address = addr_match.group(1).strip().replace("\n", "، ") if addr_match else ""

        # Financial totals
        balance  = find_first_number(raw, "الإجمالي", "اإلجمالي", "المجموع الكلي")
        paid     = find_first_number(raw, "مدفوع")
        not_paid = find_first_number(raw, "الرصيد المستحق", "المبلغ المتبقي")

        return {
            "Invoice Number":  inv_num,
            "Invoice Date":    inv_date,
            "Customer Name":   customer_name,
            "Customer Tax No": tax_numbers[1] if len(tax_numbers) > 1 else "",
            "Customer CR":     cr_numbers[1]  if len(cr_numbers)  > 1 else "",
            "Address":         address,
            "Balance":         balance,
            "Paid":            paid,
            "Not Paid":        not_paid,
            "Source File":     pdf_path.name,
        }

    except Exception as e:
        st.error(f"❌ Metadata error: {e}")
        return {}

# ── Table extraction ───────────────────────────────────────────
def is_numeric_row(row):
    return any(re.sub(r"[,٬٫\s]", "", str(c)).replace(".", "").isdigit() for c in row if c)

def extract_tables(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_rows = []
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    for row in table:
                        if row and is_numeric_row(row):
                            cleaned = [reshape(str(c)) if c else "" for c in row]
                            all_rows.append(cleaned)

            if not all_rows:
                return pd.DataFrame()

            headers = ["Total before tax", "Quantity", "Unit price", "Count", "Description", "SKU"]
            n = min(len(headers), len(all_rows[0]))
            return pd.DataFrame(all_rows, columns=headers[:n])

    except Exception as e:
        st.error(f"❌ Table error: {e}")
        return pd.DataFrame()

# ── Process single PDF ─────────────────────────────────────────
def process_pdf(pdf_path):
    meta  = extract_metadata(pdf_path)
    table = extract_tables(pdf_path)
    if not table.empty:
        for k, v in meta.items():
            table[k] = v
        return table
    return pd.DataFrame([meta])

# ── Streamlit UI ───────────────────────────────────────────────
st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True
)

# ── DEBUG toggle ───────────────────────────────────────────────
debug_mode = st.checkbox("🔍 Show raw extracted text (for debugging)", value=False)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf_paths = []

        for uf in uploaded_files:
            fp = tmp / uf.name
            fp.write_bytes(uf.read())
            if uf.name.endswith(".zip"):
                with zipfile.ZipFile(fp) as z:
                    z.extractall(tmp)
                pdf_paths.extend(tmp.glob("*.pdf"))
            else:
                pdf_paths.append(fp)

        all_data = []
        for path in pdf_paths:
            st.write(f"📄 Processing: **{path.name}**")

            # ── Show raw text if debug on ──────────────────────
            if debug_mode:
                raw = get_raw_text(path)
                st.subheader("📋 Raw extracted text:")
                st.text_area("Raw text", raw, height=300)

            df = process_pdf(path)
            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            # Clean numerics
            for col in ["Total before tax", "Balance", "Paid", "Not Paid"]:
                if col in final_df.columns:
                    final_df[col] = (
                        final_df[col].astype(str)
                        .apply(clean_number)
                        .replace("", None)
                        .astype(float)
                    )

            if "Total before tax" in final_df.columns:
                final_df["VAT 15%"]         = (final_df["Total before tax"] * 0.15).round(2)
                final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15%"]).round(2)

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            cols = [
                "Invoice Number", "Invoice Date", "Customer Name",
                "Customer Tax No", "Customer CR", "Address",
                "Balance", "Paid", "Not Paid",
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Count", "Description", "SKU",
                "Source File",
            ]
            final_df = final_df.reindex(columns=cols)

            st.success("✅ Done!")
            st.dataframe(final_df)

            out = BytesIO()
            final_df.to_excel(out, index=False, engine="openpyxl")
            out.seek(0)
            st.download_button(
                "📥 Download Excel", out, "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No data extracted.")






# # streamlit_app.py

# import streamlit as st
# import os
# import fitz  # PyMuPDF
# import pdfplumber
# import pandas as pd
# import re
# import tempfile
# import zipfile
# from pathlib import Path
# from io import BytesIO
# import arabic_reshaper
# from bidi.algorithm import get_display

# # =========================
# # Arabic Helpers
# # =========================

# def reshape_arabic_text(text):
#     try:
#         reshaped = arabic_reshaper.reshape(text)
#         bidi_text = get_display(reshaped)
#         return bidi_text
#     except:
#         return text

# # =========================
# # Metadata Extraction (PyMuPDF)
# # =========================

# def extract_metadata(pdf_path):
#     try:
#         with fitz.open(pdf_path) as doc:
#             full_text = "\n".join([page.get_text() for page in doc])

#         def find_field(text, keyword):
#             pattern = rf"{keyword}[:\s]*([^\n]*)"
#             match = re.search(pattern, text)
#             return match.group(1).strip() if match else ""

#         address_part1 = find_field(full_text, "رقم السجل")
#         address_part2 = find_field(full_text, "العنوان")

#         # === Clean customer_name ===
#         raw_customer = find_field(full_text, "فاتورة ضريبية")
#         raw_customer = re.sub(r"اسم العميل.*", "", raw_customer).strip()
#         raw_customer = re.sub(r":.*", "", raw_customer).strip()

#         # === Clean address ===
#         full_address = f"{address_part1} {address_part2}".strip()

#         metadata = {
#             "Invoice Number": find_field(full_text, "رقم الفاتورة"),
#             "Invoice Date": find_field(full_text, "تاريخ الفاتورة"),
#             "Customer Name": raw_customer,
#             "Address": full_address,
#             "Paid": find_field(full_text, "مدفوع"),
#             "Balance": find_field(full_text, "اإلجمالي"),
#             "Source File": pdf_path.name,
#             "Not Paid": find_field(full_text, "الرصيد المستحق")
#         }

#         return metadata

#     except Exception as e:
#         st.error(f"❌ Error extracting metadata from {pdf_path.name}: {e}")
#         return {}

# # =========================
# # Table Extraction (pdfplumber)
# # =========================

# def is_data_row(row):
#     return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

# def fix_shifted_rows(row):
#     if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
#         row[3] = row[4]
#         row[4] = row[5]
#         row[5] = row[6]
#         row = row[:6]
#     return row

# def extract_tables(pdf_path):
#     try:
#         with pdfplumber.open(pdf_path) as pdf:
#             all_data = []
#             for page in pdf.pages:
#                 tables = page.extract_tables()
#                 for table in tables:
#                     if table:
#                         df = pd.DataFrame(table)
#                         df = df.dropna(how="all").reset_index(drop=True)
#                         if df.empty:
#                             continue

#                         merged_rows = []
#                         temp_row = []

#                         for _, row in df.iterrows():
#                             row_values = row.fillna("").astype(str).tolist()
#                             row_values = [reshape_arabic_text(cell) for cell in row_values]
#                             row_values = fix_shifted_rows(row_values)

#                             if is_data_row(row_values):
#                                 if temp_row:
#                                     combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
#                                     merged_rows.append(combined)
#                                     temp_row = []
#                                 else:
#                                     merged_rows.append(row_values)
#                             else:
#                                 temp_row = row_values

#                         if merged_rows:
#                             num_cols = len(merged_rows[0])
#                             headers = ["Total before tax", "الكمية", "Unit price", "Quantity", "Description", "SKU", "إضافي"]
#                             df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
#                             all_data.append(df_cleaned)
#             return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

#     except Exception as e:
#         st.error(f"❌ Error extracting table from {pdf_path.name}: {e}")
#         return pd.DataFrame()

# # =========================
# # Main Process Function
# # =========================

# def process_pdf(pdf_path):
#     metadata = extract_metadata(pdf_path)
#     table_data = extract_tables(pdf_path)

#     if not table_data.empty:
#         for key, value in metadata.items():
#             table_data[key] = value
#         return table_data
#     else:
#         return pd.DataFrame([metadata])

# # =========================
# # Streamlit App UI
# # =========================

# st.set_page_config(page_title="Merged Arabic Invoice Extractor", layout="wide")
# st.title("📄 Invoice Extractor Pdf to Excel")

# uploaded_files = st.file_uploader("Upload PDF files", type=["pdf", "zip"], accept_multiple_files=True)

# if uploaded_files:
#     with tempfile.TemporaryDirectory() as temp_dir:
#         temp_dir = Path(temp_dir)
#         pdf_paths = []

#         for uploaded_file in uploaded_files:
#             file_path = temp_dir / uploaded_file.name
#             with open(file_path, "wb") as f:
#                 f.write(uploaded_file.read())

#             if uploaded_file.name.endswith(".zip"):
#                 with zipfile.ZipFile(file_path, 'r') as zip_ref:
#                     zip_ref.extractall(temp_dir)
#                 for pdf in temp_dir.glob("*.pdf"):
#                     pdf_paths.append(pdf)
#             else:
#                 pdf_paths.append(file_path)

#         all_data = []
#         for pdf_path in pdf_paths:
#             st.write(f"📄 Processing: {pdf_path.name}")
#             df = process_pdf(pdf_path)
#             if not df.empty:
#                 all_data.append(df)

#         if all_data:
#             final_df = pd.concat(all_data, ignore_index=True)

#             # ======== Cleaning Steps ========
#             if "Total before tax" in final_df.columns:
#                 final_df["Total before tax"] = (
#                     final_df["Total before tax"].astype(str)
#                     .str.replace(r"[^\d.,]", "", regex=True)
#                     .str.replace(",", "", regex=False)
#                     .replace("", None)
#                     .astype(float)
#                 )
#                 final_df["VAT 15%"] = (final_df["Total before tax"] * 0.15).round(2)
#                 final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15%"]).round(2)

#             for col in ["Paid", "Balance","Not Paid"]:
#                 if col in final_df.columns:
#                     final_df[col] = (
#                         final_df[col].astype(str)
#                         .str.replace(r"[^\d.,]", "", regex=True)
#                         .str.replace(",", "", regex=False)
#                         .replace("", None)
#                         .astype(float)
#                     )

#             # ======== Fix Invoice Date to MM/DD/YYYY ========
#             if "Invoice Date" in final_df.columns:
#                 final_df["Invoice Date"] = pd.to_datetime(
#                     final_df["Invoice Date"],
#                     errors="coerce",
#                     dayfirst=True
#                 ).dt.strftime("%m/%d/%Y")

#             # ======== Keep only required columns in order ========
#             required_columns = [
#                 "Invoice Number", "Invoice Date", "Customer Name", "Balance","Paid", "Address", 
#                 "Total before tax", "VAT 15%", "Total after tax",
#                 "Unit price", "Quantity", "Description", "SKU",
#                 "Source File"
#             ]

#             final_df = final_df.reindex(columns=required_columns)

#             st.success("✅ Extraction & cleaning complete!")
#             st.dataframe(final_df)

#             output = BytesIO()
#             final_df.to_excel(output, index=False, engine="openpyxl")
#             output.seek(0)

#             st.download_button(
#                 label="📥 Download Excel",
#                 data=output,
#                 file_name="Merged_Invoice_Data.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         else:
#             st.warning("⚠️ No data extracted from the uploaded files.")
