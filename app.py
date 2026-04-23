# streamlit_app.py

import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

# OCR
import pytesseract
from pdf2image import convert_from_path

# =========================
# OCR
# =========================

def extract_text_ocr(pdf_path):
    images = convert_from_path(pdf_path, dpi=300)
    text = ""
    for img in images:
        text += pytesseract.image_to_string(img, lang='ara+eng') + "\n"
    return text

# =========================
# TEXT NORMALIZATION
# =========================

def normalize_text(text):
    fixes = {
        "الغاتورة": "الفاتورة",
        "اإلجمالي": "الإجمالي",
        "الاجمالي": "الإجمالي",
        "تارخ": "تاريخ",
    }
    for k, v in fixes.items():
        text = text.replace(k, v)
    return text

# =========================
# FIND FIELD (FUZZY)
# =========================

def find_field(text, keywords):
    if isinstance(keywords, str):
        keywords = [keywords]

    for key in keywords:
        pattern = rf"{key}\s*[:\-]?\s*(.+)"
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return ""

# =========================
# CLEAN NUMBERS
# =========================

def clean_number(x):
    if pd.isna(x):
        return None

    x = str(x)
    x = re.sub(r"[^\d.,]", "", x)
    x = x.replace("٫", ".").replace("٬", "")

    if len(x) > 15:
        return None

    if x.count(".") > 1:
        parts = x.split(".")
        x = parts[0] + "." + "".join(parts[1:])

    try:
        return float(x)
    except:
        return None

# =========================
# METADATA (FULL)
# =========================

def extract_metadata(text, filename):

    return {
        "Invoice Number": find_field(text, ["رقم الفاتورة","فاتورة رقم"]),
        "Invoice Date": find_field(text, ["تاريخ الفاتورة","تاريخ"]),
        "Customer Name": find_field(text, ["اسم العميل","العميل"]),
        "VAT Number": find_field(text, ["الرقم الضريبي"]),
        "CR Number": find_field(text, ["السجل التجاري"]),
        "Phone": find_field(text, ["رقم الجوال"]),
        "IBAN": find_field(text, ["الايبان"]),
        "Address": find_field(text, ["العنوان"]),
        "Paid": find_field(text, ["مدفوع"]),
        "Balance": find_field(text, ["الإجمالي"]),
        "Not Paid": find_field(text, ["الرصيد المستحق"]),
        "Source File": filename
    }

# =========================
# SMART OCR TABLE PARSER 🔥
# =========================

def extract_table_from_ocr(text):

    lines = [l.strip() for l in text.split("\n") if l.strip()]
    rows = []

    for line in lines:

        # detect product line
        if any(c.isdigit() for c in line) and len(line) > 15:

            nums = re.findall(r"\d+\.?\d*", line)

            if len(nums) >= 2:
                rows.append({
                    "Description": line,
                    "Quantity": nums[-2],
                    "Unit price": nums[-1]
                })

    return pd.DataFrame(rows)

# =========================
# MAIN PROCESS
# =========================

def process_pdf(pdf_path):

    # try normal extraction
    with fitz.open(pdf_path) as doc:
        text = "\n".join([p.get_text() for p in doc])

    # fallback OCR
    if len(text.strip()) < 50:
        text = extract_text_ocr(pdf_path)

    text = normalize_text(text)

    metadata = extract_metadata(text, pdf_path.name)

    # table
    df_table = extract_table_from_ocr(text)

    if not df_table.empty:
        for k, v in metadata.items():
            df_table[k] = v
        return df_table

    return pd.DataFrame([metadata])

# =========================
# STREAMLIT UI
# =========================

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Smart Arabic Invoice Extractor")

files = st.file_uploader("Upload PDF or ZIP", type=["pdf","zip"], accept_multiple_files=True)

if files:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdfs = []

        for f in files:
            path = tmp / f.name
            with open(path, "wb") as out:
                out.write(f.read())

            if f.name.endswith(".zip"):
                with zipfile.ZipFile(path, 'r') as z:
                    z.extractall(tmp)
                pdfs += list(tmp.glob("*.pdf"))
            else:
                pdfs.append(path)

        all_data = []

        for pdf in pdfs:
            st.write(f"📄 {pdf.name}")
            df = process_pdf(pdf)
            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            # clean numbers
            for col in ["Quantity","Unit price","Paid","Balance","Not Paid"]:
                if col in final_df.columns:
                    final_df[col] = final_df[col].apply(clean_number)

            st.success("✅ Done")
            st.dataframe(final_df)

            buffer = BytesIO()
            final_df.to_excel(buffer, index=False)
            buffer.seek(0)

            st.download_button("📥 Download Excel", buffer, "invoices.xlsx")

        else:
            st.warning("⚠️ No data extracted")










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
