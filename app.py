import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
from PIL import Image
import pytesseract
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

# ── OCR: convert PDF pages to images then run Tesseract ────────
def ocr_pdf(pdf_path):
    """Extract text from scanned PDF using OCR (Arabic + English)."""
    full_text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            # Render page at high DPI for better OCR accuracy
            mat = fitz.Matrix(3, 3)  # 3x zoom = ~216 DPI
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # Run OCR with Arabic + English
            text = pytesseract.image_to_string(img, lang="ara+eng", config="--psm 6")
            full_text += text + "\n"
    return full_text

# ── Field finders ──────────────────────────────────────────────
def find_after(text, *keywords):
    for kw in keywords:
        pattern = rf"{re.escape(kw)}\s*[:\-]?\s*([^\n]+)"
        m = re.search(pattern, text)
        if m:
            val = m.group(1).strip()
            if val:
                return val
    return ""

def find_first_number(text, *keywords):
    for kw in keywords:
        pattern = rf"{re.escape(kw)}[^\d\n]*([\d,٬٫\.]+)"
        m = re.search(pattern, text)
        if m:
            return clean_number(m.group(1))
    return ""

# ── Metadata extraction ────────────────────────────────────────
def extract_metadata(pdf_path, raw_text):
    try:
        raw = raw_text

        # Debug: show what we got
        tax_numbers = re.findall(r"\b\d{15}\b", raw)
        cr_numbers  = re.findall(r"\b\d{10}\b", raw)

        # Date: DD/MM/YYYY
        date_match = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", raw)
        inv_date = date_match.group(1) if date_match else ""

        # Invoice number: 4-6 digit number
        inv_match = re.search(r"\b(\d{4,6})\b", raw)
        inv_num = inv_match.group(1) if inv_match else ""

        # Customer name after اسم العميل
        cname = find_after(raw, "اسم العميل")

        # Address
        addr_match = re.search(r"العنوان\s*[:\s]+(.+?)(?:\n\d|\Z)", raw, re.DOTALL)
        address = addr_match.group(1).strip().replace("\n", "، ") if addr_match else ""

        # Totals
        balance  = find_first_number(raw, "الإجمالي", "اإلجمالي", "الاجمالي")
        paid     = find_first_number(raw, "مدفوع")
        not_paid = find_first_number(raw, "الرصيد المستحق", "المبلغ المتبقي")

        return {
            "Invoice Number":  inv_num,
            "Invoice Date":    inv_date,
            "Customer Name":   cname,
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

# ── Table extraction from OCR text ────────────────────────────
def extract_table_from_ocr(raw_text):
    """Parse lines that look like invoice item rows."""
    rows = []
    # Look for lines with numbers that resemble item rows
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        # Line must have at least 2 numbers (qty, price, total)
        nums = re.findall(r"[\d,٬٫]+\.?\d*", line)
        if len(nums) >= 2:
            rows.append({"Raw Line": line, "Numbers": " | ".join(nums)})

    return pd.DataFrame(rows) if rows else pd.DataFrame()

# ── Process single PDF ─────────────────────────────────────────
def process_pdf(pdf_path):
    with st.spinner(f"🔍 Running OCR on {pdf_path.name}..."):
        raw_text = ocr_pdf(pdf_path)

    meta  = extract_metadata(pdf_path, raw_text)
    table = extract_table_from_ocr(raw_text)

    if not table.empty:
        for k, v in meta.items():
            table[k] = v
        return table, raw_text
    return pd.DataFrame([meta]), raw_text

# ── Streamlit UI ───────────────────────────────────────────────
st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel (OCR Mode)")

st.info("🔍 This PDF is **image-based** — using OCR (Tesseract) to extract text.")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True
)

debug_mode = st.checkbox("🔍 Show raw OCR text (for debugging)", value=True)

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
            df, raw_text = process_pdf(path)

            if debug_mode:
                st.subheader("📋 Raw OCR text:")
                st.text_area("OCR output", raw_text, height=300, key=str(path))

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            for col in ["Balance", "Paid", "Not Paid"]:
                if col in final_df.columns:
                    final_df[col] = (
                        final_df[col].astype(str)
                        .apply(clean_number)
                        .replace("", None)
                        .astype(float)
                    )

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

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
