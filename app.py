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

def reshape(text):
    try:
        return get_display(arabic_reshaper.reshape(text))
    except:
        return text

def clean_number(val):
    s = re.sub(r"[^\d.]", "", str(val).replace(",", "").replace("٬", "").replace("٫", "."))
    try:
        return float(s) if s else None
    except:
        return None

# ── Text extraction: native first, OCR fallback ───────────────
def get_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc).strip()
    if len(text) > 50:
        return text, "native"
    ocr_text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_text += pytesseract.image_to_string(img, lang="ara+eng", config="--psm 6") + "\n"
    return ocr_text, "ocr"

# ── Metadata: one dict per invoice ────────────────────────────
def extract_metadata(pdf_path, text):
    def first(*patterns):
        for p in patterns:
            m = re.search(p, text)
            if m:
                return m.group(1).strip()
        return ""

    def get_amount(*keywords):
        for kw in keywords:
            m = re.search(rf"{re.escape(kw)}[^\d\n]*([\d,٬٫]+)", text)
            if m:
                return clean_number(m.group(1))
        return None

    # Invoice number
    inv_num = first(
        r"رقم الفاتورة\s*[:\s]*(\d{4,6})",
        r"(?:قم الغاتورة|الغاتورة)\s*(\d{4,6})",
        r"\b(0\d{4,5})\b"
    )

    # Date
    date_m = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", text)
    inv_date = date_m.group(1) if date_m else ""

    # Customer name
    cname = first(
        r"اسم العميل\s*[:\s]+(.+?)(?:\n|الرقم|رقم|إلى)",
        r"العميل\s*[:\s]+(.+?)(?:\n|رقم)"
    )
    cname = re.sub(r"(قم الغاتورة|إلى|رقم).*", "", cname).strip()

    # Tax & CR numbers
    tax_nums = re.findall(r"\b\d{15}\b", text)
    cust_tax = tax_nums[1] if len(tax_nums) > 1 else (tax_nums[0] if tax_nums else "")
    cust_cr  = first(r"رقم السجل\s*[:\s]*(\d+)")

    # Address
    addr_m = re.search(r"العنوان\s*[:\s]+(.+?)(?:\n\d{4,5}|\Z)", text, re.DOTALL)
    address = re.sub(r"\s+", " ", addr_m.group(1)).strip() if addr_m else ""

    return {
        "Invoice Number":  inv_num,
        "Invoice Date":    inv_date,
        "Customer Name":   cname,
        "Customer Tax No": cust_tax,
        "Customer CR":     cust_cr,
        "Address":         address,
        "Subtotal":        get_amount("المجموع"),
        "VAT 15%":         get_amount("القيمة المضافة", "القيمه المضافة"),
        "Total":           get_amount("الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي"),
        "Paid":            get_amount("مدفوع"),
        "Not Paid":        get_amount("الرصيد المستحق"),
        "Source File":     pdf_path.name,
    }

# ── Items: one row per product line ───────────────────────────
def extract_items(pdf_path, text, mode):
    """
    Returns list of dicts:
    {Description, SKU, Unit price, Quantity, Item Total}
    """
    items = []

    # ── Try pdfplumber table first (works for native PDFs) ────
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    for row in table:
                        if not row:
                            continue
                        vals = [str(c).strip() if c else "" for c in row]
                        # Skip header rows and summary rows
                        skip_keywords = ["المجموع", "مدفوع", "الرصيد", "القيمة",
                                         "الإجمالي", "الإحمالي", "البند", "الوصف",
                                         "سعر", "الكمية", "العدد"]
                        if any(kw in v for v in vals for kw in skip_keywords):
                            continue
                        # Row must have at least 2 numeric values
                        nums = [v for v in vals if re.sub(r"[,.\s]", "", v).isdigit() and len(v) > 0]
                        if len(nums) < 2:
                            continue

                        # Column order RTL: Total | Qty(unit) | UnitPrice | Count | Description | SKU
                        item = {
                            "Description": reshape(vals[4]) if len(vals) > 4 else "",
                            "SKU":         reshape(vals[5]) if len(vals) > 5 else "",
                            "Unit price":  clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity":    clean_number(vals[1]) if len(vals) > 1 else None,
                            "Item Total":  clean_number(vals[0]) if len(vals) > 0 else None,
                        }
                        # Only add if we got at least description or a total
                        if item["Description"] or item["Item Total"]:
                            items.append(item)
    except:
        pass

    # ── OCR fallback: parse lines that look like item rows ────
    if not items:
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            skip = ["المجموع", "مدفوع", "الرصيد", "القيمة", "الإجمالي",
                    "الإحمالي", "رقم الحساب", "الضريبي", "السجل"]
            if any(kw in line for kw in skip):
                continue
            nums = re.findall(r"[\d,]+\.?\d*", line)
            if len(nums) >= 2:
                items.append({
                    "Description": line,
                    "SKU":         "",
                    "Unit price":  clean_number(nums[-2]) if len(nums) >= 2 else None,
                    "Quantity":    clean_number(nums[-3]) if len(nums) >= 3 else None,
                    "Item Total":  clean_number(nums[-1]) if nums else None,
                })

    # If still no items, return one empty item row so invoice still appears
    if not items:
        items = [{"Description": "", "SKU": "", "Unit price": None,
                  "Quantity": None, "Item Total": None}]
    return items


# ── Process single PDF → one row per item ─────────────────────
def process_pdf(pdf_path):
    text, mode = get_text(pdf_path)
    meta  = extract_metadata(pdf_path, text)
    items = extract_items(pdf_path, text, mode)

    rows = []
    for item in items:
        row = {**meta, **item}   # invoice fields + item fields merged
        rows.append(row)

    return pd.DataFrame(rows), mode


# ── Streamlit UI ───────────────────────────────────────────────
st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True
)
debug_mode = st.checkbox("🔍 Show raw extracted text", value=False)

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
            st.write(f"📄 **{path.name}**")
            with st.spinner("Extracting..."):
                df, mode = process_pdf(path)
            st.caption(f"Mode: `{mode}` — {len(df)} item row(s)")

            if debug_mode:
                raw, _ = get_text(path)
                st.text_area("Raw text", raw, height=250, key=str(path))

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            # Fix date format
            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            # Final column order
            cols = [
                "Invoice Number", "Invoice Date",
                "Customer Name", "Customer Tax No", "Customer CR", "Address",
                "Description", "SKU", "Unit price", "Quantity", "Item Total",
                "Subtotal", "VAT 15%", "Total", "Paid", "Not Paid",
                "Source File",
            ]
            final_df = final_df.reindex(columns=cols)

            st.success(f"✅ Done! {len(final_df)} total rows")
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
