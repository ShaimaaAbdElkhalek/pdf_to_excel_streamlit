import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO

# =========================
# Main Extraction Process
# =========================

def process_pdf(pdf_path):
    # Default Dictionary to hold all metadata
    metadata = {
        "Invoice Number": None,
        "Invoice Date": None,
        "Customer Name": None,
        "Address": None,
        "Paid": None,
        "Balance": None, 
        "Total before tax": None,
        "VAT 15%": None,
        "Total after tax": None,
        "Source File": Path(pdf_path).name
    }
    
    table_data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                # layout=True forces the text to maintain its visual horizontal spacing 
                text = page.extract_text(layout=True)
                if not text:
                    text = page.extract_text() # Fallback
                if text:
                    full_text += text + "\n"

            lines = full_text.split('\n')

            for line in lines:
                # Condense multiple spaces to single spaces for easier reading
                line_str = " ".join(line.split())
                if not line_str.strip():
                    continue

                # ---------------------------------
                # 1. METADATA EXTRACTION
                # ---------------------------------
                if "رقم" in line_str or "فاتورة" in line_str or "02445" in line_str:
                    num_match = re.search(r'\b\d{5}\b', line_str)
                    if num_match and not metadata["Invoice Number"]:
                        metadata["Invoice Number"] = num_match.group(0)

                date_match = re.search(r'\b\d{2}/\d{2}/\d{4}\b', line_str)
                if date_match and not metadata["Invoice Date"]:
                    metadata["Invoice Date"] = date_match.group(0)

                if "اسم العميل" in line_str or "لألأة" in line_str:
                    clean_name = re.sub(r'(اسم العميل|:|فاتورة|ضريبية|رقم|\b\d{5}\b)', '', line_str).strip()
                    if clean_name:
                        metadata["Customer Name"] = clean_name

                if "العنوان" in line_str or "محاسن" in line_str:
                    addr = re.sub(r'(العنوان|:)', '', line_str).strip()
                    if addr and len(addr) > 5:
                        metadata["Address"] = addr

                # ---------------------------------
                # 2. TOTALS EXTRACTION
                # ---------------------------------
                if "مدفوع" in line_str:
                    nums = re.findall(r'-?\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    if nums: metadata["Paid"] = nums[-1]

                if "الرصيد المستحق" in line_str or "الرصيد" in line_str:
                    nums = re.findall(r'-?\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    if nums: metadata["Balance"] = nums[-1]

                if "القيمة المضافة" in line_str:
                    nums = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    if nums: metadata["VAT 15%"] = nums[-1]

                if "الإجمالي" in line_str:
                    nums = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    if nums: metadata["Total after tax"] = nums[-1]

                # Fallback Total before tax from the "المجموع" line before VAT
                if "المجموع" in line_str and "القيمة" not in line_str:
                    nums = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    if nums and len(nums) == 1: 
                        metadata["Total before tax"] = nums[-1]

                # ---------------------------------
                # 3. LINE ITEMS (PRODUCTS) EXTRACTION
                # ---------------------------------
                # Detect lines with English characters (e.g., VEAL LEG SAHIBA)
                eng_desc_match = re.search(r'[A-Za-z\s]{4,}', line_str)
                if eng_desc_match:
                    desc = eng_desc_match.group(0).strip()
                    
                    # Find all numbers in the row
                    nums = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', line_str)
                    
                    if len(nums) >= 2:
                        prices = [n for n in nums if '.' in n]
                        unit_price = prices[0] if prices else "0"
                        
                        # Find Qty (Usually an integer without commas or dots)
                        qtys = [n for n in nums if '.' not in n and ',' not in n and int(n) <= 1000]
                        qty = qtys[0] if qtys else "0"
                        
                        # Extract Arabic SKU by removing known values (English, Prices, Quantities, and specific words like 'كرتونة')
                        sku_parts = []
                        for word in line_str.split():
                            if not re.search(r'[A-Za-z]', word) and word not in ["كرتونة", unit_price, qty]:
                                if not re.fullmatch(r'\d{1,3}(?:,\d{3})*(?:\.\d+)?', word):
                                    sku_parts.append(word)
                        
                        sku = " ".join(sku_parts)
                        
                        table_data.append({
                            "Unit price": unit_price,
                            "Quantity": qty,
                            "Description": desc,
                            "SKU": sku if len(sku) > 2 else "None"
                        })

    except Exception as e:
        st.error(f"❌ Error extracting data from {Path(pdf_path).name}: {e}")

    # Combine tables + metadata
    if not table_data:
        table_data = [{"Unit price": None, "Quantity": None, "Description": None, "SKU": None}]
        
    df = pd.DataFrame(table_data)
    for key, value in metadata.items():
        df[key] = value
        
    return df

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

            if uploaded_file.name.endswith(".zip"):
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

            # ======== Cleaning & Math Formatting ========
            
            # Clean numeric columns (removes commas, keeps negative minus signs and dots)
            numeric_cols = ["Total before tax", "VAT 15%", "Total after tax", "Paid", "Balance", "Unit price", "Quantity"]
            for col in numeric_cols:
                if col in final_df.columns:
                    final_df[col] = (
                        final_df[col].astype(str)
                        .str.replace(r'[^\d.-]', '', regex=True) # strip letters/spaces
                        .replace('', '0')                        # replace blanks
                        .replace('None', '0')
                        .astype(float)                           # convert to number
                    )

            # Fix Invoice Date Format (To standard MM/DD/YYYY)
            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"],
                    errors="coerce",
                    dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            # ======== Output Columns Order ========
            required_columns = [
                "Invoice Number", "Invoice Date", "Customer Name", "Balance", "Paid", "Address", 
                "Total before tax", "VAT 15%", "Total after tax",
                "Unit price", "Quantity", "Description", "SKU",
                "Source File"
            ]
            
            # Ensure all required columns exist so the app doesn't crash
            for col in required_columns:
                if col not in final_df.columns:
                    final_df[col] = None

            final_df = final_df.reindex(columns=required_columns)

            st.success("✅ Extraction & cleaning complete!")
            st.dataframe(final_df)

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
