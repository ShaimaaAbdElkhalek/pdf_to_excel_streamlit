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
    s = re.sub(r"[^\d.]", "", str(val).replace(",","").replace("٬","").replace("٫","."))
    try:
        return float(s) if s else None
    except:
        return None

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

def get_ocr_words(pdf_path):
    with fitz.open(pdf_path) as doc:
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(3, 3))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    data = pytesseract.image_to_data(
        img, lang="ara+eng", config="--psm 6",
        output_type=pytesseract.Output.DATAFRAME
    )
    data = data[data["conf"] > 30].dropna(subset=["text"])
    data = data[data["text"].str.strip() != ""]
    return data

def reconstruct_table_rows(word_df, y_tolerance=15):
    if word_df.empty:
        return []
    word_df = word_df.copy()
    word_df["mid_y"] = word_df["top"] + word_df["height"] / 2
    rows = []
    used = set()
    for idx, word in word_df.iterrows():
        if idx in used:
            continue
        y = word["mid_y"]
        same_row = word_df[abs(word_df["mid_y"] - y) <= y_tolerance]
        used.update(same_row.index)
        same_row = same_row.sort_values("left", ascending=False)
        row_text = " ".join(same_row["text"].astype(str).tolist())
        rows.append({"y": y, "text": row_text, "words": same_row})
    rows.sort(key=lambda r: r["y"])
    return rows

STOP_WORDS = {"اسم","العميل","فاتورة","إلى","إلىة","التتجار",
              "رقم","الفاتورة","تاريخ","الغاتورة"}

def extract_customer_name_text(text):
    part1_words = []
    m1 = re.search(
        r"اسم العميل\s*[:\s]+(.+?)(?=\s*(?:قم الغاتورة|رقم الفاتورة)\s*\d)",
        text
    )
    if m1:
        part1_words = [w for w in re.findall(r"[\u0600-\u06FF]+", m1.group(1))
                       if w not in STOP_WORDS and len(w) > 1]

    part2_words = []
    m2 = re.search(
        r"اسم العميل.+?\n(.+?)(?=تاريخ|الرقم الضريبي|رقم السجل)",
        text, re.DOTALL
    )
    if m2:
        chunk = m2.group(1).split("\n")[0]
        part2_words = [w for w in re.findall(r"[\u0600-\u06FF]+", chunk)
                       if w not in STOP_WORDS and len(w) > 1]

    all_words = part1_words + part2_words
    seen = set()
    unique = []
    for w in all_words:
        if w not in seen:
            seen.add(w)
            unique.append(w)
    return " ".join(unique).strip()

def extract_customer_name_positional(word_df, fallback_text):
    if word_df.empty:
        return extract_customer_name_text(fallback_text)
    mask = word_df["text"].str.contains("العميل", na=False)
    if not mask.any():
        return extract_customer_name_text(fallback_text)
    label_row = word_df[mask].iloc[0]
    label_y   = label_row["top"] + label_row["height"] / 2
    label_x   = label_row["left"]
    same_row = word_df[
        (abs((word_df["top"] + word_df["height"]/2) - label_y) <= 20) &
        (word_df["left"] < label_x)
    ].sort_values("left", ascending=False)
    next_row = word_df[
        ((word_df["top"] + word_df["height"]/2) > label_y + 5) &
        ((word_df["top"] + word_df["height"]/2) < label_y + 60)
    ].sort_values("left", ascending=False)
    candidates = pd.concat([same_row, next_row])
    arabic_words = []
    for _, w in candidates.iterrows():
        t = str(w["text"]).strip()
        for word in re.findall(r"[\u0600-\u06FF]+", t):
            if word not in STOP_WORDS and len(word) > 1:
                arabic_words.append(word)
    result = " ".join(arabic_words).strip()
    return result if result else extract_customer_name_text(fallback_text)

SUMMARY_KW = [
    "المجموع","مدفوع","الرصيد","القيمة","القيمه",
    "الإجمالي","الإحمالي","اإلجمالي","الاجمالي",
    "رقم الحساب","الايبان","IBAN","SA08","Kingdome","المملكة",
    "رقم الفاتورة","تاريخ","اسم العميل","الرقم الضريبي",
    "رقم السجل","العنوان","الجوال","السجل التجاري",
]
HEADER_KW = ["البند","الوصف","العدد","سعر الوحدة","الكمية","الوحدة"]

def is_summary_row(vals):
    return any(kw in " ".join(vals) for kw in SUMMARY_KW)

def extract_items_positional(word_df, text):
    if word_df.empty:
        return []
    rows = reconstruct_table_rows(word_df)
    if not rows:
        return []
    header_idx = None
    for i, row in enumerate(rows):
        if any(kw in row["text"] for kw in HEADER_KW):
            header_idx = i
            break
    summary_idx = None
    start = (header_idx + 1) if header_idx is not None else 0
    for i, row in enumerate(rows[start:], start=start):
        if any(kw in row["text"] for kw in SUMMARY_KW):
            summary_idx = i
            break
    if header_idx is None or summary_idx is None:
        return []
    items = []
    for row in rows[header_idx + 1: summary_idx]:
        t = row["text"].strip()
        if not t or any(kw in t for kw in SUMMARY_KW):
            continue
        nums = [n for n in re.findall(r"[\d,]+\.?\d*", t)
                if len(re.sub(r"[,.]","",n)) <= 8]
        words_in_row = row["words"].sort_values("left", ascending=False)
        arabic_blocks, english_blocks = [], []
        for _, w in words_in_row.iterrows():
            wt = str(w["text"]).strip()
            if re.search(r"[\u0600-\u06FF]", wt):
                arabic_blocks.append(wt)
            elif re.search(r"[A-Za-z]", wt):
                english_blocks.append(wt)
        sku  = arabic_blocks[0] if arabic_blocks else ""
        desc = " ".join(english_blocks) if english_blocks else " ".join(arabic_blocks[1:])
        if len(nums) >= 2:
            items.append({
                "SKU":         reshape(sku),
                "Description": desc,
                "Quantity":    clean_number(nums[-3]) if len(nums) >= 3 else clean_number(nums[0]),
                "Unit price":  clean_number(nums[-2]) if len(nums) >= 2 else None,
            })
    return items

def extract_items_native(pdf_path):
    items = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    for row in table:
                        if not row:
                            continue
                        vals = [str(c).strip() if c else "" for c in row]
                        if is_summary_row(vals):
                            continue
                        num_cells = [v for v in vals
                                     if re.sub(r"[,.\s]","",v).isdigit()
                                     and 1 <= len(re.sub(r"[,.\s]","",v)) <= 8]
                        if len(num_cells) < 2:
                            continue
                        items.append({
                            "Unit price":  clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity":    clean_number(vals[3]) if len(vals) > 3 else None,
                            "Description": reshape(vals[4])      if len(vals) > 4 else "",
                            "SKU":         reshape(vals[5])      if len(vals) > 5 else "",
                        })
    except:
        pass
    return items

def extract_metadata(pdf_path, text):
    def get_amount(*keywords):
        for kw in keywords:
            m = re.search(rf"{re.escape(kw)}[^\d\n]*([\d,٬٫]+\.?\d*)", text)
            if m:
                return clean_number(m.group(1))
        return None

    inv_num = ""
    for p in [r"رقم الفاتورة\s*[:\s]*(\d{4,6})",
              r"(?:قم الغاتورة|الغاتورة)\s*(\d{4,6})",
              r"\b(0\d{4,5})\b"]:
        m = re.search(p, text)
        if m:
            inv_num = m.group(1).strip()
            break

    date_m   = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", text)
    inv_date = date_m.group(1) if date_m else ""

    address = ""
    m = re.search(
        r"العنوان\s*[:\s]+(.+?)(?=\n\d{7,}|\n(?:المجموع|مدفوع|رقم الحساب)|\Z)",
        text, re.DOTALL
    )
    if m:
        address = " ".join(m.group(1).split()).strip()
        address = re.sub(r"\s*\d{10}\s*$", "", address).strip()

    total_before = get_amount("المجموع")
    vat          = get_amount("القيمة المضافة","القيمه المضافة")
    total_after  = get_amount("الإجمالي","الإحمالي","اإلجمالي","الاجمالي")
    if not total_after and total_before and vat:
        total_after = round(total_before + vat, 2)
    balance = total_after
    paid    = get_amount("مدفوع")

    return {
        "Invoice Number":   inv_num,
        "Invoice Date":     inv_date,
        "Address":          address,
        "Balance":          balance,
        "Paid":             paid,
        "Total before tax": total_before,
        "VAT 15%":          vat,
        "Total after tax":  total_after,
        "Source File":      pdf_path.name,
    }

FINAL_COLS = [
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File",
]

def process_pdf(pdf_path):
    text, mode = get_text(pdf_path)
    meta = extract_metadata(pdf_path, text)

    if mode == "ocr":
        word_df = get_ocr_words(pdf_path)
        cname   = extract_customer_name_positional(word_df, text)
        items   = extract_items_positional(word_df, text)
    else:
        word_df = pd.DataFrame()
        cname   = extract_customer_name_text(text)
        items   = extract_items_native(pdf_path)

    meta["Customer Name"] = cname

    seen = set()
    unique_items = []
    for item in items:
        key = (item.get("Description",""), item.get("Unit price"), item.get("Quantity"))
        if key not in seen:
            seen.add(key)
            unique_items.append(item)

    if not unique_items:
        unique_items = [{"Unit price": None, "Quantity": None, "Description": "", "SKU": ""}]

    rows = [{**meta, **item} for item in unique_items]
    return pd.DataFrame(rows).reindex(columns=FINAL_COLS), mode, text

# ── Streamlit UI ──────────────────────────────────────────────
st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf","zip"], accept_multiple_files=True
)
debug_mode = st.checkbox("🔍 Show full raw extracted text", value=False)

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
        for i, path in enumerate(pdf_paths):
            st.write(f"📄 **{path.name}**")
            with st.spinner("Extracting..."):
                df, mode, raw_text = process_pdf(path)
            st.caption(f"Mode: `{mode}` — {len(df)} row(s)")

            if debug_mode:
                with st.expander(f"📋 Full raw text — {path.name}", expanded=True):
                    st.text(raw_text)
                    st.caption(f"Total characters: {len(raw_text)}")
                    st.download_button(
                        label="📋 Download raw text",
                        data=raw_text,
                        file_name=f"{path.stem}_raw.txt",
                        mime="text/plain",
                        key=f"raw_dl_{i}_{path.name}"
                    )

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            st.success(f"✅ Done! {len(final_df)} total row(s)")
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
