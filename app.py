import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path
from io import BytesIO
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import arabic_reshaper
from bidi.algorithm import get_display

# =====================================================================
# 1. SHARED CONFIGURATION & HELPERS (إعدادات ودوال مشتركة بين النظامين)
# =====================================================================

FINAL_COLS =[
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File"
]

UNIT_WORDS = {
    "كرتونة", "كرتون", "قطعة", "علبة", "كيس", "طن", "كجم", "لتر", "كغ",
    "جرام", "مل", "حبة", "رول", "باكيت", "صندوق",
}

PRODUCT_CATALOG = [
    {"keywords":["صاحبة", "SAHIBA"], "sku": "فيل ليج هندي صاحبة 18 ك (510)", "desc": "VEAL LEG SAHIBA"},
    {"keywords":["الفاروق", "ELFAROUK", "ELFAROK"], "sku": "فيل ليج هندي الفاروق 18 ك", "desc": "VEAL LEG ELFAROUK"},
    {"keywords":["فوركوارتر", "FOREQUARTER", "FQ", "AMBER"], "sku": "فوركوارتر هندي عمبر", "desc": "FQ FOREQUARTER AMBER"},
    {"keywords":["كبدة", "LAMBLIVER", "JUNNE"], "sku": "كبدة ضأن استرالي جوني جولد", "desc": "LAMBLIVER JUNNE GOLD"},
    {"keywords":["عجل مقطع", "BONEINCUT"], "sku": "عجل مقطع افيكو نيوزلاندي", "desc": "BONEINCUT WAY"},
    {"keywords":["فخده", "WHOLE LEG", "رستم", "RUSTAM"], "sku": "فخده كامله هندي رستم", "desc": "WHOLE LEG RUSTAM"},
    {"keywords":["فيليه", "TENDERLOIN"], "sku": "فيليه عجل هندي عمبر 18 ك (99)", "desc": "VEAL TENDERLOIN KG"},
    {"keywords":["صدور", "BREAST", "RUSSIA"], "sku": "صدور دجاج روسي", "desc": "CHICKEN BREAST RUSSIA"},
    {"keywords": ["امامي", "FORESHANK", "استرالي"], "sku": "امامي لامب بالك استرالي", "desc": "FORESHANK AS JLO"}
]

def reshape(text):
    try:
        return get_display(arabic_reshaper.reshape(text))
    except:
        return text

def clean_number(val):
    s = str(val).replace(",", "").replace("\u066c", "").replace("\u066b", ".")
    s = re.sub(r"[^\d.]", "", s)
    try:
        if len(s.split('.')[0]) > 10: return None
        return float(s) if s else None
    except:
        return None

def standardize_product(raw_text):
    raw_upper = raw_text.upper()
    for product in PRODUCT_CATALOG:
        for kw in product["keywords"]:
            if kw.upper() in raw_upper:
                return product["sku"], product["desc"]
    return None, None

def extract_name_from_filename(pdf_path):
    stem = Path(pdf_path).stem
    name = re.sub(r"[-*\s]*\d+[-*\s]*$", "", stem).strip()
    name = re.sub(r"^[-*\s]*\d+[-*\s]*", "", name).strip()
    if re.search(r"[\u0600-\u06FF]", name): return name
    return ""

# =====================================================================
# 2. NATIVE PDF PATHWAY (نظام النصوص الأصلية - كودك القديم)
# =====================================================================

def extract_metadata_native(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        def find_field(text, keyword):
            pattern = rf"{keyword}[:\s]*([^\n]*)"
            match = re.search(pattern, text)
            return match.group(1).strip() if match else ""

        address_part1 = find_field(full_text, "رقم السجل")
        address_part2 = find_field(full_text, "العنوان")

        raw_customer = find_field(full_text, "فاتورة ضريبية")
        raw_customer = re.sub(r"اسم العميل.*", "", raw_customer).strip()
        raw_customer = re.sub(r":.*", "", raw_customer).strip()

        full_address = f"{address_part1} {address_part2}".strip()

        return {
            "Invoice Number": find_field(full_text, "رقم الفاتورة"),
            "Invoice Date": find_field(full_text, "تاريخ الفاتورة"),
            "Customer Name": raw_customer,
            "Address": full_address,
            "Paid": find_field(full_text, "مدفوع"),
            "Balance": find_field(full_text, "اإلجمالي"),
            "Source File": pdf_path.name,
        }
    except Exception as e:
        return {}

def is_data_row_native(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

def fix_shifted_rows_native(row):
    if len(row) == 7 and row[3].strip() == "" and row[4].strip() != "":
        row[3] = row[4]
        row[4] = row[5]
        row[5] = row[6]
        row = row[:6]
    return row

def process_native_pdf(pdf_path):
    metadata = extract_metadata_native(pdf_path)
    all_data =[]
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or[]):
                    df = pd.DataFrame(table).dropna(how="all").reset_index(drop=True)
                    if df.empty: continue

                    merged_rows, temp_row = [],[]
                    for _, row in df.iterrows():
                        row_values = row.fillna("").astype(str).tolist()
                        row_values =[reshape(cell) for cell in row_values]
                        row_values = fix_shifted_rows_native(row_values)

                        if is_data_row_native(row_values):
                            if temp_row:
                                combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
                                merged_rows.append(combined)
                                temp_row =[]
                            else:
                                merged_rows.append(row_values)
                        else:
                            temp_row = row_values

                    if merged_rows:
                        num_cols = len(merged_rows[0])
                        headers =["Total before tax", "الكمية", "Unit price", "Quantity", "Description", "SKU", "إضافي"]
                        df_cleaned = pd.DataFrame(merged_rows, columns=headers[:num_cols])
                        all_data.append(df_cleaned)
    except Exception as e:
        pass

    table_data = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

    if not table_data.empty:
        for key, value in metadata.items():
            table_data[key] = value
        final_df = table_data
    else:
        final_df = pd.DataFrame([metadata])

    # Cleaning native data to match FINAL_COLS
    if "Total before tax" in final_df.columns:
        final_df["Total before tax"] = final_df["Total before tax"].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "").replace("", "0").astype(float)
        final_df["VAT 15%"] = (final_df["Total before tax"] * 0.15).round(2)
        final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15%"]).round(2)
    else:
        final_df["Total before tax"] = 0.0
        final_df["VAT 15%"] = 0.0
        final_df["Total after tax"] = 0.0

    for col in ["Paid", "Balance"]:
        if col in final_df.columns:
            final_df[col] = final_df[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "").replace("", "0").astype(float)

    # توحيد أسماء المنتجات كما في الـ OCR
    if "Description" in final_df.columns and "SKU" in final_df.columns:
        for idx, row in final_df.iterrows():
            raw_text = str(row.get("Description", "")) + " " + str(row.get("SKU", ""))
            std_sku, std_desc = standardize_product(raw_text)
            if std_sku:
                final_df.at[idx, "SKU"] = std_sku
                final_df.at[idx, "Description"] = std_desc

    # ضمان وجود كل الأعمدة
    for col in FINAL_COLS:
        if col not in final_df.columns:
            final_df[col] = None

    return final_df[FINAL_COLS], "Native (Text)", "تم استخراج النص مباشرة من الملف الأصلي بدون OCR."

# =====================================================================
# 3. OCR PDF PATHWAY (نظام الصور والاسكانر - كودنا الجديد)
# =====================================================================

HEADER_KW =["البند", "الوصف", "العدد", "سعر الوحدة", "الكمية", "الوحدة"]
STOP_KWS =["المجموع", "القيمة المضافة", "الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي", "الرصيد", "الايبان", "رقم الحساب", "الإإحمالي"]
SKIP_KWS =["العنوان", "الضريبي", "السجل", "تاريخ", "العميل", "فاكس", "هاتف", "جوال", "إلى", "رقم الفاتورة", "الغاتورة", "الفاتورة", "مدفوع", "مرتجع"]

def get_ocr_text(pdf_path):
    try:
        images = convert_from_path(str(pdf_path))
        ocr_text = ""
        for i, image in enumerate(images):
            ocr_text += f"\n--- الصفحة {i+1} ---\n"
            ocr_text += pytesseract.image_to_string(image, lang="ara+eng", config="--psm 6") + "\n"
        return ocr_text
    except Exception:
        return ""

def get_nums_with_context(segment):
    seg_clean = segment.replace(",", "")
    matches = re.finditer(r"\b\d+(?:\.\d+)?\b", seg_clean)
    res =[]
    for m in matches:
        s = m.group(0)
        v = clean_number(s)
        if v is not None and v > 0:
            res.append((s, v))
    return res

def clean_sku_ocr(raw_sku):
    cleaned = re.sub(r"\|", " ", raw_sku)
    words =[w for w in cleaned.split() if w not in UNIT_WORDS and (len(w) > 1 or w == "ك")]
    return " ".join(words).strip()

def extract_sku_from_line(line):
    ar_block = re.search(r"([\u0600-\u06FF][\u0600-\u06FF\s\d\(\)ك]*)", line)
    raw = ar_block.group(1).strip() if ar_block else ""
    if not raw:
        ar_words = re.findall(r"[\u0600-\u06FF]{2,}", line)
        raw = " ".join(w for w in ar_words if w not in UNIT_WORDS)
    for b in re.findall(r"[\(\)\[\]]\s*\d+\s*[\(\)\[\]]", line):
        b_clean = "(" + re.search(r"\d+", b).group() + ")"
        if b_clean not in raw.replace(" ", ""):
            raw = raw + " " + b_clean
    return clean_sku_ocr(raw)

def parse_item_line_ocr(line):
    line_clean = re.sub(r"[\(\)\[\]]\s*\d+(?:\.\d+)?\s*[\(\)\[\]]", " ", line)
    nums = get_nums_with_context(line_clean)
    if len(nums) < 2: return None

    # حذف المجاميع إذا دخلت في السطر
    if len(nums) >= 3 and nums[-1][1] > 100 and nums[-1][1] > nums[-2][1] * 1.5:
        nums = nums[:-1]
    if len(nums) < 1: return None

    qty, unit_price = None, None
    decimals = [t for t in nums if '.' in t[0]]
    integers = [t for t in nums if '.' not in t[0]]
    sku_nums = {18, 99, 510, 106, 6, 2, 4, 3590, 10, 9, 2026}
    integers = [t for t in integers if t[1] not in sku_nums]

    # القاعدة الذهبية
    if decimals:
        price_idx = decimals[-1][0]
        unit_price = decimals[-1][1]
        possible_qtys_before = [t for t in integers if t[0] < price_idx]
        if possible_qtys_before:
            qty = possible_qtys_before[-1][1]
        else:
            possible_qtys_after = [t for t in integers if t[0] > price_idx]
            if possible_qtys_after: qty = possible_qtys_after[0][1]
            elif integers: qty = integers[0][1]
    else:
        if len(integers) >= 2:
            qty = integers[-2][1]
            unit_price = integers[-1][1]
        elif nums:
            qty, unit_price = nums[0][1], nums[0][1]

    if qty is not None:
        try: qty = int(qty) if float(qty).is_integer() else qty
        except: pass

    std_sku, std_desc = standardize_product(line)
    if std_sku:
        sku, desc = std_sku, std_desc
    else:
        all_eng = re.findall(r"[A-Za-z]{2,}", line)
        desc_words = [w for w in all_eng if len(w) >= 3 or w.isupper()]
        desc = " ".join(dict.fromkeys(desc_words)).strip()
        sku = extract_sku_from_line(line)

    if not (sku or desc): return None
    return {"SKU": sku, "Description": desc, "Quantity": qty, "Unit price": unit_price}

def extract_metadata_ocr(pdf_path, text):
    cname = ""
    m_name = re.search(r'اسم العميل\s*:\s*(.*?)(?=رقم|التاريخ|الرقم|\n)', text)
    if m_name:
        cname = re.sub(r'الغاتورة.*|الفاتورة.*|الفغاتورة.*|إلى.*', '', m_name.group(1)).strip()

    inv_num = ""
    m_inv = re.search(r'رقم\s*(?:ال[غف]اتورة|الفغاتورة|فاتورة)\s*[:\-]?\s*(\d{4,6})', text)
    if not m_inv: m_inv = re.search(r'رقم.*?\s+(\d{4,6})\b', text)
    if m_inv: inv_num = m_inv.group(1).strip()

    inv_date = ""
    m_date = re.search(r'تاريخ.*?\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})', text)
    if m_date: inv_date = m_date.group(1).strip()

    address = ""
    m_add = re.search(r'العنوان\s*:\s*(.+?)(?=\n\s*05|\n\s*\d{10}|\n\s*البند|\n\s*المجموع|05\d{8}|فيل|كبدة|عجل|فخده|فوركوارتر|فيليه|صدور|امامي)', text, re.DOTALL)
    if m_add:
        address = re.sub(r'\s*\d{10}\s*$', '', m_add.group(1).replace('\n', ' ')).strip()

    safe_text = re.sub(r'SA\d{22}|\b\d{10,}\b|\d{1,2}[/\-]\d{1,2}[/\-]\d{4}', '', text) 

    ta, tb, vat = 0.0, 0.0, 0.0
    m_tot = re.search(r'(?:الإ[جح]مالي|الإإ[جح]مالي|اإلجمالي|الاجمالي|الإجمالي)\s*[:\-]?\s*([\d.,]+)', safe_text)
    if m_tot: ta = clean_number(m_tot.group(1))

    m_sub = re.search(r'المجموع\s*[:\-]?\s*([\d.,]+)', safe_text)
    if m_sub: tb = clean_number(m_sub.group(1))

    m_vat = re.search(r'(?:القيمة المضافة|المضافة|15%)\s*[:\-]?\s*([\d.,]+)', safe_text)
    if m_vat: vat = clean_number(m_vat.group(1))

    if not ta or not tb:
        nums_raw = [clean_number(n) for n in re.findall(r"\b\d+(?:[.,]\d+)*\b", safe_text)]
        unique = sorted(set([n for n in nums_raw if n and n > 100]))
        best_diff = float("inf")
        for i, small in enumerate(unique):
            for big in unique[i + 1:]:
                if 1.10 <= big / small <= 1.20:
                    diff = abs((big / small) - 1.15)
                    if diff < best_diff:
                        best_diff, tb, ta = diff, small, big

    if ta:
        expected_tb = round(ta / 1.15, 2)
        if not tb or abs(tb - expected_tb) > 2: tb = expected_tb
        vat = round(ta - tb, 2)
    elif tb:
        ta = round(tb * 1.15, 2)
        vat = round(ta - tb, 2)

    return {
        "Invoice Number": inv_num, "Invoice Date": inv_date, "Customer Name": cname,
        "Address": address, "Balance": ta if ta else 0.0, "Paid": 0.0,
        "Total before tax": tb, "VAT 15%": vat, "Total after tax": ta,
        "Source File": pdf_path.name,
    }

def process_ocr_pdf(pdf_path):
    text = get_ocr_text(pdf_path)
    meta = extract_metadata_ocr(pdf_path, text)
    
    items =[]
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        if any(kw in line for kw in STOP_KWS) and not bool(re.search(r'[A-Za-z]{3,}', line)) and not any(h in line for h in HEADER_KW): break 
        if any(kw in line for kw in SKIP_KWS) or any(kw in line for kw in HEADER_KW): continue
            
        parsed = parse_item_line_ocr(line)
        if parsed: items.append(parsed)

    items =[i for i in items if len(i.get("Description", "")) > 2 or len(i.get("SKU", "")) > 2]
    
    if not meta["Customer Name"]:
        file_cname = extract_name_from_filename(pdf_path)
        if file_cname and len(file_cname) > 3: meta["Customer Name"] = file_cname
    if not meta["Invoice Number"]:
        m_fname_inv = re.search(r'(\d{4,6})', pdf_path.stem)
        if m_fname_inv: meta["Invoice Number"] = m_fname_inv.group(1)
    if not items:
        items =[{"Unit price": None, "Quantity": None, "Description": "", "SKU": ""}]

    rows =[{**meta, **item} for item in items]
    return pd.DataFrame(rows).reindex(columns=FINAL_COLS), "Scanned (OCR)", text

# =====================================================================
# 4. MAIN ROUTER (الموجه الذكي الذي يقرر نوع الملف)
# =====================================================================

def process_pdf(pdf_path):
    # الفحص: هل الـ PDF يحتوي على نصوص أصلية؟
    try:
        with fitz.open(pdf_path) as doc:
            text = "\n".join([page.get_text() for page in doc]).strip()
            
        # إذا كان النص المستخرج طويلاً (يحتوي على بيانات حقيقية)، نستخدم كودك القديم (Native)
        if len(text) > 100:
            return process_native_pdf(pdf_path)
        else:
            # إذا كان النص قصيراً جداً (صورة مسحوبة بسكانر)، نستخدم الـ OCR (القاعدة الذهبية)
            return process_ocr_pdf(pdf_path)
    except Exception as e:
        # في حال حدوث أي خطأ، الملجأ الأخير هو الـ OCR
        return process_ocr_pdf(pdf_path)

# =====================================================================
# 5. STREAMLIT APP UI
# =====================================================================

st.set_page_config(page_title="Smart Invoice Extractor", layout="wide")
st.title("📄 Smart Invoice Extractor (Native + OCR)")
st.markdown("يستخدم هذا النظام **التوجيه الذكي**: الفواتير الإلكترونية تُعالج بالنص الأصلي، وفواتير الإسكانر تُعالج عبر الذكاء الاصطناعي (OCR).")

uploaded_files = st.file_uploader("Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True)
debug_mode = st.checkbox("🔍 Show full raw extracted text", value=False)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf_paths =[]

        for uf in uploaded_files:
            fp = tmp / uf.name
            fp.write_bytes(uf.read())
            if uf.name.endswith(".zip"):
                with zipfile.ZipFile(fp) as z:
                    z.extractall(tmp)
                pdf_paths.extend(tmp.glob("*.pdf"))
            else:
                pdf_paths.append(fp)

        all_data =[]
        for i, path in enumerate(pdf_paths):
            st.write(f"📄 **{path.name}**")
            with st.spinner("Extracting..."):
                df, mode, raw_text = process_pdf(path)
            
            # عرض نوع الاستخراج (Native أو OCR)
            st.caption(f"Extraction Mode: `{mode}` — {len(df)} row(s)")

            if debug_mode:
                with st.expander(f"📋 Full raw text — {path.name}", expanded=False):
                    st.text(raw_text)

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            st.success(f"✅ Done! {len(final_df)} total row(s)")
            
            money_cols =["Balance", "Paid", "Total before tax", "VAT 15%", "Total after tax", "Unit price"]
            format_dict = {c: "{:.2f}" for c in money_cols if c in final_df.columns}
            st.dataframe(final_df.style.format(format_dict, na_rep=""))

            out = BytesIO()
            writer = pd.ExcelWriter(out, engine='openpyxl')
            final_df.to_excel(writer, index=False, sheet_name='Invoices')
            
            workbook = writer.book
            worksheet = writer.sheets['Invoices']
            col_indices =[final_df.columns.get_loc(c) + 1 for c in money_cols if c in final_df.columns]
            
            for row in range(2, len(final_df) + 2):
                for col_idx in col_indices:
                    cell = worksheet.cell(row=row, column=col_idx)
                    try:
                        if cell.value is not None:
                            cell.value = float(cell.value)
                            cell.number_format = '#,##0.00'
                    except: pass
                    
            writer.close()
            out.seek(0)
            
            st.download_button(
                "📥 Download Excel",
                out,
                "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No data extracted.")
