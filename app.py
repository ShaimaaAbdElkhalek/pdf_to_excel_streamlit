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
    s = re.sub(r"[^\d.]", "", str(val).replace(",", "").replace("\u066c", "").replace("\u066b", "."))
    try:
        if len(s.split('.')[0]) > 10:
            return None
        return float(s) if s else None
    except:
        return None

def extract_name_from_filename(pdf_path):
    stem = Path(pdf_path).stem
    name = re.sub(r"[-*\s]*\d+[-*\s]*$", "", stem).strip()
    name = re.sub(r"^[-*\s]*\d+[-*\s]*", "", name).strip()
    if re.search(r"[\u0600-\u06FF]", name):
        return name
    return ""

UNIT_WORDS = {
    "كرتونة", "كرتون", "قطعة", "علبة", "كيس", "طن", "كجم", "لتر", "كغ",
    "جرام", "مل", "حبة", "رول", "باكيت", "صندوق",
}

HEADER_KW =["البند", "الوصف", "العدد", "سعر الوحدة", "الكمية", "الوحدة"]
STOP_KWS =["المجموع", "القيمة المضافة", "الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي", "الرصيد", "الايبان", "رقم الحساب", "الإإحمالي"]
SKIP_KWS =["العنوان", "الضريبي", "السجل", "تاريخ", "العميل", "فاكس", "هاتف", "جوال", "إلى", "رقم الفاتورة", "رقم الغاتورة", "الفاتورة", "الغاتورة", "مدفوع", "مرتجع"]

FINAL_COLS =[
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File",
]

# 👑 الكتالوج الصارم لتوحيد أسماء المنتجات تماماً
PRODUCT_CATALOG =[
    {
        "keywords":["صاحبة", "SAHIBA"],
        "sku": "فيل ليج هندي صاحبة 18 ك (510)",
        "desc": "VEAL LEG SAHIBA"
    },
    {
        "keywords":["الفاروق", "ELFAROUK", "ELFAROK"],
        "sku": "فيل ليج هندي الفاروق 18 ك",
        "desc": "VEAL LEG ELFAROUK"
    },
    {
        "keywords":["فوركوارتر", "FOREQUARTER", "FQ", "AMBER"],
        "sku": "فوركوارتر هندي عمبر",
        "desc": "FQ FOREQUARTER AMBER"
    },
    {
        "keywords":["كبدة", "LAMBLIVER", "JUNNE"],
        "sku": "كبدة ضأن استرالي جوني جولد",
        "desc": "LAMBLIVER JUNNE GOLD"
    },
    {
        "keywords":["عجل مقطع", "BONEINCUT"],
        "sku": "عجل مقطع افيكو نيوزلاندي",
        "desc": "BONEINCUT WAY"
    },
    {
        "keywords":["فخده", "WHOLE LEG", "رستم", "RUSTAM"],
        "sku": "فخده كامله هندي رستم",
        "desc": "WHOLE LEG RUSTAM"
    },
    {
        "keywords":["فيليه", "TENDERLOIN"],
        "sku": "فيليه عجل هندي عمبر 18 ك (99)",
        "desc": "VEAL TENDERLOIN KG"
    },
    {
        "keywords":["صدور", "BREAST", "RUSSIA"],
        "sku": "صدور دجاج روسي",
        "desc": "CHICKEN BREAST RUSSIA"
    },
    {
        "keywords":["امامي", "FORESHANK", "استرالي"],
        "sku": "امامي لامب بالك استرالي",
        "desc": "FORESHANK AS JLO"
    }
]

def standardize_product(raw_text):
    raw_upper = raw_text.upper()
    for product in PRODUCT_CATALOG:
        for kw in product["keywords"]:
            if kw.upper() in raw_upper:
                return product["sku"], product["desc"]
    return None, None

def clean_sku(raw_sku):
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
    return clean_sku(raw)

def get_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc).strip()
    if len(text) > 50:
        return text, "native"
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(str(pdf_path))
        ocr_text = ""
        for i, image in enumerate(images):
            page_text = pytesseract.image_to_string(image, lang="ara+eng")
            ocr_text += f"\n--- الصفحة {i+1} ---\n{page_text}\n"
        return ocr_text, "ocr"
    except Exception:
        pass
    return "", "ocr"

def get_ocr_words(pdf_path):
    try:
        from pdf2image import convert_from_path
        images = convert_from_path(str(pdf_path))
        img = images[0]
    except Exception:
        with fitz.open(pdf_path) as doc:
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(3, 3))
        img = Image.frombytes("RGB",[pix.width, pix.height], pix.samples)
    data = pytesseract.image_to_data(
        img, lang="ara+eng", config="--psm 6",
        output_type=pytesseract.Output.DATAFRAME,
    )
    data = data[data["conf"] > 30].dropna(subset=["text"])
    data = data[data["text"].str.strip() != ""]
    return data

def reconstruct_table_rows(word_df, y_tolerance=15):
    if word_df.empty: return[]
    word_df = word_df.copy()
    word_df["mid_y"] = word_df["top"] + word_df["height"] / 2
    rows =[]
    used = set()
    for idx, word in word_df.iterrows():
        if idx in used: continue
        y = word["mid_y"]
        same_row = word_df[abs(word_df["mid_y"] - y) <= y_tolerance]
        used.update(same_row.index)
        same_row = same_row.sort_values("left", ascending=False)
        row_text = " ".join(same_row["text"].astype(str).tolist())
        rows.append({"y": y, "text": row_text, "words": same_row})
    rows.sort(key=lambda r: r["y"])
    return rows

def get_nums_with_context(segment):
    # نستخرج الأرقام كنصوص للحفاظ على العلامة العشرية لتطبيق القاعدة الذهبية
    seg_clean = segment.replace(",", "")
    matches = re.finditer(r"\b\d+(?:\.\d+)?\b", seg_clean)
    res =[]
    for m in matches:
        s = m.group(0)
        v = clean_number(s)
        if v is not None and v > 0:
            res.append((s, v))
    return res

def parse_item_line(line, tb_val=0.0):
    # مسح الأقواس لتجنب تداخل الأرقام مثل (99) أو (510)
    line_clean = re.sub(r"[\(\)\[\]]\s*\d+(?:\.\d+)?\s*[\(\)\[\]]", " ", line)

    raw_nums = get_nums_with_context(line_clean)
    if len(raw_nums) < 2: return None

    # مسح المجموع من السطر إذا كان موجوداً في النهاية لكي لا نشوش على الكمية والسعر
    if raw_nums[-1][1] > 100 and len(raw_nums) >= 3 and raw_nums[-1][1] > raw_nums[-2][1] * 1.5:
        raw_nums = raw_nums[:-1]
        
    if len(raw_nums) < 1: return None

    qty = None
    unit_price = None

    # 💡 القاعدة الذهبية: الرقم ذو العلامة العشرية (السعر)، والرقم الصحيح (الكمية)
    decimal_indices =[i for i, (s, v) in enumerate(raw_nums) if '.' in s]
    sku_nums = {18, 99, 510, 106, 3590} # أكواد مستبعدة من كونها كميات
    
    if decimal_indices:
        # نأخذ آخر رقم عشري ونعتبره السعر
        price_idx = decimal_indices[-1]
        unit_price = raw_nums[price_idx][1]
        
        # 1. نبحث عن الكمية (الرقم الصحيح) قبل السعر
        for idx in range(price_idx - 1, -1, -1):
            if '.' not in raw_nums[idx][0] and raw_nums[idx][1] not in sku_nums:
                qty = raw_nums[idx][1]
                break
                
        # 2. إذا لم نجدها قبل السعر (بسبب قلب النصوص)، نبحث بعدها
        if qty is None:
            for idx in range(price_idx + 1, len(raw_nums)):
                if '.' not in raw_nums[idx][0] and raw_nums[idx][1] not in sku_nums:
                    qty = raw_nums[idx][1]
                    break
                    
        # 3. احتياطي
        if qty is None:
            qty = raw_nums[price_idx - 1][1] if price_idx > 0 else raw_nums[0][1]
            
    else:
        # في حال لم يقرأ الـ OCR النقطة العشرية بالكامل
        if len(raw_nums) >= 2:
            qty = raw_nums[-2][1]
            unit_price = raw_nums[-1][1]
        else:
            qty = raw_nums[0][1]
            unit_price = raw_nums[0][1]

    # تحويل الكمية إلى عدد صحيح (مثال: 200.0 تصبح 200)
    if qty is not None:
        try:
            qty = int(qty) if float(qty).is_integer() else qty
        except: pass

    # توحيد اسم المنتج بشكل قاطع باستخدام الكتالوج المبرمج مسبقاً
    std_sku, std_desc = standardize_product(line)
    if std_sku:
        sku, desc = std_sku, std_desc
    else:
        all_eng = re.findall(r"[A-Za-z]{2,}", line)
        desc_words =[w for w in all_eng if len(w) >= 3 or w.isupper()]
        desc = " ".join(dict.fromkeys(desc_words)).strip()
        sku = extract_sku_from_line(line)

    if not (sku or desc): return None
    return {"SKU": sku, "Description": desc, "Quantity": qty, "Unit price": unit_price}

def extract_items_text(text, tb_val):
    items =[]
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        
        is_summary = any(kw in line for kw in STOP_KWS)
        has_english = bool(re.search(r'[A-Za-z]{3,}', line))
        
        if is_summary and not has_english and not any(h in line for h in HEADER_KW):
            break 
            
        if any(kw in line for kw in SKIP_KWS) or any(kw in line for kw in HEADER_KW):
            continue
            
        parsed = parse_item_line(line, tb_val)
        if parsed: 
            items.append(parsed)
            
    return[i for i in items if len(i.get("Description", "")) > 2 or len(i.get("SKU", "")) > 2]

def extract_metadata(pdf_path, text):
    cname = ""
    m_name = re.search(r'اسم العميل\s*:\s*(.*?)(?=رقم|التاريخ|الرقم|\n)', text)
    if m_name:
        cname = m_name.group(1).strip()
        cname = re.sub(r'الغاتورة.*|الفاتورة.*|الفغاتورة.*|إلى.*', '', cname).strip()

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
        address = m_add.group(1).replace('\n', ' ').strip()
        address = re.sub(r'\s*\d{10}\s*$', '', address).strip()

    tb = ta = vat = paid = bal = 0.0

    # 💡 تنظيف النص من أرقام البنوك والتواريخ حتى لا تتدخل في المجاميع المالية
    safe_text = re.sub(r'SA\d{22}', '', text)
    safe_text = re.sub(r'\b\d{10,}\b', '', safe_text)
    safe_text = re.sub(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{4}', '', safe_text) 

    m_tot = re.search(r'(?:الإ[جح]مالي|الإإ[جح]مالي|اإلجمالي|الاجمالي|الإجمالي)\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_tot: ta = clean_number(m_tot.group(1))

    m_sub = re.search(r'المجموع\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_sub: tb = clean_number(m_sub.group(1))

    m_vat = re.search(r'(?:القيمة المضافة|المضافة|15%)\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_vat: vat = clean_number(m_vat.group(1))

    if not ta or not tb:
        nums_raw =[clean_number(n) for n in re.findall(r"[\d,]+\.?\d*", safe_text) if clean_number(n) and clean_number(n) > 100]
        unique = sorted(set(nums_raw))
        best_diff = float("inf")
        best_ta, best_tb = None, None
        
        for i, small in enumerate(unique):
            for big in unique[i + 1:]:
                if 1.10 <= big / small <= 1.20:
                    diff = abs((big / small) - 1.15)
                    if diff < best_diff:
                        best_diff = diff
                        best_tb = small
                        best_ta = big
                        
        if not tb and best_tb: tb = best_tb
        if not ta and best_ta: ta = best_ta

    if ta:
        expected_tb = round(ta / 1.15, 2)
        expected_vat = round(ta - expected_tb, 2)
        if not tb or abs(tb - expected_tb) > 2: tb = expected_tb
        if not vat or abs(vat - expected_vat) > 2: vat = expected_vat
    elif tb:
        ta = round(tb * 1.15, 2)
        vat = round(ta - tb, 2)

    return {
        "Invoice Number": inv_num,
        "Invoice Date": inv_date,
        "Customer Name": cname,
        "Address": address,
        "Balance": ta if ta else 0.0,
        "Paid": 0.0,
        "Total before tax": tb,
        "VAT 15%": vat,
        "Total after tax": ta,
        "Source File": pdf_path.name,
    }

def process_pdf(pdf_path):
    text, mode = get_text(pdf_path)
    meta = extract_metadata(pdf_path, text)
    tb_val = meta.get("Total before tax", 0.0)

    items = extract_items_text(text, tb_val)

    if not items and mode == "ocr":
        word_df = get_ocr_words(pdf_path)
        if not word_df.empty:
            rows = reconstruct_table_rows(word_df)
            reconstructed_text = "\n".join([r["text"] for r in rows])
            items = extract_items_text(reconstructed_text, tb_val)

    file_cname = extract_name_from_filename(pdf_path)
    if file_cname and len(file_cname) > 3:
        meta["Customer Name"] = file_cname

    if not meta["Invoice Number"]:
        m_fname_inv = re.search(r'(\d{4,6})', pdf_path.stem)
        if m_fname_inv:
            meta["Invoice Number"] = m_fname_inv.group(1)

    if not items:
        items =[{"Unit price": None, "Quantity": None, "Description": "", "SKU": ""}]

    rows =[{**meta, **item} for item in items]
    return pd.DataFrame(rows).reindex(columns=FINAL_COLS), mode, text

# =====================
# Streamlit App UI
# =====================
st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel")

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
            st.caption(f"Mode: `{mode}` — {len(df)} row(s)")

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
            
            # 💡 عرض البيانات في الموقع مع تثبيت العلامة العشرية (18.00)
            money_cols =["Balance", "Paid", "Total before tax", "VAT 15%", "Total after tax", "Unit price"]
            format_dict = {c: "{:.2f}" for c in money_cols if c in final_df.columns}
            st.dataframe(final_df.style.format(format_dict, na_rep=""))

            # 💡 تحميل ملف الإكسيل مع إجبار الإكسيل على كتابة (.00)
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
