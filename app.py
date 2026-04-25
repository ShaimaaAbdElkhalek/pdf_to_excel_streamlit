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
STOP_KWS =["المجموع", "القيمة المضافة", "الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي", "الرصيد", "الايبان", "رقم الحساب"]
SKIP_KWS =["العنوان", "الضريبي", "السجل", "تاريخ", "العميل", "فاكس", "هاتف", "جوال", "إلى", "رقم الفاتورة", "رقم الغاتورة", "الفاتورة", "الغاتورة", "مدفوع", "مرتجع"]

FINAL_COLS =[
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File",
]

# قاموس لتصحيح أخطاء الـ OCR الطفيفة
SKU_CORRECTIONS = {
    "فيل ليج هندي صاحبة 18 (510)": "فيل ليج هندي صاحبة 18 ك (510)",
    "فيل ليج هندي الفاروق 18": "فيل ليج هندي الفاروق 18 ك",
}

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
    for b in re.findall(r"\(\s*\d+\s*\)", line):
        b_clean = "(" + re.search(r"\d+", b).group() + ")"
        if b_clean not in raw.replace(" ", ""):
            raw = raw + " " + b_clean
            
    sku = clean_sku(raw)
    for wrong, correct in SKU_CORRECTIONS.items():
        if sku == wrong or wrong in sku:
            sku = correct
            break
            
    return sku

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
    ocr_text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img = Image.frombytes("RGB",[pix.width, pix.height], pix.samples)
            ocr_text += pytesseract.image_to_string(img, lang="ara+eng", config="--psm 6") + "\n"
    return ocr_text, "ocr"

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

def get_nums(segment):
    return [n for n in re.findall(r"[\d,]+\.?\d*", segment) if clean_number(n) not in (0, None) and len(re.sub(r"[,.]", "", n)) <= 8]

def parse_item_line(line, tb_val=0.0):
    eng_matches = list(re.finditer(r"[A-Za-z]{2,}", line))
    if eng_matches:
        first_start = eng_matches[0].start()
        last_end    = eng_matches[-1].end()
        before_nums = get_nums(line[:first_start])
        middle_nums = get_nums(line[first_start:last_end])
        after_nums  = get_nums(line[last_end:])
        if len(after_nums) >= 2: all_nums = after_nums
        elif len(middle_nums) >= 1: all_nums = middle_nums + after_nums
        else: all_nums = before_nums + after_nums
    else:
        all_nums = get_nums(line)

    if len(all_nums) < 2: return None

    pack_bracket = set()
    for m in re.finditer(r"\(\s*(\d+)\s*\)", line):
        pack_bracket.add(m.group(1))

    if len(all_nums) == 2:
        candidates =[n for n in all_nums if n not in pack_bracket]
        row_total = None
    else:
        candidates =[n for n in all_nums[:-1] if n not in pack_bracket]
        rt = clean_number(all_nums[-1])
        row_total = rt if rt and rt > 100 else None

    if not candidates: return None

    qty = None
    unit_price = None
    
    cand_floats = [clean_number(n) for n in candidates if clean_number(n)]
    valid_strs = [s for s in candidates if clean_number(s)]
    matched = False

    targets =[]
    if row_total and row_total > 0: targets.append(row_total)
    if tb_val and tb_val > 0: targets.append(tb_val)

    # 1. الاستخراج الذكي بالضرب الرياضي
    for target in targets:
        for i, v1 in enumerate(cand_floats):
            for j, v2 in enumerate(cand_floats):
                if i >= j: continue
                if abs(v1 * v2 - target) / target < 0.05:
                    has_frac1 = not float(v1).is_integer()
                    has_frac2 = not float(v2).is_integer()
                    
                    if has_frac1 or has_frac2:
                        leftovers =[v for idx, v in enumerate(cand_floats) if idx not in (i, j)]
                        
                        # فلترة ذكية لرفض الرقم 18 ككمية إذا كان هناك كمية أوضح (مثل 49)
                        ints =[v for v in leftovers if float(v).is_integer()]
                        if len(ints) > 1 and 18 in ints:
                            ints = [v for v in ints if v != 18]
                            
                        qty = max(ints) if ints else leftovers[0] if leftovers else None
                        unit_price = v1 if has_frac1 and not has_frac2 else v2 if has_frac2 and not has_frac1 else min(v1, v2)
                    else:
                        qty = min(v1, v2)
                        unit_price = max(v1, v2)
                    matched = True
                    break
            if matched: break
        if matched: break

    # 2. الاستخراج في حال فشل الضرب (مثل الفاتورة 2567)
    if not matched:
        has_dot_idx =[idx for idx, s in enumerate(valid_strs) if '.' in s]
        no_dot_idx =[idx for idx, s in enumerate(valid_strs) if '.' not in s]
        
        if has_dot_idx and no_dot_idx:
            unit_price = cand_floats[has_dot_idx[0]]
            
            # تجاهل رقم 18 ككمية إذا وجدنا كمية أخرى (مثل 200)
            possible_qtys = [cand_floats[idx] for idx in no_dot_idx]
            if len(possible_qtys) > 1 and 18 in possible_qtys:
                possible_qtys = [q for q in possible_qtys if q != 18]
                
            qty = max(possible_qtys) if possible_qtys else cand_floats[no_dot_idx[0]]
            
        elif len(cand_floats) >= 2:
            qty = min(cand_floats[0], cand_floats[1])
            unit_price = max(cand_floats[0], cand_floats[1])
        elif cand_floats:
            qty = cand_floats[0]
            unit_price = cand_floats[0]

    # تحويل الكمية لعدد صحيح نظيف
    if qty is not None:
        try:
            qty = int(qty) if float(qty).is_integer() else qty
        except:
            pass

    # استخراج الكلمات الإنجليزية بذكاء
    all_eng = re.findall(r"[A-Za-z]{2,}", line)
    desc_words =[w for w in all_eng if len(w) >= 3 or w.isupper()]
    seen_w, deduped = set(),[]
    for w in desc_words:
        if w.upper() not in seen_w:
            seen_w.add(w.upper())
            deduped.append(w)
    desc = " ".join(deduped).strip()
    
    sku = extract_sku_from_line(line)

    if not (sku or desc): return None
    return {"SKU": sku, "Description": desc, "Quantity": qty, "Unit price": unit_price}

def extract_items_positional(word_df, text, tb_val):
    items =[]
    
    if not word_df.empty:
        rows = reconstruct_table_rows(word_df)
        for row in rows:
            t = row["text"].strip()
            if not t: continue
            
            is_summary = any(kw in t for kw in STOP_KWS)
            has_english = bool(re.search(r'[A-Za-z]{3,}', t))
            
            if is_summary and not has_english and not any(h in t for h in HEADER_KW):
                break
                
            if any(kw in t for kw in SKIP_KWS) or any(kw in t for kw in HEADER_KW):
                continue
                
            parsed = parse_item_line(t, tb_val)
            if parsed: items.append(parsed)

    if not items:
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

    valid_items =[]
    for item in items:
        if len(item.get("Description", "")) < 3 and len(item.get("SKU", "")) < 3:
            continue
        valid_items.append(item)

    return valid_items

def is_summary_row(vals):
    return any(kw in " ".join(vals) for kw in STOP_KWS)

def extract_items_native(pdf_path, tb_val):
    items =[]
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or[]):
                    for row in table:
                        if not row: continue
                        vals =[str(c).strip() if c else "" for c in row]
                        if is_summary_row(vals): continue
                        num_cells =[v for v in vals if re.sub(r"[,.\s]", "", v).isdigit() and 1 <= len(re.sub(r"[,.\s]", "", v)) <= 8]
                        if len(num_cells) < 2: continue
                        
                        raw_sku  = reshape(vals[5]) if len(vals) > 5 else ""
                        raw_desc = reshape(vals[4]) if len(vals) > 4 else ""
                        sku = clean_sku(raw_sku)
                        
                        for wrong, correct in SKU_CORRECTIONS.items():
                            if sku == wrong or wrong in sku:
                                sku = correct
                                break

                        items.append({
                            "Unit price": clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity": clean_number(vals[3]) if len(vals) > 3 else None,
                            "Description": raw_desc,
                            "SKU": sku,
                        })
    except Exception: pass
    return items

def extract_metadata(pdf_path, text):
    cname = ""
    m_name = re.search(r'اسم العميل\s*:\s*(.*?)(?=رقم|التاريخ|الرقم|\n)', text)
    if m_name:
        cname = m_name.group(1).strip()
        cname = re.sub(r'الغاتورة.*|الفاتورة.*|الفغاتورة.*|إلى.*', '', cname).strip()

    inv_num = ""
    m_inv = re.search(r'رقم\s*(?:ال[غف]اتورة|الفغاتورة|فاتورة)\s*[:\-]?\s*(\d{4,6})', text)
    if not m_inv:
        m_inv = re.search(r'رقم.*?\s+(\d{4,6})\b', text)
    if m_inv: inv_num = m_inv.group(1).strip()

    inv_date = ""
    m_date = re.search(r'تاريخ.*?\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})', text)
    if m_date: inv_date = m_date.group(1).strip()

    address = ""
    m_add = re.search(r'العنوان\s*:\s*(.+?)(?=\n\s*05|\n\s*\d{10}|\n\s*البند|\n\s*المجموع|05\d{8}|فيل|كبدة|عجل|فخده|فوركوارتر)', text, re.DOTALL)
    if m_add:
        address = m_add.group(1).replace('\n', ' ').strip()
        address = re.sub(r'\s*\d{10}\s*$', '', address).strip()

    tb = ta = vat = paid = bal = 0.0

    safe_text = re.sub(r'SA\d{22}', '', text)
    safe_text = re.sub(r'\b\d{10,}\b', '', safe_text)

    m_tot = re.search(r'الإ[جح]مالي\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_tot: ta = clean_number(m_tot.group(1))

    m_sub = re.search(r'المجموع\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_sub: tb = clean_number(m_sub.group(1))

    m_vat = re.search(r'(?:القيمة المضافة|المضافة|15%)\s*[:\-]?\s*([\d,]+\.?\d*)', safe_text)
    if m_vat: vat = clean_number(m_vat.group(1))

    if not ta or not tb:
        nums_raw =[]
        for n in re.findall(r"[\d,]+\.?\d*", safe_text):
            v = clean_number(n)
            if v and v > 100: nums_raw.append(v)
        
        unique = sorted(set(nums_raw))
        best_diff = float("inf")
        found_ta, found_tb = ta, tb
        
        for i, small in enumerate(unique):
            for big in unique[i + 1:]:
                r = big / small
                if 1.10 <= r <= 1.20:
                    diff = abs(r - 1.15)
                    if diff < best_diff:
                        best_diff = diff
                        found_tb = small
                        found_ta = big
                        
        if not ta and found_ta: ta = found_ta
        if not tb and found_tb: tb = found_tb

    if ta:
        expected_tb = round(ta / 1.15, 2)
        expected_vat = round(ta - expected_tb, 2)
        if not tb or abs(tb - expected_tb) > 2: tb = expected_tb
        if not vat or abs(vat - expected_vat) > 2: vat = expected_vat
    elif tb:
        ta = round(tb * 1.15, 2)
        vat = round(ta - tb, 2)

    paid = 0.0
    bal = ta if ta else 0.0

    return {
        "Invoice Number": inv_num,
        "Invoice Date": inv_date,
        "Customer Name": cname,
        "Address": address,
        "Balance": bal,
        "Paid": paid,
        "Total before tax": tb,
        "VAT 15%": vat,
        "Total after tax": ta,
        "Source File": pdf_path.name,
    }

def process_pdf(pdf_path):
    text, mode = get_text(pdf_path)
    meta = extract_metadata(pdf_path, text)
    tb_val = meta.get("Total before tax", 0.0)

    if mode == "ocr":
        word_df = get_ocr_words(pdf_path)
        items   = extract_items_positional(word_df, text, tb_val)
    else:
        word_df = pd.DataFrame()
        items   = extract_items_native(pdf_path, tb_val)
        if not items:
            items = extract_items_positional(pd.DataFrame(), text, tb_val)

    file_cname = extract_name_from_filename(pdf_path)
    if file_cname and len(file_cname) > 3:
        meta["Customer Name"] = file_cname

    if not meta["Invoice Number"]:
        m_fname_inv = re.search(r'(\d{4,6})', pdf_path.stem)
        if m_fname_inv:
            meta["Invoice Number"] = m_fname_inv.group(1)

    # 💡 تم مسح نظام الحذف تماماً، الكود سيستخرج كل السطور كما هي بدون حذف
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
            st.dataframe(final_df)

            out = BytesIO()
            final_df.to_excel(out, index=False, engine="openpyxl")
            out.seek(0)
            st.download_button(
                "📥 Download Excel",
                out,
                "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ No data extracted.")
