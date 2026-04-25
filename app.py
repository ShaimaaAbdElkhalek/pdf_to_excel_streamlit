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

SKU_TO_DESC = {
    "فيل ليج هندي صاحبة": "VEAL LEG SAHIBA",
    "فيل ليج هندي": "VEAL LEG HINDI",
    "فيل ليج": "VEAL LEG",
    "فوركوارتر هندي": "FOREQUARTER HINDI",
    "فوركوارتر": "FOREQUARTER",
}

# 💡 قاموس التصحيح التلقائي لحل مشكلة الـ OCR في الفاتورة الثالثة وتوحيد الاسم
SKU_CORRECTIONS = {
    "فيل ليج هندي صاحبة 18 (510)": "فيل ليج هندي صاحبة 18 ك (510)",
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
    # تطبيق التصحيح
    for wrong, correct in SKU_CORRECTIONS.items():
        if sku == wrong or wrong in sku:
            sku = correct
            
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
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
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
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    data = pytesseract.image_to_data(
        img, lang="ara+eng", config="--psm 6",
        output_type=pytesseract.Output.DATAFRAME,
    )
    data = data[data["conf"] > 30].dropna(subset=["text"])
    data = data[data["text"].str.strip() != ""]
    return data

def reconstruct_table_rows(word_df, y_tolerance=15):
    if word_df.empty: return []
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

    # 💡 خوارزمية ذكية لاستخراج العدد والسعر (تعرف السعر من العلامة العشرية)
    for target in targets:
        for i, v1 in enumerate(cand_floats):
            for j, v2 in enumerate(cand_floats):
                if i >= j: continue
                if abs(v1 * v2 - target) / target < 0.05:
                    leftovers =[v for idx, v in enumerate(cand_floats) if idx not in (i, j)]
                    if leftovers:
                        qty = leftovers[0]
                        s1, s2 = valid_strs[i], valid_strs[j]
                        # الرقم الذي يحتوى على فاصلة عشرية هو السعر
                        if '.' in s1 and '.' not in s2:
                            unit_price = v1
                        elif '.' in s2 and '.' not in s1:
                            unit_price = v2
                        else:
                            unit_price = min(v1, v2)
                    else:
                        s1, s2 = valid_strs[i], valid_strs[j]
                        if float(v1).is_integer() and not float(v2).is_integer():
                            qty, unit_price = v1, v2
                        elif float(v2).is_integer() and not float(v1).is_integer():
                            qty, unit_price = v2, v1
                        else:
                            if '.' in s1 and '.' not in s2:
                                qty, unit_price = v2, v1
                            elif '.' in s2 and '.' not in s1:
                                qty, unit_price = v1, v2
                            else:
                                qty, unit_price = min(v1, v2), max(v1, v2)
                    matched = True
                    break
            if matched: break
        if matched: break

    # حالة احتياطية (Fallback)
    if not matched:
        has_dot =[idx for idx, s in enumerate(valid_strs) if '.' in s]
        if len(has_dot) == 1:
            unit_price = cand_floats[has_dot[0]]
            others =[v for idx, v in enumerate(cand_floats) if idx != has_dot[0]]
            qty = max(others) if others else unit_price
        elif len(cand_floats) >= 2:
            qty = cand_floats[0]
            unit_price = cand_floats[1]
        elif cand_floats:
            qty = cand_floats[0]
            unit_price = cand_floats[0]

    if qty is not None:
        try:
            qty = int(qty) if float(qty).is_integer() else qty
        except:
            pass

    all_eng = re.findall(r"[A-Za-z]{2,}", line)
    desc_words =[w for w in all_eng if len(w) >= 4 or w.isupper()]
    seen_w, deduped = set(),[]
    for w in desc_words:
        if w.upper() not in seen_w:
            seen_w.add(w.upper())
            deduped.append(w)
    desc = " ".join(deduped).strip()
    sku = extract_sku_from_line(line)

    if sku and desc:
        for ar_key, en_val in SKU_TO_DESC.items():
            if ar_key in sku:
                dwords = desc.upper().split()
                if all(w in en_val.upper() for w in dwords) and desc.
