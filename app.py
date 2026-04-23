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
        return float(s) if s else None
    except:
        return None


def extract_name_from_filename(pdf_path):
    stem = Path(pdf_path).stem
    name = re.sub(r"[-_\s]*\d+[-_\s]*$", "", stem).strip()
    name = re.sub(r"^[-_\s]*\d+[-_\s]*", "", name).strip()
    if re.search(r"[\u0600-\u06FF]", name):
        return name
    return ""


UNIT_WORDS = {
    "\u0643\u0631\u062a\u0648\u0646\u0629", "\u0643\u0631\u062a\u0648\u0646",
    "\u0642\u0637\u0639\u0629", "\u0639\u0644\u0628\u0629", "\u0643\u064a\u0633",
    "\u0637\u0646", "\u0643\u062c\u0645", "\u0644\u062a\u0631", "\u0643\u063a",
    "\u062c\u0631\u0627\u0645", "\u0645\u0644", "\u062d\u0628\u0629",
    "\u0631\u0648\u0644", "\u0628\u0627\u0643\u064a\u062a", "\u0635\u0646\u062f\u0648\u0642",
}

HEADER_KW = [
    "\u0627\u0644\u0628\u0646\u062f", "\u0627\u0644\u0648\u0635\u0641",
    "\u0627\u0644\u0639\u062f\u062f", "\u0633\u0639\u0631 \u0627\u0644\u0648\u062d\u062f\u0629",
    "\u0627\u0644\u0643\u0645\u064a\u0629", "\u0627\u0644\u0648\u062d\u062f\u0629",
]

SUMMARY_KW = [
    "\u0627\u0644\u0645\u062c\u0645\u0648\u0639", "\u0645\u062f\u0641\u0648\u0639",
    "\u0627\u0644\u0631\u0635\u064a\u062f", "\u0627\u0644\u0642\u064a\u0645\u0629",
    "\u0627\u0644\u0642\u064a\u0645\u0647", "\u0627\u0644\u0625\u062c\u0645\u0627\u0644\u064a",
    "\u0627\u0644\u0625\u062d\u0645\u0627\u0644\u064a", "\u0627\u0625\u0644\u062c\u0645\u0627\u0644\u064a",
    "\u0627\u0644\u0627\u062c\u0645\u0627\u0644\u064a",
    "\u0631\u0642\u0645 \u0627\u0644\u062d\u0633\u0627\u0628",
    "\u0627\u0644\u0627\u064a\u0628\u0627\u0646", "IBAN", "SA08", "Kingdome",
    "\u0627\u0644\u0645\u0645\u0644\u0643\u0629",
    "\u0631\u0642\u0645 \u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629",
    "\u062a\u0627\u0631\u064a\u062e", "\u0627\u0633\u0645 \u0627\u0644\u0639\u0645\u064a\u0644",
    "\u0627\u0644\u0631\u0642\u0645 \u0627\u0644\u0636\u0631\u064a\u0628\u064a",
    "\u0631\u0642\u0645 \u0627\u0644\u0633\u062c\u0644", "\u0627\u0644\u0639\u0646\u0648\u0627\u0646",
    "\u0627\u0644\u062c\u0648\u0627\u0644", "\u0627\u0644\u0633\u062c\u0644 \u0627\u0644\u062a\u062c\u0627\u0631\u064a",
    "\u0645\u0631\u062a\u062c\u0639",
]

FINAL_COLS = [
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File",
]

SKU_TO_DESC = {
    "\u0641\u064a\u0644 \u0644\u064a\u062c \u0647\u0646\u062f\u064a \u0635\u0627\u062d\u0628\u0629": "VEAL LEG SAHIBA",
    "\u0641\u064a\u0644 \u0644\u064a\u062c \u0647\u0646\u062f\u064a": "VEAL LEG HINDI",
    "\u0641\u064a\u0644 \u0644\u064a\u062c": "VEAL LEG",
    "\u0641\u0648\u0631\u0643\u0648\u0627\u0631\u062a\u0631 \u0647\u0646\u062f\u064a": "FOREQUARTER HINDI",
    "\u0641\u0648\u0631\u0643\u0648\u0627\u0631\u062a\u0631": "FOREQUARTER",
}


def clean_sku(raw_sku):
    cleaned = re.sub(r"\|", " ", raw_sku)
    words = [w for w in cleaned.split()
             if w not in UNIT_WORDS and (len(w) > 1 or w == "\u0643")]
    return " ".join(words).strip()


def extract_sku_from_line(line):
    ar_block = re.search(r"([\u0600-\u06FF][\u0600-\u06FF\s\d\(\)\u0643]*)", line)
    raw = ar_block.group(1).strip() if ar_block else ""
    if not raw:
        ar_words = re.findall(r"[\u0600-\u06FF]{2,}", line)
        raw = " ".join(w for w in ar_words if w not in UNIT_WORDS)
    for b in re.findall(r"\(\s*\d+\s*\)", line):
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
            ocr_text += f"\n--- \u0627\u0644\u0635\u0641\u062d\u0629 {i+1} ---\n{page_text}\n"
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


def extract_customer_name_text(text):
    m = re.search(
        r"\u0627\u0633\u0645 \u0627\u0644\u0639\u0645\u064a\u0644\s*[:\s]+(.+?)"
        r"(?=\u0627\u0644\u0631\u0642\u0645 \u0627\u0644\u0636\u0631\u064a\u0628\u064a"
        r"|\u0631\u0642\u0645 \u0627\u0644\u0633\u062c\u0644|\u0627\u0644\u0639\u0646\u0648\u0627\u0646)",
        text, re.DOTALL,
    )
    if not m:
        return ""
    chunk = m.group(1)
    chunk = re.sub(r"(?:\u0642\u0645 \u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629|\u0631\u0642\u0645 \u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629)\s*\d+", "", chunk)
    chunk = re.sub(r"\d{4,}", "", chunk)
    chunk = re.sub(r"[a-zA-Z]{2,}", "", chunk)
    stop = {"\u0627\u0633\u0645", "\u0627\u0644\u0639\u0645\u064a\u0644", "\u0641\u0627\u062a\u0648\u0631\u0629",
            "\u0625\u0644\u0649", "\u0625\u0644\u0649\u0629", "\u0627\u0644\u062a\u062a\u062c\u0627\u0631",
            "\u0631\u0642\u0645", "\u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629",
            "\u062a\u0627\u0631\u064a\u062e", "\u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629", "\u0625\u0644\u0649\u0629"}
    arabic_words = [w for w in re.findall(r"[\u0600-\u06FF]+", chunk) if w not in stop and len(w) > 1]
    seen = set()
    unique = []
    for w in arabic_words:
        if w not in seen:
            seen.add(w)
            unique.append(w)
    return " ".join(unique).strip()


def get_nums(segment):
    return [
        n for n in re.findall(r"[\d,]+\.?\d*", segment)
        if clean_number(n) not in (0, None)
        and len(re.sub(r"[,.]", "", n)) <= 8
    ]


def parse_item_line(line):
    """
    Extracts SKU, Description, Quantity, Unit price from one invoice product line.

    Number routing:
      Layout A (numbers after last English word):  after_nums >= 2 → use after_nums
      Layout B (numbers between English words):     middle_nums >= 1 → middle + after
      Fallback:                                     before + after

    Price / Qty assignment (column-order first, then cross-check):
      Step 1: first_whole=qty_candidate, first_decimal=price_candidate
              If first_whole * first_decimal ≈ row_total → accept (handles most cases)
      Step 2: If step 1 fails, find ANY (v1,v2) pair where v1*v2 ≈ row_total
              Use position order: earlier candidate = qty, later = price
      Step 3: Fallback: first whole = qty, first decimal = price

    Description: collect all English words ≥ 4 chars from entire line (reduces noise)
    SKU: Arabic block + any (NNN) bracket appended if missing from Arabic zone
    """
    eng_matches = list(re.finditer(r"[A-Za-z]{2,}", line))

    if eng_matches:
        first_start = eng_matches[0].start()
        last_end    = eng_matches[-1].end()
        before_nums = get_nums(line[:first_start])
        middle_nums = get_nums(line[first_start:last_end])
        after_nums  = get_nums(line[last_end:])

        if len(after_nums) >= 2:
            all_nums = after_nums
        elif len(middle_nums) >= 1:
            all_nums = middle_nums + after_nums
        else:
            all_nums = before_nums + after_nums
    else:
        all_nums = get_nums(line)

    if len(all_nums) < 2:
        return None

    # Strip pack-count brackets like (510)
    pack_bracket = set()
    for m in re.finditer(r"\(\s*(\d+)\s*\)", line):
        pack_bracket.add(m.group(1))

    if len(all_nums) == 2:
        candidates = [n for n in all_nums if n not in pack_bracket]
        row_total = None
    else:
        candidates = [n for n in all_nums[:-1] if n not in pack_bracket]
        rt = clean_number(all_nums[-1])
        row_total = rt if rt and rt > 100 else None

    if not candidates:
        return None

    # Column-order: first whole = qty, first decimal = price
    first_whole   = next((clean_number(n) for n in candidates if "." not in str(n) and clean_number(n)), None)
    first_decimal = next((clean_number(n) for n in candidates if "." in     str(n)), None)

    qty = first_whole
    unit_price = first_decimal

    # Validate against row total
    if row_total and qty and unit_price:
        if abs(qty * unit_price - row_total) / row_total > 0.05:
            # Column-order doesn't match — search for matching pair
            cand_vals = [clean_number(n) for n in candidates if clean_number(n)]
            best_diff = float("inf")
            found = False
            for i, v1 in enumerate(cand_vals):
                for j, v2 in enumerate(cand_vals):
                    if i == j:
                        continue
                    if row_total > 0 and abs(v1 * v2 - row_total) / row_total < 0.05:
                        diff = abs(v1 * v2 - row_total)
                        if diff < best_diff:
                            best_diff = diff
                            # Earlier position = qty, later = price
                            if i < j:
                                qty, unit_price = v1, v2
                            else:
                                qty, unit_price = v2, v1
                            found = True
            # If still no match, keep column-order defaults

    # Final fallback: no decimal found
    if unit_price is None:
        vals = sorted([clean_number(n) for n in candidates if clean_number(n)])
        if len(vals) >= 2:
            qty = vals[0]
            unit_price = vals[-1]
        elif vals:
            unit_price = vals[0]

    # Description: English words ≥ 4 chars (reduces short OCR noise like 'ghd', 'pxS')
    all_eng = re.findall(r"[A-Za-z]{2,}", line)
    desc_words = [w for w in all_eng if len(w) >= 4 or w.isupper()]
    seen_w: set = set()
    deduped = []
    for w in desc_words:
        key = w.upper()
        if key not in seen_w:
            seen_w.add(key)
            deduped.append(w)
    desc = " ".join(deduped).strip()

    # SKU
    sku = extract_sku_from_line(line)

    # Description fallback via SKU lookup
    if sku and desc:
        for ar_key, en_val in SKU_TO_DESC.items():
            if ar_key in sku:
                dwords = desc.upper().split()
                if all(w in en_val.upper() for w in dwords) and desc.upper() != en_val.upper():
                    desc = en_val
                break
    elif sku and not desc:
        for ar_key, en_val in SKU_TO_DESC.items():
            if ar_key in sku:
                desc = en_val
                break

    if not (sku or desc):
        return None

    return {"SKU": sku, "Description": desc, "Quantity": qty, "Unit price": unit_price}


def extract_items_positional(word_df, text):
    items = []

    if not word_df.empty:
        rows = reconstruct_table_rows(word_df)
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
        if header_idx is not None and summary_idx is not None:
            for row in rows[header_idx + 1: summary_idx]:
                t = row["text"].strip()
                if not t or any(kw in t for kw in SUMMARY_KW):
                    continue
                parsed = parse_item_line(t)
                if parsed:
                    items.append(parsed)

    if not items:
        lines = text.split("\n")
        in_table = False
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if any(h in line for h in HEADER_KW):
                in_table = True
                continue
            if in_table and any(kw in line for kw in [
                "\u0627\u0644\u0645\u062c\u0645\u0648\u0639",
                "\u0627\u0644\u0642\u064a\u0645\u0629",
                "\u0627\u0644\u0625\u062c\u0645\u0627\u0644\u064a",
                "\u0627\u0644\u0625\u062d\u0645\u0627\u0644\u064a",
                "\u0645\u062f\u0641\u0648\u0639",
                "\u0645\u0631\u062a\u062c\u0639",
            ]):
                break
            if not in_table:
                continue
            parsed = parse_item_line(line)
            if parsed:
                items.append(parsed)

    if not items:
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            if any(kw in line for kw in SUMMARY_KW + HEADER_KW):
                continue
            has_arabic  = bool(re.search(r"[\u0600-\u06FF]{2,}", line))
            has_english = bool(re.search(r"[A-Za-z]{2,}", line))
            has_nums    = len(re.findall(r"[\d,]+\.?\d*", line)) >= 2
            if not (has_arabic and has_english and has_nums):
                continue
            parsed = parse_item_line(line)
            if parsed and (parsed["SKU"] or parsed["Description"]):
                items.append(parsed)
                break

    return items


def is_summary_row(vals):
    return any(kw in " ".join(vals) for kw in SUMMARY_KW)


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
                        num_cells = [
                            v for v in vals
                            if re.sub(r"[,.\s]", "", v).isdigit()
                            and 1 <= len(re.sub(r"[,.\s]", "", v)) <= 8
                        ]
                        if len(num_cells) < 2:
                            continue
                        raw_sku  = reshape(vals[5]) if len(vals) > 5 else ""
                        raw_desc = reshape(vals[4]) if len(vals) > 4 else ""
                        items.append({
                            "Unit price":  clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity":    clean_number(vals[3]) if len(vals) > 3 else None,
                            "Description": raw_desc,
                            "SKU":         clean_sku(raw_sku),
                        })
    except Exception:
        pass
    return items


def extract_financials(text):
    """
    Ratio-based financial extraction (handles split-column OCR layouts).
    Finds number pairs where large/small ~1.15 (the 15% VAT relationship).
    Paid is always set to None (user requested Paid=0 always — handled in process_pdf).
    """
    pos = max(text.find("\u0627\u0644\u0645\u062c\u0645\u0648\u0639"), 0)
    fin = text[pos:]

    nums_raw = []
    for n in re.findall(r"[\d,]+\.?\d*", fin):
        v = clean_number(n)
        if v and v > 100 and v != 15 and v != 150:
            nums_raw.append(v)
    unique = sorted(set(nums_raw))

    tb = ta = vat = None

    # Strategy 1: pair with ratio ~1.15
    best = float("inf")
    for i, small in enumerate(unique):
        for big in unique[i + 1:]:
            r = big / small
            if 1.10 <= r <= 1.20:
                s = abs(r - 1.15)
                if s < best:
                    best = s
                    tb = small
                    ta = big

    # Strategy 2: (vat, ta) where vat/ta ~0.13
    if ta is None:
        for big in reversed(unique):
            for small in unique:
                if small < big and 0.125 <= small / big <= 0.135:
                    ta = big
                    vat = small
                    tb = round(big - small, 2)
                    break
            if ta:
                break

    # Strategy 3: largest number = ta
    if ta is None and unique:
        ta = max(unique)
        tb = round(ta / 1.15, 2)

    if tb and ta and not vat:
        vat = round(ta - tb, 2)
    if vat and ta and not tb:
        tb = round(ta - vat, 2)
    if tb and vat and not ta:
        ta = round(tb + vat, 2)

    return (
        round(tb, 2) if tb else None,
        round(vat, 2) if vat else None,
        round(ta, 2) if ta else None,
    )


def extract_metadata(pdf_path, text):
    inv_num = ""
    for p in [
        r"\u0631\u0642\u0645 \u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629\s*[:\s]*(\d{4,6})",
        r"(?:\u0642\u0645 \u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629|\u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629)\s*(\d{4,6})",
        r"\b(0\d{4,5})\b",
    ]:
        m = re.search(p, text)
        if m:
            inv_num = m.group(1).strip()
            break

    date_m   = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", text)
    inv_date = date_m.group(1) if date_m else ""

    address = ""
    m = re.search(
        r"\u0627\u0644\u0639\u0646\u0648\u0627\u0646\s*[:\s]+(.+?)"
        r"(?=\n\d{7,}|\n(?:\u0627\u0644\u0645\u062c\u0645\u0648\u0639|\u0645\u062f\u0641\u0648\u0639|\u0631\u0642\u0645 \u0627\u0644\u062d\u0633\u0627\u0628)|\Z)",
        text, re.DOTALL,
    )
    if m:
        address = " ".join(m.group(1).split()).strip()
        address = re.sub(r"\s*\d{10}\s*$", "", address).strip()

    total_before, vat, total_after = extract_financials(text)

    return {
        "Invoice Number":   inv_num,
        "Invoice Date":     inv_date,
        "Address":          address,
        "Balance":          total_after,
        "Paid":             0,          # always 0 as requested
        "Total before tax": total_before,
        "VAT 15%":          vat,
        "Total after tax":  total_after,
        "Source File":      pdf_path.name,
    }


def process_pdf(pdf_path):
    text, mode = get_text(pdf_path)
    meta = extract_metadata(pdf_path, text)

    if mode == "ocr":
        word_df = get_ocr_words(pdf_path)
        items   = extract_items_positional(word_df, text)
    else:
        word_df = pd.DataFrame()
        items   = extract_items_native(pdf_path)
        if not items:
            items = extract_items_positional(pd.DataFrame(), text)

    cname = extract_name_from_filename(pdf_path)
    if not cname or len(cname) < 4:
        cname = extract_customer_name_text(text)
    if not cname or len(cname) < 4:
        cname = ""
    meta["Customer Name"] = cname

    seen = set()
    unique_items = []
    for item in items:
        key = (item.get("Description", ""), item.get("Unit price"), item.get("Quantity"))
        if key not in seen:
            seen.add(key)
            unique_items.append(item)

    if not unique_items:
        unique_items = [{"Unit price": None, "Quantity": None, "Description": "", "SKU": ""}]

    rows = [{**meta, **item} for item in unique_items]
    return pd.DataFrame(rows).reindex(columns=FINAL_COLS), mode, text


st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("\U0001f4c4 Invoice Extractor \u2014 PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True
)
debug_mode = st.checkbox("\U0001f50d Show full raw extracted text", value=False)

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
            st.write(f"\U0001f4c4 **{path.name}**")
            with st.spinner("Extracting..."):
                df, mode, raw_text = process_pdf(path)
            st.caption(f"Mode: `{mode}` \u2014 {len(df)} row(s)")

            if debug_mode:
                with st.expander(f"\U0001f4cb Full raw text \u2014 {path.name}", expanded=True):
                    st.text(raw_text)
                    st.caption(f"Total characters: {len(raw_text)}")
                    st.download_button(
                        label="\U0001f4cb Download raw text",
                        data=raw_text,
                        file_name=f"{path.stem}_raw.txt",
                        mime="text/plain",
                        key=f"raw_dl_{i}_{path.stem}",
                    )

            if not df.empty:
                all_data.append(df)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)

            if "Invoice Date" in final_df.columns:
                final_df["Invoice Date"] = pd.to_datetime(
                    final_df["Invoice Date"], errors="coerce", dayfirst=True
                ).dt.strftime("%m/%d/%Y")

            st.success(f"\u2705 Done! {len(final_df)} total row(s)")
            st.dataframe(final_df)

            out = BytesIO()
            final_df.to_excel(out, index=False, engine="openpyxl")
            out.seek(0)
            st.download_button(
                "\U0001f4e5 Download Excel",
                out,
                "Invoices.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("\u26a0\ufe0f No data extracted.")
