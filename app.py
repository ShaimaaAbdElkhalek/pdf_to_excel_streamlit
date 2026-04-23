import asyncio
import re
import tempfile
import unicodedata
import zipfile
from io import BytesIO
from pathlib import Path

import fitz
import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image

try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    def reshape(text):
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
except ImportError:
    def reshape(text):
        return text


# ---------------------------------------------------------------------------
# Arabic normalization
# ---------------------------------------------------------------------------

def normalize_arabic(text: str) -> str:
    """Convert Arabic presentation forms (FExxx) to standard codepoints."""
    return unicodedata.normalize("NFKC", text)


def fix_ocr_numbers(text: str) -> str:
    """Fix common OCR artifacts in numeric contexts (I→1, O→0, l→1)."""
    # Replace capital I/O only when surrounded by digits or at start of digit run
    text = re.sub(r"(?<!\w)([IO])(\d)", lambda m: ("1" if m.group(1) == "I" else "0") + m.group(2), text)
    text = re.sub(r"(\d)([IO])(?!\w)", lambda m: m.group(1) + ("1" if m.group(2) == "I" else "0"), text)
    return text


# ---------------------------------------------------------------------------
# Windows OCR (winocr) — replaces pytesseract
# ---------------------------------------------------------------------------

async def _winocr_async(img: Image.Image):
    import winocr
    return await winocr.recognize_pil(img, "ar")


def run_winocr(img: Image.Image):
    """Synchronous wrapper; safe whether or not an event loop is already running."""
    try:
        # No running loop → standard path
        asyncio.get_running_loop()
        # A loop IS running (e.g. Streamlit) → run in a fresh thread
        import concurrent.futures
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
            future = ex.submit(asyncio.run, _winocr_async(img))
            return future.result(timeout=60)
    except RuntimeError:
        # No running loop
        return asyncio.run(_winocr_async(img))


def pdf_page_to_image(pdf_path, page_index: int = 0, scale: float = 3.0) -> Image.Image:
    with fitz.open(str(pdf_path)) as doc:
        pix = doc[page_index].get_pixmap(matrix=fitz.Matrix(scale, scale))
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean_number(val):
    s = re.sub(r"[^\d.]", "", str(val).replace(",", "").replace("\u066c", "").replace("\u066b", "."))
    try:
        return float(s) if s else None
    except Exception:
        return None


def extract_name_from_filename(pdf_path) -> str:
    stem = Path(pdf_path).stem
    name = re.sub(r"[-_\s]*\d+[-_\s]*$", "", stem).strip()
    name = re.sub(r"^[-_\s]*\d+[-_\s]*", "", name).strip()
    return name if re.search(r"[\u0600-\u06FF]", name) else ""


UNIT_WORDS = {
    "\u0643\u0631\u062a\u0648\u0646\u0629", "\u0643\u0631\u062a\u0648\u0646",
    "\u0642\u0637\u0639\u0629", "\u0639\u0644\u0628\u0629", "\u0643\u064a\u0633",
    "\u0637\u0646", "\u0643\u062c\u0645", "\u0644\u062a\u0631", "\u0643\u063a",
    "\u062c\u0631\u0627\u0645", "\u0645\u0644", "\u062d\u0628\u0629",
    "\u0631\u0648\u0644", "\u0628\u0627\u0643\u064a\u062a", "\u0635\u0646\u062f\u0648\u0642",
    "\u0643\u0631\u0646\u0648\u0646\u0629",  # extra OCR variant كرنونة
}

# Keywords that mark the table header row
HEADER_KW = [
    "\u0627\u0644\u0628\u0646\u062f", "\u0627\u0644\u0648\u0635\u0641",
    "\u0627\u0644\u0639\u062f\u062f", "\u0633\u0639\u0631 \u0627\u0644\u0648\u062d\u062f\u0629",
    "\u0627\u0644\u0643\u0645\u064a\u0629", "\u0627\u0644\u0648\u062d\u062f\u0629",
    "\u0627\u0644\u0648\u0635\u0641",  # الوصف
]

# Keywords that mark the financial summary section
SUMMARY_KW = [
    "\u0627\u0644\u0645\u062c\u0645\u0648\u0639",  # المجموع
    "\u0645\u062f\u0641\u0648\u0639",              # مدفوع
    "\u0645\u062f\u0647\u0648\u0639",              # مدهوع (OCR variant)
    "\u0627\u0644\u0631\u0635\u064a\u062f",        # الرصيد
    "\u0627\u0644\u0642\u064a\u0645\u0629",        # القيمة
    "\u0627\u0644\u0642\u064a\u0645\u0647",
    "\u0627\u0644\u0625\u062c\u0645\u0627\u0644\u064a", # الإجمالي
    "\u0627\u0644\u0625\u062d\u0645\u0627\u0644\u064a",
    "\u0627\u0644\u0627\u062c\u0645\u0627\u0644\u064a",
    "\u0631\u0642\u0645 \u0627\u0644\u062d\u0633\u0627\u0628",
    "\u0627\u0644\u0627\u064a\u0628\u0627\u0646", "IBAN", "SA08",
    "\u0627\u0644\u0645\u0645\u0644\u0643\u0629", "Kingdome",
    "\u0631\u0642\u0645 \u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629",
    "\u062a\u0627\u0631\u064a\u062e", "\u0627\u0633\u0645 \u0627\u0644\u0639\u0645\u064a\u0644",
    "\u0627\u0644\u0631\u0642\u0645 \u0627\u0644\u0636\u0631\u064a\u0628\u064a",
    "\u0631\u0642\u0645 \u0627\u0644\u0633\u062c\u0644",
    "\u0627\u0644\u0633\u062c\u0644 \u0627\u0644\u062a\u062c\u0627\u0631\u064a",
    "\u0645\u0631\u062a\u062c\u0639",
    # OCR-garbled variants
    "\u0627\u0644\u0645\u062d\u0645\u0648\u062c",  # المحموج (garbled المجموع)
    "\u0627\u0644\u0645\u062d\u0645\u0648\u0639",  # المحموع (OCR variant)
    "\u0627\u0644\u0635\u0627\u062d\u0647",        # المصاحه (garbled)
    "\u0627\u0644\u0645\u0645\u0647",              # الممه (garbled)
    "\u0627\u0644\u0625\u062d\u0645\u0627\u0644\u064a",
    "\u0645\u0631\u062a\u062d\u0639",              # مرتحع (OCR variant)
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


def clean_sku(raw_sku: str) -> str:
    cleaned = re.sub(r"\|", " ", raw_sku)
    words = [w for w in cleaned.split()
             if w not in UNIT_WORDS and (len(w) > 1 or w == "\u0643")]
    return " ".join(words).strip()


def extract_sku_from_line(line: str) -> str:
    ar_block = re.search(r"([\u0600-\u06FF][\u0600-\u06FF\s\d()\u0643]*)", line)
    raw = ar_block.group(1).strip() if ar_block else ""
    if not raw:
        ar_words = re.findall(r"[\u0600-\u06FF]{2,}", line)
        raw = " ".join(w for w in ar_words if w not in UNIT_WORDS)
    for b in re.finditer(r"\(\s*\d+\s*\)", line):
        b_clean = "(" + re.search(r"\d+", b.group()).group() + ")"
        if b_clean not in raw.replace(" ", ""):
            raw = raw + " " + b_clean
    return clean_sku(raw)


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def get_text(pdf_path):
    """Return (text, mode) where mode is 'native' or 'ocr'."""
    with fitz.open(str(pdf_path)) as doc:
        raw = "\n".join(page.get_text() for page in doc).strip()

    if raw:
        raw = normalize_arabic(raw)
        raw = fix_ocr_numbers(raw)

    if len(raw) > 50:
        return raw, "native"

    # Image-based PDF — use Windows OCR
    try:
        img = pdf_page_to_image(pdf_path, scale=3.0)
        result = run_winocr(img)
        lines = [" ".join(w.text for w in line.words) for line in result.lines]
        text = fix_ocr_numbers("\n".join(lines))
        return text, "ocr"
    except Exception:
        return "", "ocr"


def get_ocr_words(pdf_path) -> pd.DataFrame:
    """Return DataFrame of words with bounding boxes from Windows OCR."""
    try:
        img = pdf_page_to_image(pdf_path, scale=3.0)
        result = run_winocr(img)
        rows = []
        for line in result.lines:
            for word in line.words:
                r = word.bounding_rect
                rows.append({
                    "left": r.x,
                    "top": r.y,
                    "width": r.width,
                    "height": r.height,
                    "text": fix_ocr_numbers(word.text),
                    "conf": 90,
                })
        return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame()


# ---------------------------------------------------------------------------
# Table row reconstruction
# ---------------------------------------------------------------------------

def reconstruct_table_rows(word_df: pd.DataFrame, y_tolerance: int = 15):
    if word_df.empty:
        return []
    word_df = word_df.copy()
    word_df["mid_y"] = word_df["top"] + word_df["height"] / 2
    rows = []
    used: set = set()
    for idx, word in word_df.iterrows():
        if idx in used:
            continue
        y = word["mid_y"]
        same_row = word_df[abs(word_df["mid_y"] - y) <= y_tolerance]
        used.update(same_row.index)
        # Sort left-to-right so English words inside Arabic lines keep correct order
        same_row = same_row.sort_values("left", ascending=True)
        row_text = " ".join(same_row["text"].astype(str).tolist())
        rows.append({"y": y, "text": row_text, "words": same_row})
    rows.sort(key=lambda r: r["y"])
    return rows


# ---------------------------------------------------------------------------
# Customer name extraction
# ---------------------------------------------------------------------------

def extract_customer_name_text(text: str) -> str:
    """Extract customer name from invoice text."""
    # Pattern handles both standard Arabic and OCR-garbled variants
    m = re.search(
        r"(?:\u0627\u0633\u0645|\u0627\u0633\u0645\u0647)"   # اسم / اسمه
        r"[\s\n]*"
        r"(?:\u0627\u0644\u0639\u0645\u064a\u0644|\u0627\u0644\u0639\u0645\u0628\u0644)"  # العميل / العمبل
        r"[\s\n:]*(.+?)(?=\n|:|\u0627\u0644\u0631\u0642\u0645|\u0627\u0644\u0636\u0631\u064a)",
        text, re.DOTALL,
    )
    if not m:
        return ""
    chunk = m.group(1).strip()
    chunk = re.sub(r"\d{4,}", "", chunk)
    stop = {
        "\u0627\u0633\u0645", "\u0627\u0644\u0639\u0645\u064a\u0644",
        "\u0641\u0627\u062a\u0648\u0631\u0629", "\u0625\u0644\u0649",
        "\u0631\u0642\u0645", "\u062a\u0627\u0631\u064a\u062e",
    }
    ar_words = [w for w in re.findall(r"[\u0600-\u06FF]+", chunk)
                if w not in stop and len(w) > 1]
    seen: set = set()
    unique = []
    for w in ar_words:
        if w not in seen:
            seen.add(w)
            unique.append(w)
    return " ".join(unique).strip()


# ---------------------------------------------------------------------------
# Item parsing
# ---------------------------------------------------------------------------

def get_nums(segment: str):
    return [
        n for n in re.findall(r"[\d,]+\.?\d*", segment)
        if clean_number(n) not in (0, None)
        and len(re.sub(r"[,.]", "", n)) <= 8
    ]


def parse_item_line(line: str):
    """
    Extract SKU, Description, Quantity, Unit price from one invoice line.

    Strategy: find the (qty, price, total) triple where qty*price ≈ total.
    Numbers may appear before or after the English product code, so we search
    ALL numbers on the line. Original string form is used to detect whether a
    number was written with a decimal point (e.g. "18.00" = price-like vs
    "200" = qty-like).
    """
    # Exclude pack-count bracket contents like (510) and unit-adjacent numbers
    pack_bracket: set = set()
    for m in re.finditer(r"\(\s*(\d+)\s*\)", line):
        pack_bracket.add(m.group(1))

    unit_adjacent: set = set()
    # Digit directly followed by Arabic unit (no space required — Arabic is non-ASCII so no \b issues)
    for m in re.finditer(r"(\d+)\s*(?:كجم|كغ|قطع|قطعة)", line):
        unit_adjacent.add(m.group(1))
    # Standalone ك preceded by a digit and followed by space/end (e.g. "18 ك")
    for m in re.finditer(r"(\d+)\s+ك(?:\s|$)", line):
        unit_adjacent.add(m.group(1))

    all_raw = get_nums(line)
    cand_pairs = [
        (n, clean_number(n))
        for n in all_raw
        if n not in pack_bracket and n not in unit_adjacent and clean_number(n)
    ]

    if len(cand_pairs) < 2:
        return None

    # Numbers whose original string has an explicit decimal point → price-like
    def has_decimal(n_str: str) -> bool:
        return "." in n_str

    cand_vals = [v for _, v in cand_pairs]

    row_total = None
    qty = None
    unit_price = None

    # Step 1: find triple (t, q, p) where q * p ≈ t (within 5%)
    unique_vals = sorted(set(cand_vals), reverse=True)
    for total_cand in unique_vals:
        if total_cand < 200:
            continue
        rest_pairs = [(s, v) for s, v in cand_pairs if v != total_cand or cand_vals.count(v) > 1]
        # Remove one occurrence of total_cand
        tc_removed = False
        rest_filtered = []
        for s, v in cand_pairs:
            if v == total_cand and not tc_removed:
                tc_removed = True
                continue
            rest_filtered.append((s, v))

        best_err = 0.05
        for i, (s1, v1) in enumerate(rest_filtered):
            for s2, v2 in rest_filtered[i:]:
                if v1 == 0 or v2 == 0:
                    continue
                err = abs(v1 * v2 - total_cand) / total_cand
                if err < best_err:
                    best_err = err
                    row_total = total_cand
                    # Written-decimal → price-like, no decimal → qty-like
                    if has_decimal(s1) and not has_decimal(s2):
                        qty, unit_price = v2, v1
                    elif has_decimal(s2) and not has_decimal(s1):
                        qty, unit_price = v1, v2
                    elif not has_decimal(s1) and not has_decimal(s2):
                        qty, unit_price = min(v1, v2), max(v1, v2)
                    else:
                        qty, unit_price = v1, v2
        if row_total is not None:
            break

    # Step 2: fallback — no matching triple; prefer explicit-decimal as price
    if qty is None or unit_price is None:
        qty_like = [(s, v) for s, v in cand_pairs if not has_decimal(s)]
        price_like = [(s, v) for s, v in cand_pairs if has_decimal(s)]
        if qty_like and price_like:
            qty = min(v for _, v in qty_like)
            unit_price = min(v for _, v in price_like)
        elif len(cand_pairs) >= 2:
            vals = sorted(v for _, v in cand_pairs)
            qty = vals[0]
            unit_price = vals[-1]

    # Step 3: infer qty from total/price if still missing
    if row_total and unit_price and (qty is None or qty == unit_price):
        computed = round(row_total / unit_price)
        if computed > 0 and abs(computed * unit_price - row_total) / row_total < 0.05:
            qty = computed

    # Step 4: if price looks like a weight/count (< 50 SAR) and qty is far larger,
    # they are likely swapped (e.g. 12.2 kg at 2850 SAR/kg)
    if qty is not None and unit_price is not None:
        if unit_price < 50 and qty > unit_price * 30:
            qty, unit_price = unit_price, qty

    # Description: collect English tokens including alphanumeric codes (e.g. BONEINCUT6WAY)
    # Pattern: sequence starting and ending with a letter, may contain digits in middle
    all_eng = re.findall(r"[A-Za-z][A-Za-z\d]*[A-Za-z]|[A-Za-z]{2,}", line)
    desc_words = [w for w in all_eng if len(w) >= 3 or w.isupper()]
    seen_w: set = set()
    deduped = []
    for w in desc_words:
        key = w.upper()
        if key not in seen_w:
            seen_w.add(key)
            deduped.append(w)
    desc = " ".join(deduped).strip()

    sku = extract_sku_from_line(line)

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


# ---------------------------------------------------------------------------
# Item extraction strategies
# ---------------------------------------------------------------------------

def _is_summary_with_number(line: str) -> bool:
    """True only when a summary keyword AND a number appear on the same line.
    A lone keyword without a number is just a column header, not the summary."""
    has_kw = any(kw in line for kw in SUMMARY_KW)
    has_num = bool(re.search(r"\d{3,}", line))
    return has_kw and has_num


def extract_items_multiline(text: str):
    """
    Extract items from native PDFs where each product spans multiple lines.
    Finds the table section, removes pipe separators, groups lines by product,
    then parses each group as a combined item block.
    """
    lines = text.split("\n")
    in_table = False
    table_lines = []
    for line in lines:
        if any(h in line for h in HEADER_KW):
            in_table = True
            continue
        # Only treat as summary if keyword co-occurs with a number on the same line
        if in_table and _is_summary_with_number(line):
            break
        if in_table:
            cleaned = line.strip().replace("|", " ").replace("—", "").strip()
            if cleaned:
                table_lines.append(cleaned)

    if not table_lines:
        return []

    # Split table lines into per-item blocks: a new block starts when we find
    # a line containing an English product word (>=4 chars uppercase)
    blocks = []
    current: list = []
    for line in table_lines:
        if re.search(r"[A-Z]{3,}", line) and current:
            blocks.append(" ".join(current))
            current = [line]
        else:
            current.append(line)
    if current:
        blocks.append(" ".join(current))

    items = []
    for block in blocks:
        if not block.strip():
            continue
        parsed = parse_item_line(block)
        if parsed and (parsed.get("Description") or parsed.get("SKU")):
            items.append(parsed)
    return items


def extract_items_positional(word_df: pd.DataFrame, text: str):
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
            if in_table and any(kw in line for kw in SUMMARY_KW):
                break
            if not in_table:
                continue
            parsed = parse_item_line(line)
            if parsed:
                items.append(parsed)

    # Fallback: any line with Arabic + English + 2+ numbers
    if not items:
        for line in text.split("\n"):
            line = line.strip()
            if not line or any(kw in line for kw in SUMMARY_KW + HEADER_KW):
                continue
            has_arabic = bool(re.search(r"[\u0600-\u06FF]{2,}", line))
            has_english = bool(re.search(r"[A-Za-z]{2,}", line))
            has_nums = len(re.findall(r"[\d,]+\.?\d*", line)) >= 2
            if not (has_english and has_nums):
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
        with pdfplumber.open(str(pdf_path)) as pdf:
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
                        raw_sku = reshape(vals[5]) if len(vals) > 5 else ""
                        raw_desc = reshape(vals[4]) if len(vals) > 4 else ""
                        items.append({
                            "Unit price": clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity": clean_number(vals[3]) if len(vals) > 3 else None,
                            "Description": raw_desc,
                            "SKU": clean_sku(raw_sku),
                        })
    except Exception:
        pass
    return items


# ---------------------------------------------------------------------------
# Financial extraction
# ---------------------------------------------------------------------------

def extract_financials(text: str):
    """Find total-before-tax, VAT, and total-after-tax using ratio 1.15."""
    pos = max(text.find("\u0627\u0644\u0645\u062c\u0645\u0648\u0639"), 0)  # المجموع
    # Also look for OCR-garbled variant
    pos2 = max(text.find("\u0627\u0644\u0645\u062d\u0645\u0648\u062c"), 0)  # المحموج
    fin_start = max(pos, pos2) if min(pos, pos2) == 0 else min(pos, pos2)
    fin = text[fin_start:] if fin_start else text

    nums_raw = []
    for n in re.findall(r"[\d,]+\.?\d*", fin):
        v = clean_number(n)
        if v and v > 100 and v not in (15, 150):
            nums_raw.append(v)
    unique = sorted(set(nums_raw))

    tb = ta = vat = None

    # Strategy 1: pair with ratio ~1.15 (VAT relationship)
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

    # Strategy 2: (vat, total) where vat/total ~0.13
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

    # Strategy 3: largest number = total-after-tax
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


# ---------------------------------------------------------------------------
# Metadata extraction
# ---------------------------------------------------------------------------

def extract_invoice_number(text: str) -> str:
    """Extract invoice number using multiple strategies."""
    for pattern in [
        r"(?:\u0631\u0642\u0645|\u0631\u0647\u0645).{0,15}(?:\u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629|\u0627\u0644\u0645\u0627\u0631\u0648\u0631\u0629|\u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629).{0,5}(\d{4,6})",
        r"(?:\u0627\u0644\u0641\u0627\u062a\u0648\u0631\u0629|\u0627\u0644\u063a\u0627\u062a\u0648\u0631\u0629|\u0627\u0644\u0645\u0627\u0631\u0648\u0631\u0629).{0,5}(\d{4,6})",
        r"\b(0\d{4,5})\b",  # starts with 0 + 4-5 digits
    ]:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return ""


def extract_address(text: str) -> str:
    """Extract address after العنوان keyword."""
    m = re.search(
        r"(?:\u0627\u0644\u0639\u0646\u0648\u0627\u0646|\u0627\u0644\u0639\u0648\u0627\u0646)"  # العنوان / العوان (OCR)
        r"[\s:]*(.+?)(?=\n\d{7,}|\n(?:\u0627\u0644\u0645\u062c\u0645\u0648\u0639|\u0645\u062f\u0641\u0648\u0639|\u0631\u0642\u0645 \u0627\u0644\u062d\u0633\u0627\u0628)|\Z)",
        text, re.DOTALL,
    )
    if not m:
        return ""
    address = " ".join(m.group(1).split()).strip()
    return re.sub(r"\s*\d{10}\s*$", "", address).strip()


def extract_metadata(pdf_path, text: str) -> dict:
    inv_num = extract_invoice_number(text)

    date_m = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", text)
    inv_date = date_m.group(1) if date_m else ""

    address = extract_address(text)

    total_before, vat, total_after = extract_financials(text)

    return {
        "Invoice Number": inv_num,
        "Invoice Date": inv_date,
        "Address": address,
        "Balance": total_after,
        "Paid": 0,
        "Total before tax": total_before,
        "VAT 15%": vat,
        "Total after tax": total_after,
        "Source File": Path(pdf_path).name,
    }


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process_pdf(pdf_path):
    pdf_path = Path(pdf_path)
    text, mode = get_text(pdf_path)
    meta = extract_metadata(pdf_path, text)

    if mode == "ocr":
        word_df = get_ocr_words(pdf_path)
        items = extract_items_positional(word_df, text)
    else:
        word_df = pd.DataFrame()
        items = extract_items_native(pdf_path)
        if not items:
            items = extract_items_multiline(text)
        if not items:
            items = extract_items_positional(pd.DataFrame(), text)

    # Customer name: filename is the most reliable source
    cname = extract_name_from_filename(pdf_path)
    if not cname or len(cname) < 4:
        cname = extract_customer_name_text(text)
    if not cname or len(cname) < 4:
        cname = ""
    meta["Customer Name"] = cname

    # Deduplicate items
    seen: set = set()
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


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("\U0001f4c4 Invoice Extractor — PDF to Excel")

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
                pdf_paths.extend(tmp.glob("**/*.pdf"))
            else:
                pdf_paths.append(fp)

        all_data = []
        for i, path in enumerate(pdf_paths):
            st.write(f"\U0001f4c4 **{path.name}**")
            with st.spinner("Extracting..."):
                try:
                    df, mode, raw_text = process_pdf(path)
                except Exception as exc:
                    st.error(f"Failed: {exc}")
                    continue
            st.caption(f"Mode: `{mode}` — {len(df)} row(s)")

            if debug_mode:
                with st.expander(f"\U0001f4cb Full raw text — {path.name}", expanded=True):
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
