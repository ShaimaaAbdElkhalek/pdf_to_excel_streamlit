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


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

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


def extract_name_from_filename(pdf_path):
    stem = Path(pdf_path).stem
    name = re.sub(r"[-_\s]*\d+[-_\s]*$", "", stem).strip()
    name = re.sub(r"^[-_\s]*\d+[-_\s]*", "", name).strip()
    if re.search(r"[\u0600-\u06FF]", name):
        return name
    return ""


# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────

UNIT_WORDS = {
    "كرتونة", "كرتون", "قطعة", "علبة", "كيس", "طن", "كجم",
    "لتر", "كغ", "جرام", "مل", "حبة", "رول", "باكيت", "صندوق",
}

HEADER_KW = ["البند", "الوصف", "العدد", "سعر الوحدة", "الكمية", "الوحدة"]

SUMMARY_KW = [
    "المجموع", "مدفوع", "الرصيد", "القيمة", "القيمه",
    "الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي",
    "رقم الحساب", "الايبان", "IBAN", "SA08", "Kingdome", "المملكة",
    "رقم الفاتورة", "تاريخ", "اسم العميل", "الرقم الضريبي",
    "رقم السجل", "العنوان", "الجوال", "السجل التجاري",
]

FINAL_COLS = [
    "Invoice Number", "Invoice Date", "Customer Name",
    "Address", "Balance", "Paid",
    "Total before tax", "VAT 15%", "Total after tax",
    "Unit price", "Quantity", "Description", "SKU",
    "Source File",
]

# Pack-size pattern: "18 ك", "(510)", "18 كجم", "18 كغ", "12×200"
PACK_SIZE_PAT = re.compile(
    r"\b\d+\s*ك\b"
    r"|\(\s*\d+\s*\)"
    r"|\b\d+\s*كجم\b"
    r"|\b\d+\s*كغ\b"
    r"|\b\d+\s*×\s*\d+\b",
    re.UNICODE,
)

# Known Arabic SKU → full English description mapping.
# Used as fallback when OCR misreads English words as Arabic (e.g. VEAL → ادعلا).
# Key = Arabic product name (after clean_sku), Value = correct full English desc.
SKU_TO_DESC = {
    "فيل ليج هندي صاحبة": "VEAL LEG SAHIBA",
    "فيل ليج هندي":        "VEAL LEG HINDI",
    "فيل ليج":             "VEAL LEG",
}


# ─────────────────────────────────────────────
# SKU cleaner
# ─────────────────────────────────────────────

def clean_sku(raw_sku):
    """Remove pack-size tokens and stray digits; keep meaningful Arabic words."""
    cleaned = PACK_SIZE_PAT.sub(" ", raw_sku)
    cleaned = re.sub(r"\b\d+\b", " ", cleaned)
    words = [w for w in cleaned.split() if w not in UNIT_WORDS and len(w) > 1]
    return " ".join(words).strip()


# ─────────────────────────────────────────────
# PDF text extraction
# ─────────────────────────────────────────────

def get_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = "\n".join(page.get_text() for page in doc).strip()
    if len(text) > 50:
        return text, "native"
    # Fallback 1: pdf2image + pytesseract
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
    # Fallback 2: PyMuPDF pixmap + pytesseract
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


# ─────────────────────────────────────────────
# Row reconstruction from word-level OCR
# ─────────────────────────────────────────────

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


# ─────────────────────────────────────────────
# Customer name extraction
# ─────────────────────────────────────────────

def extract_customer_name_text(text):
    m = re.search(
        r"اسم العميل\s*[:\s]+(.+?)(?=الرقم الضريبي|رقم السجل|العنوان)",
        text, re.DOTALL,
    )
    if not m:
        return ""
    chunk = m.group(1)
    chunk = re.sub(r"(?:قم الغاتورة|رقم الفاتورة)\s*\d+", "", chunk)
    chunk = re.sub(r"\d{4,}", "", chunk)
    chunk = re.sub(r"[a-zA-Z]{2,}", "", chunk)
    stop = {
        "اسم", "العميل", "فاتورة", "إلى", "إلىة", "التتجار",
        "رقم", "الفاتورة", "تاريخ", "الغاتورة", "إلىة",
    }
    arabic_words = [
        w for w in re.findall(r"[\u0600-\u06FF]+", chunk)
        if w not in stop and len(w) > 1
    ]
    seen = set()
    unique = []
    for w in arabic_words:
        if w not in seen:
            seen.add(w)
            unique.append(w)
    return " ".join(unique).strip()


# ─────────────────────────────────────────────
# Number helpers
# ─────────────────────────────────────────────

def get_nums(segment):
    """Return valid candidate numbers from a text segment."""
    return [
        n for n in re.findall(r"[\d,]+\.?\d*", segment)
        if clean_number(n) not in (0, None)
        and len(re.sub(r"[,.]", "", n)) <= 8
    ]


# ─────────────────────────────────────────────
# Line parser  (core extraction logic)
# ─────────────────────────────────────────────

def parse_item_line(line):
    """
    Extract SKU / Description / Quantity / Unit price from one product line.

    Handles two real OCR layouts found in these invoices:

    Layout A (2445 / 2975):
        Arabic-SKU [pack-spec] | English-Desc | qty | price | total
        Numbers are AFTER the last English word.
        e.g. "فيل ليج هندي صاحبة 18 ك (510) VEAL LEG SAHIBA |200 |350.00 0 كرتونة 70,000"

    Layout B (2832):
        Arabic-SKU | [noise] qty price | [noise] | English-Desc | total
        Numbers are BETWEEN English words.
        e.g. "كبدة ضأن ... | pS 2850 12.20 | 190 | LAMBLIVER(JUNNE GOLD) 34,770"

    Routing logic (tested against real OCR output):
    - If after_nums (after last English) has >= 2 items  → Layout A, use after_nums only
    - If after_nums has < 2 but middle_nums exist        → Layout B, use middle + after
    - Fallback: use before + after

    Numbers in "before" zone are ignored when above strategies work, because
    before-English numbers in these invoices are always pack specs (18, part of SKU).

    Pack-size numbers like (510) are stripped from candidates via regex.

    Description fallback via SKU_TO_DESC:
    When OCR misreads English words as Arabic (e.g. VEAL → ادعلا on invoice 2445),
    the extracted desc is partial (e.g. "LEG SAHIBA"). We check if all desc words
    are a subset of a known full description mapped from the Arabic SKU, and if so
    replace with the full description.
    """

    eng_matches = list(re.finditer(r"[A-Za-z]{2,}", line))

    if eng_matches:
        first_start = eng_matches[0].start()
        last_end    = eng_matches[-1].end()
        before_nums = get_nums(line[:first_start])
        middle_nums = get_nums(line[first_start:last_end])
        after_nums  = get_nums(line[last_end:])

        if len(after_nums) >= 2:
            # Layout A: all data numbers live after the last English token
            all_nums = after_nums
        elif len(middle_nums) >= 1:
            # Layout B: data numbers are between / after English tokens
            all_nums = middle_nums + after_nums
        else:
            # Fallback
            all_nums = before_nums + after_nums
    else:
        all_nums = get_nums(line)

    if len(all_nums) < 2:
        return None

    # Last number = row total; the rest are candidates (qty, price, pack-counts)
    candidates = all_nums[:-1]

    # Strip pack-count numbers like (510) from candidates
    pack_bracket_nums = set()
    for m in re.finditer(r"\(\s*(\d+)\s*\)", line):
        pack_bracket_nums.add(m.group(1))
    candidates = [n for n in candidates if n not in pack_bracket_nums]

    if not candidates:
        return None

    # Unit price = first candidate that has a decimal point
    unit_price = None
    for n in candidates:
        if "." in str(n):
            unit_price = clean_number(n)
            break

    # Quantity = smallest whole number among candidates
    whole_nums = [
        clean_number(n) for n in candidates
        if "." not in str(n) and clean_number(n) and clean_number(n) != unit_price
    ]
    qty = min(whole_nums) if whole_nums else None

    # Fallback: no decimal found
    if unit_price is None:
        vals = sorted([clean_number(n) for n in candidates if clean_number(n)])
        if len(vals) >= 2:
            qty = vals[0];  unit_price = vals[-1]
        elif vals:
            unit_price = vals[0]

    # ── Description: scan the FULL line for English words ──────────────────
    all_eng = re.findall(r"[A-Za-z]{2,}", line)
    desc_words = [w for w in all_eng if len(w) >= 3 or w.isupper()]
    seen_w: set = set()
    deduped = []
    for w in desc_words:
        key = w.upper()
        if key not in seen_w:
            seen_w.add(key)
            deduped.append(w)
    desc = " ".join(deduped).strip()

    # ── SKU: Arabic block stripped of pack-size tokens ─────────────────────
    sku_m = re.search(r"([\u0600-\u06FF][\u0600-\u06FF\s\d\(\)ك]+)", line)
    sku = clean_sku(sku_m.group(1).strip()) if sku_m else ""
    if not sku:
        ar_words = re.findall(r"[\u0600-\u06FF]{2,}", line)
        sku = clean_sku(" ".join(w for w in ar_words if w not in UNIT_WORDS))

    # ── Description fallback via SKU lookup ────────────────────────────────
    # Handles OCR misreads: e.g. VEAL OCR'd as Arabic "ادعلا" on invoice 2445
    # → desc = "LEG SAHIBA" (partial). We detect the partial match and replace.
    if sku and desc:
        for ar_key, en_val in SKU_TO_DESC.items():
            if ar_key in sku:
                desc_upper  = desc.upper()
                en_val_upper = en_val.upper()
                # If all current desc words are present in the full expected desc
                # but the desc is shorter → OCR dropped some words → use full
                desc_words_u = desc_upper.split()
                all_present  = all(w in en_val_upper for w in desc_words_u)
                if all_present and desc_upper != en_val_upper:
                    desc = en_val
                break
    elif sku and not desc:
        for ar_key, en_val in SKU_TO_DESC.items():
            if ar_key in sku:
                desc = en_val
                break

    if not (sku or desc):
        return None

    return {
        "SKU":         sku,
        "Description": desc,
        "Quantity":    qty,
        "Unit price":  unit_price,
    }


# ─────────────────────────────────────────────
# Item extraction — OCR path (positional)
# ─────────────────────────────────────────────

def extract_items_positional(word_df, text):
    items = []

    # Pass 1: word-level positional rows
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

    # Pass 2: text-line fallback (header detected)
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
            if in_table and any(
                kw in line for kw in ["المجموع", "القيمة", "الإجمالي", "الإحمالي", "مدفوع"]
            ):
                break
            if not in_table:
                continue
            parsed = parse_item_line(line)
            if parsed:
                items.append(parsed)

    # Pass 3: headerless fallback
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


# ─────────────────────────────────────────────
# Item extraction — native PDF path (pdfplumber)
# ─────────────────────────────────────────────

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
                        sku_words = [
                            w for w in re.findall(r"[\u0600-\u06FF\d\(\)ك]+", raw_sku)
                            if w not in UNIT_WORDS
                        ]
                        items.append({
                            "Unit price":  clean_number(vals[2]) if len(vals) > 2 else None,
                            "Quantity":    clean_number(vals[3]) if len(vals) > 3 else None,
                            "Description": raw_desc,
                            "SKU":         clean_sku(" ".join(sku_words)),
                        })
    except Exception:
        pass
    return items


# ─────────────────────────────────────────────
# Metadata extraction
# ─────────────────────────────────────────────

def extract_metadata(pdf_path, text):
    def get_amount_clean(*keywords):
        for kw in keywords:
            m = re.search(rf"{re.escape(kw)}[^\d\n]*([\d,٬٫]+\.?\d*)", text)
            if m:
                return clean_number(m.group(1))
        return None

    # Invoice number
    inv_num = ""
    for p in [
        r"رقم الفاتورة\s*[:\s]*(\d{4,6})",
        r"(?:قم الغاتورة|الغاتورة)\s*(\d{4,6})",
        r"\b(0\d{4,5})\b",
    ]:
        m = re.search(p, text)
        if m:
            inv_num = m.group(1).strip()
            break

    # Invoice date
    date_m   = re.search(r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})", text)
    inv_date = date_m.group(1) if date_m else ""

    # Address
    address = ""
    m = re.search(
        r"العنوان\s*[:\s]+(.+?)(?=\n\d{7,}|\n(?:المجموع|مدفوع|رقم الحساب)|\Z)",
        text, re.DOTALL,
    )
    if m:
        address = " ".join(m.group(1).split()).strip()
        address = re.sub(r"\s*\d{10}\s*$", "", address).strip()

    # Financial fields
    vat          = get_amount_clean("القيمة المضافة", "القيمه المضافة")
    total_before = get_amount_clean("المجموع")

    if total_before and vat and vat > 0:
        ratio = total_before / vat
        if ratio > 10:
            s = str(int(total_before))
            if len(s) > 4:
                candidate = clean_number(s[1:])
                if candidate and abs(candidate / vat - 100 / 15) < 1:
                    total_before = candidate

    # VAT sanity check: OCR sometimes prepends a stray digit to the VAT number
    # e.g. real VAT=5,175 but OCR reads 35,175 (﷼ symbol digit bleeds in)
    # Fix: if VAT > 20% of total_before (impossible for 15%), strip leading digit
    if vat and total_before and total_before > 0:
        if vat > total_before * 0.20:
            expected_vat = total_before * 0.15
            s = str(int(vat))
            if len(s) > 3:
                candidate = clean_number(s[1:])
                if candidate and abs(candidate - expected_vat) / expected_vat < 0.05:
                    vat = candidate
            # Also try dividing by 10
            if vat > total_before * 0.20:
                candidate2 = vat / 10
                if abs(candidate2 - expected_vat) / expected_vat < 0.05:
                    vat = candidate2

    total_after = get_amount_clean("الإجمالي", "الإحمالي", "اإلجمالي", "الاجمالي")
    if not total_after and total_before and vat:
        total_after = round(total_before + vat, 2)

    # Paid — try direct regex first
    paid_m = re.search(r"مدفوع[^\d\n]*([\d,٬٫]+\.?\d*)", text)
    paid   = clean_number(paid_m.group(1)) if paid_m else None

    # Paid validation: OCR garbage like "مدفوع 5 - يلد" gives paid=5 which is wrong.
    # Discard if paid is unreasonably small compared to total (< 1% of total_after).
    if paid and total_after and paid < total_after * 0.01:
        paid = None

    # Paid fallback: OCR often garbles مدفوع line (e.g. "مدفوع 5 - يلد", "E9920 80,500-")
    # If الرصيد المستحق = 0.00, the invoice is fully paid → Paid = Total after tax
    if not paid:
        balance_m = re.search(r"الرصيد[^\d\n]*([\d,٬٫]+\.?\d*)", text)
        balance   = clean_number(balance_m.group(1)) if balance_m else None
        if balance is not None and balance == 0 and total_after:
            paid = total_after

    return {
        "Invoice Number":   inv_num,
        "Invoice Date":     inv_date,
        "Address":          address,
        "Balance":          total_after,
        "Paid":             paid,
        "Total before tax": total_before,
        "VAT 15%":          vat,
        "Total after tax":  total_after,
        "Source File":      pdf_path.name,
    }


# ─────────────────────────────────────────────
# Main per-PDF processor
# ─────────────────────────────────────────────

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

    # Deduplicate items
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


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

st.set_page_config(page_title="Invoice Extractor", layout="wide")
st.title("📄 Invoice Extractor — PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload PDF or ZIP files", type=["pdf", "zip"], accept_multiple_files=True
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
