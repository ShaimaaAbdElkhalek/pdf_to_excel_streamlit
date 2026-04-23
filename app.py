import re
import pandas as pd

# =========================
# CLEAN TEXT
# =========================

def clean_text(text):
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# SMART PRODUCT EXTRACTION (REAL FIX)
# =========================

def extract_items(text):

    text = clean_text(text)

    rows = []

    # 🔥 STEP 1: isolate product area
    # (between "العدد" and "المجموع")
    try:
        product_block = re.search(r"العدد(.*?)المجموع", text).group(1)
    except:
        product_block = text

    # 🔥 STEP 2: split using product name patterns
    # products usually contain English words or capital letters
    candidates = re.split(r"(?=[A-Z][A-Z\s]+)", product_block)

    for c in candidates:

        c = c.strip()
        if len(c) < 10:
            continue

        # extract all numbers
        nums = re.findall(r"\d+\.\d+|\d+", c)

        if len(nums) < 2:
            continue

        # extract description (remove numbers)
        desc = re.sub(r"\d+\.\d+|\d+", "", c)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        quantity = nums[-2]
        price = nums[-1]

        rows.append({
            "SKU / Description": desc,
            "Quantity": quantity,
            "Unit Price": price
        })

    return pd.DataFrame(rows)

# =========================
# TEST WITH YOUR TEXT
# =========================

text = """
شركة بداية ونهاية الجودة التجارية ...
العدد سعر الوحدة الكمية المجموع
58,584.42 كجم 69 18.00 200 BONE IN CUT 6 WAY عجل مقطع افيكو نيوزلاندي
17,450 كجم 5 20.00 49 WHOLE LEG RUSTAM
المجموع 76,034.42
"""

df = extract_items(text)
print(df)
