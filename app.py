import re
import pandas as pd

def extract_items(text):

    # normalize
    text = re.sub(r"‏", "", text)
    text = re.sub(r"\s+", " ", text)

    rows = []

    # STEP 1: find ALL numbers with positions
    numbers = [(m.group(), m.start()) for m in re.finditer(r"\d+\.\d+|\d+", text)]

    # STEP 2: slide through numbers in pairs (qty + price patterns)
    for i in range(len(numbers) - 1):

        num1, pos1 = numbers[i]
        num2, pos2 = numbers[i + 1]

        # ignore big totals like 76,034 or IBAN parts
        if len(num1) > 5 or len(num2) > 5:
            continue

        # STEP 3: extract surrounding context window
        start = max(0, pos1 - 120)
        end = min(len(text), pos2 + 120)

        window = text[start:end]

        # must contain letters (product)
        if not any(c.isalpha() for c in window):
            continue

        # remove numbers to get product name
        desc = re.sub(r"\d+\.\d+|\d+", "", window)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 6:
            continue

        # filter invoice noise
        if "المجموع" in desc or "الإحمالي" in desc or "الرصيد" in desc:
            continue

        rows.append({
            "SKU / Description": desc,
            "Quantity": num1,
            "Unit Price": num2
        })

    # remove duplicates (very important)
    df = pd.DataFrame(rows).drop_duplicates()

    return df
