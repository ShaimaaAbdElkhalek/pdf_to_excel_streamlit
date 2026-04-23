import re
import pandas as pd

def extract_table(text):

    # normalize
    text = text.replace("‏", "")
    text = re.sub(r"\s+", " ", text)

    rows = []

    # STEP 1: find all number clusters (this is your REAL table anchor)
    pattern = re.finditer(r"(\d{1,3}(?:,\d{3})*\.?\d*)", text)

    numbers = [(m.group(), m.start()) for m in pattern]

    # STEP 2: group numbers in triplets (qty / price / extra noise)
    for i in range(len(numbers) - 2):

        n1, p1 = numbers[i]
        n2, p2 = numbers[i + 1]
        n3, p3 = numbers[i + 2]

        # skip huge numbers (totals, IBAN, etc.)
        if len(n1) > 6 or len(n2) > 6 or len(n3) > 6:
            continue

        # STEP 3: extract surrounding text window
        start = max(0, p1 - 120)
        end = min(len(text), p3 + 120)

        window = text[start:end]

        # must contain product text
        if not any(c.isalpha() for c in window):
            continue

        # STEP 4: clean description
        desc = re.sub(r"\d{1,3}(?:,\d{3})*\.?\d*", "", window)
        desc = re.sub(r"[^\w\s\u0600-\u06FF]", " ", desc)
        desc = re.sub(r"\s+", " ", desc).strip()

        if len(desc) < 5:
            continue

        rows.append({
            "Description": desc,
            "Quantity": n2,
            "Unit Price": n3
        })

    return pd.DataFrame(rows)
