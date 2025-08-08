# second version of your code using pdfplumber instead of tabula

import streamlit as st
import os
import shutil
import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
import re
import tempfile
import zipfile
from pathlib import Path

# =========================
# Helper Functions
# =========================

def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

def find_field(text, keyword):
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""

def process_pdf(pdf_path, safe_folder):
    all_rows = []
    try:
        with fitz.open(pdf_path) as doc:
            full_text = "\n".join([page.get_text() for page in doc])

        invoice_number = find_field(full_text, "رقم الفاتورة")
        invoice_date = find_field(full_text, "تاريخ الفاتورة")
        customer_name = find_field(full_text, "فاتورة ضريبية")
        address_part2 = find_field(full_text, "العنوان")
        address_part1 = find_field(full_text, "رقم السجل")
        address = f"{address_part1} {address_part2}" if address_part1 or address_part2 else ""
        paid_value = find_field(full_text, "مدفوع")
        balance_value = find_field(full_text, "الرصيد المستحق")

        ascii_name = f"bill_{pdf_path.stem.encode('ascii', errors='ignore').decode()}.pdf"
        safe_pdf_path = safe_folder / ascii_name
        shutil.copy(pdf_path, safe_pdf_path)

        with pdfplumber.open(safe_pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    merged_rows = []
                    temp_row = []

                    for row in table:
                        row_values = [cell if cell else "" for cell in row]

                        if is_data_row(row_values):
                            if temp_row:
                                combined = [temp_row[0] + " " + row_values[0]] + row_values[1:]
                                merged_rows.append(combined)
                                temp_row = []
                            else:
                                merged_rows.append(row_values)
                        else:
                            temp_row = row_values

                    if merged_rows:
                        headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند"]
                        df_merged = pd.DataFrame(merged_rows, columns=headers[:len(merged_rows[0])])
                        df_merged["Invoice Number"] = invoice_number
                        df_merged["Invoice Date"] = invoice_date
                        df_merged["Customer Name"] = customer_name
                        df_merged["Address"] = address
                        df_merged["Paid"] = paid_value
                        df_merged["Balance"] = balance_value
                        df_merged["Source File"] = pdf_path.name
                        all_rows.append(df_merged)
    except Exception as e:
        st.error(f"❌ Error in {pdf_path.name}: {e}")
    return all_rows
