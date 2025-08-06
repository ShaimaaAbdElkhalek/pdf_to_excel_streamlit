import streamlit as st
import os
import tabula
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path
import re
from io import BytesIO

# =========================
# Helper Functions
# =========================
def is_data_row(row):
    """Check if a row contains numeric values."""
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)

def find_field(text, keyword):
    """Extracts value following a given keyword from the text."""
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip()
    return ""

def process_pdfs(uploaded_files):
    all_rows = []

    for uploaded_file in uploaded_files:
        try:
            st.write(f"📄 Processing: {uploaded_file.name}")

            # Save uploaded file temporarily
            temp_pdf_path = f"temp_{uploaded_file.name}"
            with open(temp_pdf_path, "wb") as f:
                f.write(uploaded_file.read())

            # Extract full text
            with fitz.open(temp_pdf_path) as doc:
                full_text = "\n".join([page.get_text() for page in doc])

            # Extract non-table fields
            invoice_number = find_field(full_text, "رقم الفاتورة")
            invoice_date = find_field(full_text, "تاريخ الفاتورة")
            customer_name = find_field(full_text, "فاتورة ضريبية")
            address_part2 = find_field(full_text, "العنوان")
            address_part1 = find_field(full_text, "رقم السجل")
            address = f"{address_part1} {address_part2}" if address_part1 or address_part2 else ""
            paid_value = find_field(full_text, "مدفوع")
            balance_value = find_field(full_text, "الرصيد المستحق")

            # Extract tables
            tables = tabula.read_pdf(temp_pdf_path, pages='all', multiple_tables=True, stream=True)

            for table in tables:
                if not table.empty:
                    merged_rows = []
                    temp_row = []

                    for _, row in table.iterrows():
                        row_values = row.fillna("").astype(str).tolist()

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
                        column_headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند"]
                        df_merged = pd.DataFrame(merged_rows, columns=column_headers[:len(merged_rows[0])])

                        # Add extracted fields
                        df_merged["Invoice Number"] = invoice_number
                        df_merged["Invoice Date"] = invoice_date
                        df_merged["Customer Name"] = customer_name
                        df_merged["Address"] = address
                        df_merged["Paid"] = paid_value
                        df_merged["Balance"] = balance_value
                        df_merged["Source File"] = uploaded_file.name

                        all_rows.append(df_merged)

            os.remove(temp_pdf_path)

        except Exception as e:
            st.error(f"❌ Error in {uploaded_file.name}: {e}")

    if not all_rows:
        return pd.DataFrame()

    # Combine data
    final_df = pd.concat(all_rows, ignore_index=True)

    # Clean Customer Name
    final_df["Customer Name"] = (
        final_df["Customer Name"].astype(str)
        .str.replace(r"اسم العميل\s*[:：]?\s*", "", regex=True)
        .str.strip(" :：﹕")
    )

    # Clean Address
    final_df["Address"] = (
        final_df["Address"].astype(str)
        .str.replace(r"العنوان\s*[:：]?\s*", "", regex=True)
        .str.strip(" :：﹕")
    )

    # Clean Paid and Balance
    for col in ["Paid", "Balance"]:
        final_df[col] = (
            final_df[col].astype(str)
            .str.replace(r"[^\d.,]", "", regex=True)
            .str.replace(",", "", regex=False)
            .astype(float)
        )

    # Ensure numeric
    final_df["العدد"] = pd.to_numeric(final_df["العدد"], errors="coerce")
    final_df["المجموع"] = (
        final_df["المجموع"].astype(str)
        .str.replace(r"[^\d.,]", "", regex=True)
        .str.replace(",", "", regex=False)
        .astype(float)
    )

    # VAT
    final_df["VAT 15% Calc"] = (final_df["المجموع"] * 0.15).round(2)

    # Rename columns
    final_df = final_df.rename(columns={
        "المجموع": "Total before tax",
        "سعر الوحدة": "Unit price",
        "العدد": "Quantity",
        "الوصف": "Description",
        "البند": "SKU"
    })

    # Calculate totals
    final_df["VAT 15% Calc"] = (final_df["Total before tax"] * 0.15).round(2)
    final_df["Total after tax"] = (final_df["Total before tax"] + final_df["VAT 15% Calc"]).round(2)

    # Reorder
    final_df = final_df[
        [
            "Invoice Number", "Invoice Date", "Customer Name", "Address", "Paid", "Balance",
            "Total before tax", "VAT 15% Calc", "Total after tax",
            "Unit price", "Quantity", "Description", "SKU",
            "Source File"
        ]
    ]

    return final_df


# =========================
# Streamlit App
# =========================
st.title("📄 PDF Invoice Extractor & Cleaner")

uploaded_files = st.file_uploader(
    "Upload one or more PDF invoices",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    df = process_pdfs(uploaded_files)

    if not df.empty:
        st.success("✅ Data extracted successfully!")
        st.dataframe(df)

        # Download Excel
        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name="Cleaned_Combined_Tables.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid tables found.")
