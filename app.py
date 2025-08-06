import streamlit as st
import fitz  # PyMuPDF
import tabula
import pandas as pd
import re
import tempfile
from pathlib import Path
import zipfile
# =========================
# Helper Functions
# =========================
def is_data_row(row):
    return any(str(cell).replace(",", "").replace("٫", ".").replace("٬", ".").replace(" ", "").isdigit() for cell in row)
def find_field(text, keyword):
    pattern = rf"{keyword}[:\s]*([^\n]*)"
    match = re.search(pattern, text)
    return match.group(1).strip() if match else ""
def process_pdf(pdf_path):
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
        tables = tabula.read_pdf(str(pdf_path), pages='all', multiple_tables=True, stream=True)
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
                    headers = ["المجموع", "الكمية", "سعر الوحدة", "العدد", "الوصف", "البند"]
                    df = pd.DataFrame(merged_rows, columns=headers[:len(merged_rows[0])])
                    df["Invoice Number"] = invoice_number
                    df["Invoice Date"] = invoice_date
                    df["Customer Name"] = customer_name
                    df["Address"] = address
                    df["Paid"] = paid_value
                    df["Balance"] = balance_value
                    df["Source File"] = pdf_path.name
                    all_rows.append(df)
        if all_rows:
            return pd.concat(all_rows, ignore_index=True)
    except Exception as e:
        st.warning(f"❌ Failed to process {pdf_path.name}: {e}")
    return pd.DataFrame()
def clean_df(df):
    df["Customer Name"] = df["Customer Name"].astype(str).str.replace(r"^\s*اسم العميل\s*[:：﹕٭‪]?\s*", "", regex=True).str.strip(" :：﹕")
    df["Address"] = df["Address"].astype(str).str.replace(r"^\s*العنوان\s*[:：﹕٭‪]?\s*", "", regex=True).str.replace(r"رقم الجوال.*", "", regex=True).str.strip(" :：﹕")
    for col in ["Paid", "Balance"]:
        df[col] = df[col].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "").astype(float)
    df["العدد"] = pd.to_numeric(df["العدد"], errors="coerce")
    df["المجموع"] = df["المجموع"].astype(str).str.replace(r"[^\d.,]", "", regex=True).str.replace(",", "").astype(float)
    df["VAT 15% Calc"] = (df["المجموع"] * 0.15).round(2)
    df = df.rename(columns={
        "المجموع": "Total before tax",
        "سعر الوحدة": "Unit price",
        "العدد": "Quantity",
        "الوصف": "Description",
        "البند": "SKU"
    })
    df["VAT 15% Calc"] = (df["Total before tax"] * 0.15).round(2)
    df["Total after tax"] = (df["Total before tax"] + df["VAT 15% Calc"]).round(2)
    return df[[
        "Invoice Number", "Invoice Date", "Customer Name", "Address", "Paid", "Balance",
        "Total before tax", "VAT 15% Calc", "Total after tax",
        "Unit price", "Quantity", "Description", "SKU", "Source File"
    ]]
# =========================
# Streamlit UI
# =========================
st.title("📄 Arabic Invoice PDF to Excel Converter")
st.markdown("Upload one or more PDF invoices (or a ZIP folder of PDFs), and get a cleaned Excel sheet with extracted data.")
uploaded_file = st.file_uploader("Upload PDF files or ZIP folder", type=["pdf", "zip"], accept_multiple_files=True)
if uploaded_file:
    with st.spinner("Processing files..."):
        temp_dir = tempfile.TemporaryDirectory()
        pdf_paths = []
        for file in uploaded_file:
            suffix = Path(file.name).suffix.lower()
            if suffix == ".pdf":
                file_path = Path(temp_dir.name) / file.name
                with open(file_path, "wb") as f:
                    f.write(file.read())
                pdf_paths.append(file_path)
            elif suffix == ".zip":
                zip_path = Path(temp_dir.name) / file.name
                with open(zip_path, "wb") as f:
                    f.write(file.read())
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir.name)
                    pdf_paths += list(Path(temp_dir.name).rglob("*.pdf"))
        all_dataframes = [process_pdf(pdf) for pdf in pdf_paths]
        all_dataframes = [df for df in all_dataframes if not df.empty]
        if all_dataframes:
            final_df = pd.concat(all_dataframes, ignore_index=True)
            cleaned_df = clean_df(final_df)
            excel_path = Path(temp_dir.name) / "Cleaned_Invoices.xlsx"
            cleaned_df.to_excel(excel_path, index=False)
            st.success("✅ Done! Download the Excel file below.")
            with open(excel_path, "rb") as f:
                st.download_button("⬇️ Download Excel File", f, file_name="Invoices_Cleaned.xlsx")
        else:
            st.warning("⚠️ No valid data found in uploaded PDFs.")
