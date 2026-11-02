import io
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
import pdfplumber
import streamlit as st


# ---------------------------------------------------------------------------
# Data Extraction Helper Functions
# ---------------------------------------------------------------------------

def extract_date(text: str) -> Optional[str]:
    """
    Extracts the document date from the text using common prefixes or standalone date patterns.
    Supports formats like DD.MM.YYYY, DD/MM/YYYY, YYYY-MM-DD, etc.
    """
    # 1. Check for labeled dates (e.g., Date:, Tarih:, Invoice Date:)
    labeled_pattern = re.compile(
        r"(?:date|invoice\s*date|issue\s*date|tarih|fatura\s*tarihi|d[üu]zenleme\s*tarihi)"
        r"[\s\:\-]+(?P<date>\d{1,2}[\.\/\-]\d{1,2}[\.\/\-]\d{2,4}|\d{4}[\.\/\-]\d{1,2}[\.\/\-]\d{1,2})",
        re.IGNORECASE,
    )
    match = labeled_pattern.search(text)
    if match:
        return match.group("date").strip()

    # 2. Fallback: Search for any standard date formatted string
    date_fallback_pattern = re.compile(
        r"\b(\d{1,2}[\.\/\-]\d{1,2}[\.\/\-]\d{2,4}|\d{4}[\.\/\-]\d{1,2}[\.\/\-]\d{1,2})\b"
    )
    matches = date_fallback_pattern.findall(text)
    if matches:
        return matches[0]

    return None


def extract_amount(text: str) -> Optional[float]:
    """
    Extracts the grand total or amount from text.
    Searches for terms like 'Total', 'Grand Total', 'Toplam Tutar', 'Ödenecek Tutar'.
    Handles both period and comma as decimal and thousand separators, with support
    for various currencies (USD, EUR, TRY, TL, $, €, £, ₺).
    """
    # Specific keywords indicating a total balance
    keywords = [
        r"grand\s*total",
        r"total\s*amount",
        r"total\s*due",
        r"amount\s*due",
        r"[öo]denecek\s*tutar",
        r"genel\s*toplam",
        r"toplam\s*tutar",
        r"toplam",
        r"total",
        r"tutar",
    ]
    keyword_group = "|".join(keywords)

    # Unicode-safe currency pattern ($ € £ ₺ USD EUR TRY TL)
    curr_pattern = r"(?:[\$\u20ac\u00a3\u20ba]|USD|EUR|TRY|TL)?"

    # Pattern capturing amount:
    # 1. Thousand separated: e.g. 1,250.00 or 1.250,00
    # 2. Direct decimals or integers: e.g. 250.00 or 250
    pattern = re.compile(
        rf"(?:{keyword_group})[\s:\-]+{curr_pattern}\s*"
        r"([0-9]{1,3}(?:[.,\s][0-9]{3})*(?:[.,][0-9]{2})|[0-9]+(?:[.,][0-9]{2})?)",
        re.IGNORECASE,
    )

    # Find all matches and prioritize the last match (totals are usually at bottom of invoices)
    all_matches = list(pattern.finditer(text))
    raw_amount = None

    if all_matches:
        raw_amount = all_matches[-1].group(1).strip()
    else:
        # Fallback: look for lines ending in currency format
        fallback_pattern = re.compile(
            rf"{curr_pattern}\s*([0-9]{{1,3}}(?:[.,\s][0-9]{{3}})*(?:[.,][0-9]{{2}})|[0-9]+(?:[.,][0-9]{{2}})?)\s*"
            rf"(?:[\$\u20ac\u00a3\u20ba]|USD|EUR|TRY|TL)\b",
            re.IGNORECASE,
        )
        fallback_matches = fallback_pattern.findall(text)
        if fallback_matches:
            raw_amount = fallback_matches[-1]

    if raw_amount:
        # Normalize the amount string to float
        cleaned = raw_amount.replace(" ", "")

        # Handle comma as decimal separator (e.g. 1.250,50 -> 1250.50)
        if "," in cleaned and "." in cleaned:
            if cleaned.rfind(",") > cleaned.rfind("."):
                # 1.250,50 -> comma is decimal
                cleaned = cleaned.replace(".", "").replace(",", ".")
            else:
                # 1,250.50 -> period is decimal
                cleaned = cleaned.replace(",", "")
        elif "," in cleaned:
            # Check if comma indicates cents (e.g. 125,50)
            parts = cleaned.split(",")
            if len(parts[-1]) == 2:
                cleaned = cleaned.replace(",", ".")
            else:
                cleaned = cleaned.replace(",", "")

        try:
            return float(cleaned)
        except ValueError:
            return None

    return None


def extract_company_name(text: str) -> Optional[str]:
    """
    Extracts company or vendor name.
    First looks for explicit labels like 'Company:', 'Vendor:', 'Firma:', 'Satıcı:'.
    If not found, scans the top non-empty lines ignoring generic receipt headings.
    """
    # 1. Check for labeled vendor name
    label_pattern = re.compile(
        r"(?:company(?:\s*name)?|vendor|supplier|seller|firma(?:\s*ad[ıi])?|sat[ıi]c[ıi]|unvan|ünvan)"
        r"[\s\:\-]+([^\n\r]+)",
        re.IGNORECASE,
    )
    match = label_pattern.search(text)
    if match:
        name = match.group(1).strip()
        if len(name) > 2:
            return name

    # 2. Fallback: inspect top lines of the document
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    skip_keywords = [
        "invoice", "receipt", "fatura", "e-fatura", "e-arşiv", "makbuz",
        "tax invoice", "bill", "fiş", "bilgi fişi", "page", "sayfa"
    ]

    for line in lines[:8]:  # inspect first 8 lines
        lower_line = line.lower()
        if any(keyword in lower_line for keyword in skip_keywords):
            continue
        # Avoid lines that are just numbers or dates
        if re.match(r"^[\d\W]+$", line):
            continue
        # A good candidate is usually between 3 and 60 characters
        if 3 <= len(line) <= 60:
            return line

    return None


def extract_invoice_number(text: str) -> Optional[str]:
    """
    Extracts invoice or receipt document reference number.
    """
    pattern = re.compile(
        r"(?:invoice\s*(?:no|number|#)|fatura\s*no|receipt\s*(?:no|#)|belge\s*no|fi[şs]\s*no)"
        r"[\s\:\-]+([A-Za-z0-9\-_/]+)",
        re.IGNORECASE,
    )
    match = pattern.search(text)
    if match:
        return match.group(1).strip()
    return None


def process_pdf(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """
    Reads a single PDF file using pdfplumber and extracts structured fields.
    """
    full_text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
    except Exception as exc:
        return {
            "Filename": filename,
            "Company": "Error reading PDF",
            "Date": None,
            "Invoice No": None,
            "Total Amount": None,
            "Status": f"Failed: {str(exc)}",
            "Raw Text": "",
        }

    if not full_text.strip():
        return {
            "Filename": filename,
            "Company": "Empty or Scanned PDF",
            "Date": None,
            "Invoice No": None,
            "Total Amount": None,
            "Status": "Warning: No readable text found (may be an image/scanned PDF)",
            "Raw Text": "",
        }

    company = extract_company_name(full_text)
    date_val = extract_date(full_text)
    invoice_no = extract_invoice_number(full_text)
    total_amount = extract_amount(full_text)

    status = "Success"
    if not (company and date_val and total_amount is not None):
        status = "Partial (Check fields)"

    return {
        "Filename": filename,
        "Company": company or "N/A",
        "Date": date_val or "N/A",
        "Invoice No": invoice_no or "N/A",
        "Total Amount": total_amount,
        "Status": status,
        "Raw Text": full_text,
    }


# ---------------------------------------------------------------------------
# Excel Generation Helper Function (openpyxl)
# ---------------------------------------------------------------------------

def generate_excel(df: pd.DataFrame) -> bytes:
    """
    Converts a pandas DataFrame into a formatted Excel file using openpyxl.
    Applies custom headers, column widths, number formatting, and row styling.
    """
    # Exclude raw text column from the final export
    export_df = df.drop(columns=["Raw Text"], errors="ignore")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Extracted Data")
        workbook = writer.book
        worksheet = writer.sheets["Extracted Data"]

        # Styling definitions
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Calibri", size=10)
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")

        # Apply header styling
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Format rows and columns
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.font = regular_font
                col_name = export_df.columns[cell.column - 1]

                if col_name == "Total Amount":
                    cell.alignment = right_align
                    # Apply currency formatting
                    cell.number_format = "#,##0.00"
                elif col_name in ["Date", "Invoice No", "Status"]:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

        # Auto-adjust column widths
        for col_idx, col in enumerate(worksheet.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for cell in col:
                val = str(cell.value or "")
                max_len = max(max_len, len(val))
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    return output.getvalue()


# ---------------------------------------------------------------------------
# Streamlit User Interface
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="PDF to Excel Data Extractor",
        page_icon="📊",
        layout="wide",
    )

    # Title & Subtitle
    st.title("📄 PDF to Excel Data Extractor")
    st.markdown(
        "Upload one or more invoice/receipt PDF files to automatically extract **Company**, "
        "**Date**, **Invoice No**, and **Total Amount**, then export the table to an Excel spreadsheet."
    )

    # Sidebar instructions
    with st.sidebar:
        st.header("ℹ️ Instructions")
        st.markdown(
            """
            1. **Upload PDFs**: Select or drag & drop one or multiple PDF documents.
            2. **Review Extracted Data**: Inspect the resulting table in the main panel.
            3. **Edit in Place**: You can double-click cells in the table to correct any value.
            4. **Export**: Click **Download as Excel** to save the `.xlsx` file.
            """
        )
        st.markdown("---")
        st.caption("Powered by **Streamlit**, **pdfplumber**, **pandas**, and **openpyxl**.")

    # 1. File Upload Section
    uploaded_files = st.file_uploader(
        "Select PDF files to process",
        type=["pdf"],
        accept_multiple_files=True,
        help="Upload single or multiple receipt / invoice PDF documents.",
    )

    if uploaded_files:
        st.info(f"Loaded **{len(uploaded_files)}** PDF document(s). Processing...")

        results = []
        progress_bar = st.progress(0)

        for index, uploaded_file in enumerate(uploaded_files):
            file_bytes = uploaded_file.read()
            extracted = process_pdf(file_bytes, uploaded_file.name)
            results.append(extracted)
            progress_bar.progress((index + 1) / len(uploaded_files))

        df = pd.DataFrame(results)

        # Summary Metrics
        st.subheader("📈 Summary")
        total_files = len(df)
        success_count = (df["Status"] == "Success").sum()
        total_sum = df["Total Amount"].dropna().sum()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Documents", total_files)
        col2.metric("Fully Parsed", f"{success_count}/{total_files}")
        col3.metric("Total Calculated Amount", f"{total_sum:,.2f}")

        # 2. Interactive Data Editor / Table
        st.subheader("📋 Extracted Data Table")
        st.caption("You can directly edit any cell before downloading to Excel.")

        display_columns = ["Filename", "Company", "Date", "Invoice No", "Total Amount", "Status"]
        edited_df = st.data_editor(
            df[display_columns],
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Total Amount": st.column_config.NumberColumn(
                    "Total Amount",
                    format="%.2f",
                ),
                "Status": st.column_config.TextColumn(
                    "Status",
                    disabled=True,
                ),
            },
        )

        # Merge edited fields back into full dataframe for export
        for col in display_columns:
            df[col] = edited_df[col]

        # 3. Export to Excel Button
        st.subheader("💾 Export Data")
        excel_data = generate_excel(df)
        current_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_filename = f"extracted_invoices_{current_time_str}.xlsx"

        st.download_button(
            label="📥 Download as Excel (.xlsx)",
            data=excel_data,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # 4. Raw Text Inspector (Optional Debugging / Verification)
        with st.expander("🔍 View Raw Extracted Text by File"):
            selected_file = st.selectbox("Select document to inspect:", df["Filename"].tolist())
            matched_row = df[df["Filename"] == selected_file].iloc[0]
            st.text_area(
                "Raw Text Content",
                value=matched_row["Raw Text"] if matched_row["Raw Text"] else "(No text extracted)",
                height=250,
            )

    else:
        st.write("---")
        st.info("👆 Please upload one or more PDF files using the file uploader above to begin extraction.")


if __name__ == "__main__":
    main()
