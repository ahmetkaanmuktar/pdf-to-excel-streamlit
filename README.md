# PDF to Excel Data Extractor Tool

A desktop tool built with **Streamlit**, **pdfplumber**, **pandas**, and **openpyxl** to automatically parse invoices, receipts, and simple PDF documents and export structured data into an Excel spreadsheet.

## Features
- **Multi-PDF Upload**: Drag and drop one or several PDF files simultaneously.
- **Smart Extraction Logic**:
  - **Company / Vendor Name** (explicit labels or document top-section analysis)
  - **Document Date** (supports standard formats like `DD.MM.YYYY`, `DD/MM/YYYY`, `YYYY-MM-DD`)
  - **Invoice / Receipt Number**
  - **Grand Total / Amount** (handles `$`, `€`, `£`, `₺`, `USD`, `EUR`, `TRY`, `TL`, and comma/dot decimals)
- **Interactive Data Table**: Edit cells directly in the Streamlit UI prior to exporting.
- **Professional Excel Export**: Styled header, auto-adjusted column widths, and proper currency number formatting.
- **Raw Text Inspector**: View extracted textual content for transparency and auditing.

## Installation

Install the required Python libraries using `pip`:

```bash
pip install streamlit pdfplumber pandas openpyxl
```

Or using `requirements.txt`:

```bash
pip install -r requirements.txt
```

## Running the Application

Execute the following command in your terminal:

```bash
streamlit run app.py
```

Or double-click `run_app.bat` on Windows.
