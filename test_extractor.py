"""
Self-test script for the PDF to Excel extractor functions.
"""
import io
import pandas as pd
from app import extract_company_name, extract_date, extract_amount, extract_invoice_number, generate_excel

# Sample raw text simulating an invoice
sample_text_1 = """
ACME TEKNOLOJİ VE BİLİŞİM A.Ş.
Vergi No: 1234567890
Fatura No: INV-2024-00981
Fatura Tarihi: 14.05.2024

Açıklama             Miktar    Birim Fiyat    Tutar
Yazılım Danışmanlık       1      10.000,00    10.000,00
KDV %20                                        2.000,00
-------------------------------------------------------
Genel Toplam:                               12.000,00 TL
"""

sample_text_2 = """
GLOBAL SUPPLIERS LLC
Invoice Number: INV-88219
Invoice Date: 2024-04-18

Item Description      Qty       Rate        Amount
Cloud Hosting           1     250.00        250.00
--------------------------------------------------
Grand Total:                               $250.00
"""

def test():
    # Test 1
    comp1 = extract_company_name(sample_text_1)
    date1 = extract_date(sample_text_1)
    inv1 = extract_invoice_number(sample_text_1)
    amt1 = extract_amount(sample_text_1)
    print("Test 1 Result:", comp1, date1, inv1, amt1)
    assert comp1 == "ACME TEKNOLOJİ VE BİLİŞİM A.Ş."
    assert date1 == "14.05.2024"
    assert inv1 == "INV-2024-00981"
    assert amt1 == 12000.0

    # Test 2
    comp2 = extract_company_name(sample_text_2)
    date2 = extract_date(sample_text_2)
    inv2 = extract_invoice_number(sample_text_2)
    amt2 = extract_amount(sample_text_2)
    print("Test 2 Result:", comp2, date2, inv2, amt2)
    assert comp2 == "GLOBAL SUPPLIERS LLC"
    assert date2 == "2024-04-18"
    assert inv2 == "INV-88219"
    assert amt2 == 250.0

    # Test Excel generation
    df = pd.DataFrame([
        {"Filename": "test1.pdf", "Company": comp1, "Date": date1, "Invoice No": inv1, "Total Amount": amt1, "Status": "Success"},
        {"Filename": "test2.pdf", "Company": comp2, "Date": date2, "Invoice No": inv2, "Total Amount": amt2, "Status": "Success"}
    ])
    excel_bytes = generate_excel(df)
    assert len(excel_bytes) > 0
    print("Excel generation successful! File size:", len(excel_bytes), "bytes")
    print("All unit tests passed successfully!")

if __name__ == "__main__":
    test()
