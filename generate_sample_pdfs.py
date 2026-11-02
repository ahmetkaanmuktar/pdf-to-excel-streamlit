"""
Script to generate sample invoice PDF files for testing the Streamlit app.
"""
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def create_sample_pdf(filename: str, company: str, inv_no: str, date_str: str, items: list, total: str):
    c = canvas.Canvas(filename, pagesize=letter)
    width, height = letter

    # Header / Company
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 60, company)

    # Document details
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 85, f"Invoice Number: {inv_no}")
    c.drawString(50, height - 100, f"Invoice Date: {date_str}")
    c.drawString(50, height - 115, "Status: Paid")

    # Table Header
    c.line(50, height - 135, width - 50, height - 135)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, height - 150, "Item Description")
    c.drawString(320, height - 150, "Qty")
    c.drawString(400, height - 150, "Price")
    c.drawString(480, height - 150, "Total")
    c.line(50, height - 160, width - 50, height - 160)

    # Items
    y = height - 180
    c.setFont("Helvetica", 10)
    for desc, qty, price, line_total in items:
        c.drawString(50, y, desc)
        c.drawString(320, y, str(qty))
        c.drawString(400, y, price)
        c.drawString(480, y, line_total)
        y -= 25

    # Footer / Grand Total
    c.line(50, y - 10, width - 50, y - 10)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(350, y - 30, "Grand Total:")
    c.drawString(470, y - 30, total)

    c.save()
    print(f"Generated {filename}")

if __name__ == "__main__":
    create_sample_pdf(
        "sample_invoice_1.pdf",
        company="NEXUS CLOUD SOLUTIONS LTD",
        inv_no="INV-2024-1082",
        date_str="25.04.2024",
        items=[
            ("Dedicated Server Hosting", 1, "$450.00", "$450.00"),
            ("SSL Certificate & DNS Setup", 2, "$75.00", "$150.00"),
            ("Database Backup Service", 1, "$120.00", "$120.00"),
        ],
        total="$720.00",
    )

    create_sample_pdf(
        "sample_invoice_2.pdf",
        company="ALPHA LOGISTICS & TRADE GMBH",
        inv_no="INV-88910",
        date_str="12.05.2024",
        items=[
            ("Freight Transport Service", 2, "EUR 600.00", "EUR 1,200.00"),
            ("Warehouse Storage Fee", 1, "EUR 350.00", "EUR 350.00"),
        ],
        total="EUR 1,550.00",
    )
