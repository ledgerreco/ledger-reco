"""Generate realistic sample ledgers: Tally-style books (xlsx) + vendor statement (xlsx, csv, pdf x2)."""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font
import pandas as pd

HERE = Path(__file__).parent

# ---------------- Our books: Tally "Ledger Vouchers" export for vendor Sharma Traders --------------
# Purchases are Credit to the vendor, payments are Debit.
tally = [
    # date, particulars, vch type, vch no, debit, credit
    ("1-Apr-2024", "Opening Balance", "", "", None, 50000.00),
    ("5-Apr-2024", "Purchase A/c", "Purchase", "SHT/101", None, 118000.00),
    (None, "Being goods purchased vide bill no 101", "", "", None, None),          # Tally narration line
    ("12-Apr-2024", "Purchase A/c", "Purchase", "SHT/102", None, 59000.00),
    ("20-Apr-2024", "HDFC Bank", "Payment", "15", 100000.00, None),
    (None, "NEFT UTR HDFCN24111045", "", "", None, None),
    ("3-May-2024", "Purchase A/c", "Purchase", "SHT/103", None, 23600.00),           # vendor has 24,600
    ("15-May-2024", "Purchase A/c", "Purchase", "SHT/104", None, 35400.00),
    ("28-May-2024", "HDFC Bank", "Payment", "18", 50000.00, None),                  # vendor books on 15-Jun
    ("18-Jun-2024", "Purchase A/c", "Purchase", "SHT/106", None, 70800.00),
    ("5-Jul-2024", "HDFC Bank", "Payment", "22", 80000.00, None),
    (None, "Chq No 004512", "", "", None, None),
    ("31-Jul-2024", "TDS on Purchase 194Q", "Journal", "7", 1180.00, None),          # vendor 1,180.40
    ("10-Aug-2024", "Purchase A/c", "Purchase", "SHT/107", None, 94400.00),
    ("22-Aug-2024", "Purchase A/c", "Purchase", "SHT/108", None, 11800.00),
    ("23-Aug-2024", "Purchase A/c", "Purchase", "SHT/108", None, 11800.00),          # duplicate entry!
    ("15-Sep-2024", "HDFC Bank", "Payment", "31", 120000.00, None),
    ("30-Sep-2024", "Purchase Return", "Debit Note", "DN/04", 5900.00, None),        # vendor hasn't booked
]
deb = sum(r[4] or 0 for r in tally[1:])
cre = sum(r[5] or 0 for r in tally)
closing_cr = cre - deb  # credit balance

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Ledger"
for line in ["ABC Enterprises Pvt Ltd", "Plot 12, MIDC Andheri East, Mumbai", "Sharma Traders",
             "Ledger Account", "1-Apr-2024 to 30-Sep-2024", None]:
    ws.append([line])
ws["A1"].font = Font(bold=True, size=13)
ws.append(["Date", "Particulars", "Vch Type", "Vch No.", "Debit", "Credit"])
for r in tally:
    ws.append(list(r))
ws.append([None, "Closing Balance", "", "", closing_cr, None])      # Tally puts it on opposite side
ws.append([None, None, None, None, deb + closing_cr, cre])            # totals row
for col, w in zip("ABCDEF", (14, 32, 12, 10, 14, 14)):
    ws.column_dimensions[col].width = w
wb.save(HERE / "our_books_tally_sharma_traders.xlsx")

# ---------------- Vendor statement (their books: sales Debit to us, receipts Credit) ----------------
vendor = [
    ("01/04/2024", "", "Opening Balance", 50000.00, None),
    ("05/04/2024", "SI-101", "Sales Invoice", 118000.00, None),
    ("12/04/2024", "SI-102", "Sales Invoice", 59000.00, None),
    ("22/04/2024", "RCPT-0045", "Receipt - NEFT", None, 100000.00),
    ("03/05/2024", "SI-103", "Sales Invoice", 24600.00, None),       # amount mismatch vs 23,600
    ("15/05/2024", "SI-104", "Sales Invoice", 35400.00, None),
    ("15/06/2024", "RCPT-0061", "Receipt - RTGS", None, 50000.00),   # 18 days after our payment
    ("18/06/2024", "SI-106", "Sales Invoice", 70800.00, None),
    ("08/07/2024", "CHQ 004512", "Receipt - Cheque", None, 80000.00),
    ("31/07/2024", "TDS-Q1", "TDS deducted by customer", None, 1180.40),
    ("09/08/2024", "SI-107", "Sales Invoice", 94400.00, None),
    ("22/08/2024", "SI-108", "Sales Invoice", 11800.00, None),
    ("16/09/2024", "RCPT-0088", "Receipt - NEFT", None, 120000.00),
    ("25/09/2024", "SI-110", "Sales Invoice", 29500.00, None),       # not in our books (in transit)
    ("30/09/2024", "INT/09", "Interest on delayed payment", 1250.00, None),  # not in our books
]
bal, rows = 0.0, []
for d, ref, desc, dr, cr in vendor:
    bal += (dr or 0) - (cr or 0)
    rows.append((d, ref, desc, dr, cr, f"{bal:,.2f} Dr"))
cols = ["Date", "Invoice No", "Description", "Debit (₹)", "Credit (₹)", "Balance"]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Statement"
ws.append(["SHARMA TRADERS"])
ws.append(["Statement of Account — Customer: ABC Enterprises Pvt Ltd"])
ws.append(["Period: 01-04-2024 to 30-09-2024"])
ws.append([])
ws.append(cols)
for r in rows:
    ws.append(list(r))
wb.save(HERE / "party_statement_sharma_traders.xlsx")
pd.DataFrame(rows, columns=cols).to_csv(HERE / "party_statement_sharma_traders.csv", index=False)

# ---------------- PDFs ----------------
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.pdfgen import canvas

fmt = lambda v: "" if v is None else f"{v:,.2f}"  # noqa: E731
styles = getSampleStyleSheet()
doc = SimpleDocTemplate(str(HERE / "party_statement_sharma_traders.pdf"), pagesize=A4)
data = [["Date", "Invoice No", "Description", "Debit", "Credit", "Balance"]]
data += [[d, ref, desc, fmt(dr), fmt(cr), b] for d, ref, desc, dr, cr, b in rows]
t = Table(data, repeatRows=1)
t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                       ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                       ("ALIGN", (3, 1), (-1, -1), "RIGHT"), ("FONTSIZE", (0, 0), (-1, -1), 8)]))
doc.build([Paragraph("SHARMA TRADERS", styles["Title"]),
           Paragraph("Statement of Account — Customer: ABC Enterprises Pvt Ltd (01-04-2024 to 30-09-2024)",
                     styles["Normal"]), Spacer(1, 12), t])

# Plain-text PDF (no ruled table) -> exercises the text-line fallback parser
c = canvas.Canvas(str(HERE / "party_statement_sharma_traders_plain.pdf"), pagesize=A4)
c.setFont("Helvetica-Bold", 12)
c.drawString(40, 800, "SHARMA TRADERS - Statement of Account")
c.setFont("Helvetica", 9)
c.drawString(40, 785, "Customer: ABC Enterprises Pvt Ltd   Period: 01-Apr-2024 to 30-Sep-2024")
y = 760
xs = {"Date": 40, "Ref": 105, "Desc": 175}
right = {"Debit": 400, "Credit": 470, "Balance": 560}
c.setFont("Helvetica-Bold", 9)
c.drawString(xs["Date"], y, "Date"); c.drawString(xs["Ref"], y, "Ref"); c.drawString(xs["Desc"], y, "Particulars")
for k, x in right.items():
    c.drawRightString(x, y, k)
c.setFont("Helvetica", 9)
for d, ref, desc, dr, cr, b in rows:
    y -= 16
    dd = pd.to_datetime(d, dayfirst=True).strftime("%d-%b-%Y")
    c.drawString(xs["Date"], y, dd); c.drawString(xs["Ref"], y, ref); c.drawString(xs["Desc"], y, desc)
    c.drawRightString(right["Debit"], y, fmt(dr)); c.drawRightString(right["Credit"], y, fmt(cr))
    c.drawRightString(right["Balance"], y, b)
c.save()

# Image-only PDF (simulates a scanned document; no text layer)
c = canvas.Canvas(str(HERE / "scanned_example.pdf"), pagesize=A4)
c.rect(50, 500, 400, 200, fill=1)
c.save()
print("samples written; our closing Cr =", closing_cr, "vendor closing Dr =", bal)


# ======================= AI samples =======================================================================
# v2 pair: everything above PLUS cases that one-to-one rules cannot solve:
#   * one payment of ₹1,47,500 in our books vs two receipts (₹59,000 + ₹88,500) in the party's books
#   * one party invoice SI-113 ₹35,400 booked by us as two entries SHT/113 ₹23,600 + SHT/113A ₹11,800
tally_v2 = tally + [
    ("2-Sep-2024", "Purchase A/c", "Purchase", "SHT/111", None, 59000.00),
    ("4-Sep-2024", "Purchase A/c", "Purchase", "SHT/112", None, 88500.00),
    ("6-Sep-2024", "Purchase A/c", "Purchase", "SHT/113", None, 23600.00),
    ("6-Sep-2024", "Purchase A/c (freight portion)", "Purchase", "SHT/113A", None, 11800.00),
    ("20-Sep-2024", "HDFC Bank", "Payment", "34", 147500.00, None),
    (None, "NEFT against bills 111 & 112", "", "", None, None),
]
# insert before the last (30-Sep) entry so dates stay in order; narration line stays after its voucher
tally_v2 = tally[:-1] + tally_v2[len(tally):] + [tally[-1]]
deb2 = sum(r[4] or 0 for r in tally_v2[1:])
cre2 = sum(r[5] or 0 for r in tally_v2)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Ledger"
for line in ["ABC Enterprises Pvt Ltd", "Plot 12, MIDC Andheri East, Mumbai", "Sharma Traders",
             "Ledger Account", "1-Apr-2024 to 30-Sep-2024", None]:
    ws.append([line])
ws.append(["Date", "Particulars", "Vch Type", "Vch No.", "Debit", "Credit"])
for r in tally_v2:
    ws.append(list(r))
ws.append([None, "Closing Balance", "", "", cre2 - deb2, None])
ws.append([None, None, None, None, cre2, cre2])
wb.save(HERE / "ai_our_books_tally_v2.xlsx")

vendor_v2 = vendor + [
    ("02/09/2024", "SI-111", "Sales Invoice", 59000.00, None),
    ("04/09/2024", "SI-112", "Sales Invoice", 88500.00, None),
    ("06/09/2024", "SI-113", "Sales Invoice (incl. freight)", 35400.00, None),
    ("21/09/2024", "RCPT-0091", "Receipt against SI-111", None, 59000.00),
    ("21/09/2024", "RCPT-0092", "Receipt against SI-112", None, 88500.00),
]
vendor_v2.sort(key=lambda r: pd.to_datetime(r[0], dayfirst=True))
bal, rows2 = 0.0, []
for d, ref, desc, dr, cr in vendor_v2:
    bal += (dr or 0) - (cr or 0)
    rows2.append((d, ref, desc, dr, cr, f"{bal:,.2f} Dr"))
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["SHARMA TRADERS"])
ws.append(["Statement of Account — Customer: ABC Enterprises Pvt Ltd"])
ws.append([])
ws.append(cols)
for r in rows2:
    ws.append(list(r))
wb.save(HERE / "ai_party_statement_v2.xlsx")

# "Scanned" party statement: render the table PDF to a slightly rotated, noisy grayscale image,
# saved as JPG and as an image-only PDF (no text layer).
import random
import pypdfium2 as pdfium
from PIL import Image, ImageFilter

random.seed(7)
page = pdfium.PdfDocument(str(HERE / "party_statement_sharma_traders.pdf"))[0]
img = page.render(scale=2.2).to_pil().convert("L").rotate(0.6, expand=True, fillcolor=255)
px = img.load()
for _ in range(25000):
    x, y = random.randrange(img.width), random.randrange(img.height)
    px[x, y] = random.choice((150, 190, 230))
img = img.filter(ImageFilter.GaussianBlur(0.5))
img.save(HERE / "scanned_party_statement.jpg", quality=80)
img.convert("RGB").save(HERE / "scanned_party_statement.pdf", "PDF", resolution=200)
print("AI samples written")
