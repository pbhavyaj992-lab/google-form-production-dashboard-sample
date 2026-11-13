"""Generate the synthetic sample data and the Excel dashboard workbook.

Run: python3 generate.py
Outputs:
  data/sample_production_records.csv
  excel/production_dashboard_sample.xlsx

All data is synthetic and invented for demonstration only.
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
random.seed(42)

LINES = ["Line A", "Line B", "Line C"]
PRODUCTS = {
    "Line A": ["Detergent Powder 1kg", "Detergent Powder 5kg"],
    "Line B": ["Liquid Soap 500ml", "Liquid Soap 1L"],
    "Line C": ["Floor Cleaner 1L", "Dishwash Liquid 750ml"],
}
MATERIALS = ["Caustic Soda", "Soda Ash", "SLES", "Fragrance", "Packaging Film", "HDPE Bottles"]
SUPPLIERS = ["Supplier Alpha", "Supplier Beta", "Supplier Gamma"]
SHIFTS = ["Morning (06:00-14:00)", "Afternoon (14:00-22:00)", "Night (22:00-06:00)"]
SUPERVISORS = ["Supervisor 1", "Supervisor 2", "Supervisor 3"]
OPERATORS = ["Operator A", "Operator B", "Operator C", "Operator D", "Operator E", "Operator F"]

START = date(2026, 10, 26)
DAYS = 14

rows = []
for day in range(DAYS):
    d = START + timedelta(days=day)
    for shift in SHIFTS:
        for line in LINES:
            product = random.choice(PRODUCTS[line])
            batch = f"B{str(d).replace('-', '')[2:]}-{line[-1]}{SHIFTS.index(shift) + 1}"
            material = random.choice(MATERIALS)
            supplier = random.choice(SUPPLIERS)
            # One inward receipt row and one outward issue row per line/shift,
            # mirroring a Form that logs one material movement per submission.
            qty_received = random.choice([0, 0, 250, 500, 750, 1000])
            qty_issued = random.randint(120, 480)
            units = random.randint(800, 4200)
            defects = random.randint(0, max(1, units // 60))
            downtime = random.choice([0, 0, 10, 15, 20, 30, 45, 60])
            supervisor = SUPERVISORS[SHIFTS.index(shift)]
            operator = random.choice(OPERATORS)
            rows.append([
                d.isoformat(), shift, supervisor, line, product, batch,
                material, supplier, "Inward", qty_received, 0,
                units, defects, downtime, operator,
            ])
            rows.append([
                d.isoformat(), shift, supervisor, line, product, batch,
                material, supplier, "Outward", 0, qty_issued,
                units, defects, downtime, operator,
            ])

# Shuffle so the file reads like real form submissions arriving over time.
random.shuffle(rows)

HEADERS = [
    "date", "shift", "supervisor", "production_line", "product", "batch_no",
    "material", "supplier", "movement_type", "qty_received_kg",
    "qty_issued_kg", "units_produced", "defects", "downtime_minutes", "operator",
]

csv_path = ROOT / "data" / "sample_production_records.csv"
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(HEADERS)
    w.writerows(rows)
print(f"Wrote {csv_path} ({len(rows)} data rows)")

# ---------------------------------------------------------------- Excel workbook
wb = openpyxl.Workbook()

NAVY = "1F3864"
TEAL = "2E75B6"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
RED = "FFC7CE"
GREEN = "C6EFCE"
hdr_font = Font(bold=True, color="FFFFFF", size=11)
hdr_fill = PatternFill("solid", fgColor=NAVY)
sub_font = Font(bold=True, size=11, color=NAVY)
kpi_font = Font(bold=True, size=16, color=NAVY)
kpi_label = Font(size=10, color="595959")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---- Data sheet: the clean table the Sheet/App Script would maintain.
ws = wb.active
ws.title = "Data"
for col, h in enumerate(HEADERS, 1):
    c = ws.cell(row=1, column=col, value=h.replace("_", " ").title())
    c.font = hdr_font
    c.fill = hdr_fill
    c.border = border
for r, row in enumerate(rows, 2):
    for col, v in enumerate(row, 1):
        c = ws.cell(row=r, column=col, value=v)
        c.border = border
        if r % 2 == 0:
            c.fill = PatternFill("solid", fgColor=GREY)
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions
widths = [12, 24, 14, 15, 24, 14, 16, 15, 13, 16, 14, 15, 10, 17, 12]
for i, wdt in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = wdt

n = len(rows) + 1  # last data row (row 1 is header)

# ---- Summary sheet: pivot-style tables built with live formulas.
sm = wb.create_sheet("Summary")
sm["A1"] = "Production summary (live formulas over the Data sheet)"
sm["A1"].font = Font(bold=True, size=13, color=NAVY)

sm["A3"] = "Units produced by production line"
sm["A3"].font = sub_font
sm["A4"] = "Production line"
sm["B4"] = "Units produced"
sm["C4"] = "Defects"
sm["D4"] = "Defect rate"
for col in "ABCD":
    sm[f"{col}4"].font = hdr_font
    sm[f"{col}4"].fill = hdr_fill
for i, line in enumerate(LINES):
    r = 5 + i
    sm.cell(row=r, column=1, value=line)
    # Units/defects counted once per line/shift: only Inward rows carry them
    # (each line/shift appears twice, Inward + Outward, same units).
    sm.cell(row=r, column=2).value = (
        f'=SUMIFS(Data!$L$2:$L${n},Data!$D$2:$D${n},$A{r},Data!$I$2:$I${n},"Inward")'
    )
    sm.cell(row=r, column=3).value = (
        f'=SUMIFS(Data!$M$2:$M${n},Data!$D$2:$D${n},$A{r},Data!$I$2:$I${n},"Inward")'
    )
    sm.cell(row=r, column=4).value = f'=IF(B{r}=0,"-",C{r}/B{r})'
    sm.cell(row=r, column=4).number_format = "0.00%"
tr = 5 + len(LINES)
sm.cell(row=tr, column=1, value="Total").font = Font(bold=True)
sm.cell(row=tr, column=2).value = f"=SUM(B5:B{tr - 1})"
sm.cell(row=tr, column=2).font = Font(bold=True)
sm.cell(row=tr, column=3).value = f"=SUM(C5:C{tr - 1})"
sm.cell(row=tr, column=3).font = Font(bold=True)
sm.cell(row=tr, column=4).value = f"=IF(B{tr}=0,\"-\",C{tr}/B{tr})"
sm.cell(row=tr, column=4).number_format = "0.00%"
sm.cell(row=tr, column=4).font = Font(bold=True)

r0 = tr + 2
sm.cell(row=r0, column=1, value="Material received vs issued (kg)").font = sub_font
sm.cell(row=r0 + 1, column=1, value="Material").font = hdr_font
sm.cell(row=r0 + 1, column=1).fill = hdr_fill
sm.cell(row=r0 + 1, column=2, value="Received (kg)").font = hdr_font
sm.cell(row=r0 + 1, column=2).fill = hdr_fill
sm.cell(row=r0 + 1, column=3, value="Issued (kg)").font = hdr_font
sm.cell(row=r0 + 1, column=3).fill = hdr_fill
sm.cell(row=r0 + 1, column=4, value="Balance (kg)").font = hdr_font
sm.cell(row=r0 + 1, column=4).fill = hdr_fill
for i, mat in enumerate(MATERIALS):
    r = r0 + 2 + i
    sm.cell(row=r, column=1, value=mat)
    sm.cell(row=r, column=2).value = f"=SUMIFS(Data!$J$2:$J${n},Data!$G$2:$G${n},$A{r})"
    sm.cell(row=r, column=3).value = f"=SUMIFS(Data!$K$2:$K${n},Data!$G$2:$G${n},$A{r})"
    sm.cell(row=r, column=4).value = f"=B{r}-C{r}"
tr2 = r0 + 2 + len(MATERIALS)
sm.cell(row=tr2, column=1, value="Total").font = Font(bold=True)
for col in (2, 3, 4):
    sm.cell(row=tr2, column=col).value = f"=SUM({get_column_letter(col)}{r0 + 2}:{get_column_letter(col)}{tr2 - 1})"
    sm.cell(row=tr2, column=col).font = Font(bold=True)

r1 = tr2 + 2
sm.cell(row=r1, column=1, value="Daily units produced").font = sub_font
sm.cell(row=r1 + 1, column=1, value="Date").font = hdr_font
sm.cell(row=r1 + 1, column=1).fill = hdr_fill
sm.cell(row=r1 + 1, column=2, value="Units produced").font = hdr_font
sm.cell(row=r1 + 1, column=2).fill = hdr_fill
sm.cell(row=r1 + 1, column=3, value="Downtime (min)").font = hdr_font
sm.cell(row=r1 + 1, column=3).fill = hdr_fill
for i in range(DAYS):
    r = r1 + 2 + i
    d = (START + timedelta(days=i)).isoformat()
    sm.cell(row=r, column=1, value=d)
    sm.cell(row=r, column=2).value = (
        f'=SUMIFS(Data!$L$2:$L${n},Data!$A$2:$A${n},$A{r},Data!$I$2:$I${n},"Inward")'
    )
    sm.cell(row=r, column=3).value = (
        f'=SUMIFS(Data!$N$2:$N${n},Data!$A$2:$A${n},$A{r},Data!$I$2:$I${n},"Inward")'
    )
for col, wdt in zip("ABCD", [26, 18, 12, 14]):
    sm.column_dimensions[col].width = wdt

# ---- Dashboard sheet: KPI cards (formulas) + charts fed by Summary.
db = wb.create_sheet("Dashboard", 0)
db.sheet_properties.pageSetUpPr.fitToPage = True
db.page_setup.orientation = "landscape"
db.page_setup.fitToWidth = 1
db.page_setup.fitToHeight = 1
db["A1"] = "Production Dashboard  (sample)"
db["A1"].font = Font(bold=True, size=16, color=NAVY)
db["A2"] = "Google Form-linked production dashboard — sample data, 26 Oct – 8 Nov 2026 (synthetic)"
db["A2"].font = Font(size=11, color="595959")
db["A3"] = "In the live system these KPIs read from the Google Sheet fed by the Form. Here they read from the Data sheet in this workbook."
db["A3"].font = Font(size=10, color="595959", italic=True)

kpis = [
    ("Total units produced", f"=Summary!B{tr}", "#,##0"),
    ("Total defects", f"=Summary!C{tr}", "#,##0"),
    ("Overall defect rate", f"=Summary!D{tr}", "0.00%"),
    ("Total downtime (min)", f"=SUM(Summary!C{r1 + 2}:C{r1 + 1 + DAYS})", "#,##0"),
    ("Material received (kg)", f"=Summary!B{tr2}", "#,##0"),
    ("Material issued (kg)", f"=Summary!C{tr2}", "#,##0"),
]
db["A5"] = "Key metrics"
db["A5"].font = sub_font
for i, (label, formula, fmt) in enumerate(kpis):
    r = 6 + i
    db.cell(row=r, column=1, value=label).font = Font(size=11)
    v = db.cell(row=r, column=2)
    v.value = formula
    v.font = kpi_font
    v.number_format = fmt
db.column_dimensions["A"].width = 30
db.column_dimensions["B"].width = 20

# Charts from the Summary sheet.
bar = BarChart()
bar.type = "col"
bar.title = "Units produced by production line"
bar.y_axis.title = "Units"
bar.x_axis.title = "Line"
bar.height = 7.5
bar.width = 14
data = Reference(sm, min_col=2, min_row=4, max_row=4 + len(LINES))
cats = Reference(sm, min_col=1, min_row=5, max_row=4 + len(LINES))
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.shape = 4
db.add_chart(bar, "D5")

line = LineChart()
line.title = "Daily units produced"
line.y_axis.title = "Units"
line.x_axis.title = "Date"
line.height = 7.5
line.width = 14
data2 = Reference(sm, min_col=2, min_row=r1 + 1, max_row=r1 + 1 + DAYS)
cats2 = Reference(sm, min_col=1, min_row=r1 + 2, max_row=r1 + 1 + DAYS)
line.add_data(data2, titles_from_data=True)
line.set_categories(cats2)
db.add_chart(line, "D22")

db["A14"] = "How to refresh (live system)"
db["A14"].font = sub_font
notes = [
    "1. Supervisors submit the Google Form from any phone.",
    "2. Responses land in the linked Google Sheet; Apps Script (Code.gs) cleans and flags them.",
    "3. This Excel file pulls the Sheet via Power Query (Data > Refresh All) and updates in under 30 seconds.",
    "4. Print this sheet for the daily summary page.",
]
for i, note in enumerate(notes):
    db.cell(row=15 + i, column=1, value=note).font = Font(size=10)
db.merge_cells("A15:B15")
db.merge_cells("A16:B16")
db.merge_cells("A17:B17")
db.merge_cells("A18:B18")

# ---- How-to sheet: Power Query steps.
hq = wb.create_sheet("How to refresh")
hq["A1"] = "Connecting this dashboard to the live Google Sheet (Power Query)"
hq["A1"].font = Font(bold=True, size=13, color=NAVY)
steps = [
    "1. In Google Sheets: File > Share > Publish to web > choose the 'Clean_Data' tab > CSV > Publish. Copy the link.",
    "2. In Excel: Data tab > Get Data > From Other Sources > From Web. Paste the published CSV link.",
    "3. In the Power Query Editor: keep the same column names as the Data sheet here, set Date to Date type and quantities to Whole Number.",
    "4. Close & Load To... > Existing worksheet > the Data sheet, replacing the sample rows.",
    "5. Tick Data > Queries & Connections > right-click the query > Properties > 'Refresh data when opening the file'.",
    "6. From then on, opening the workbook (or Data > Refresh All) pulls the latest Form responses in under 30 seconds.",
    "",
    "Note: publishing the CSV link makes that tab readable by anyone with the link. Share the Sheet itself only with the Plant Manager and Owner.",
]
for i, s in enumerate(steps):
    c = hq.cell(row=3 + i, column=1, value=s)
    c.font = Font(size=11)
    hq.merge_cells(start_row=3 + i, start_column=1, end_row=3 + i, end_column=8)
hq.column_dimensions["A"].width = 140

xlsx_path = ROOT / "excel" / "production_dashboard_sample.xlsx"
wb.save(xlsx_path)
print(f"Wrote {xlsx_path}")
