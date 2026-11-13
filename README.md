# Google Form-Linked Production Dashboard — SAMPLE

> **Sample project.** Built to demonstrate a Google Form-linked production
> dashboard for a batch-manufacturing unit. It is a demonstration, not a
> client delivery: all data in this repo is synthetic sample data, invented
> for testing the flow. There is no real client, no real factory and no real
> results behind it.

Inspired by a real Freelancer brief for a "Google Form-Linked Production
Dashboard": a batch-manufacturing unit where three supervisors submit a daily
Google Form (Batch Manufacturing Record, shift details and raw-material
movements), the Plant Manager monitors the live Google Sheet, and the Owner
gets an automatically refreshed Excel dashboard with the key KPIs.

## The problem it solves

Paper or WhatsApp-based shift reporting means the manager chases supervisors
for numbers, stock figures drift because "remaining stock" is typed by hand,
and the owner only sees production totals days later. This sample wires up
the three-tier flow from that brief:

1. **Supervisors** submit one short Google Form per material movement, from
   any phone. They never touch the Sheet directly.
2. **Plant Manager** watches a validated, read-only Sheet view with filters
   for date, shift, batch and material, plus colour alerts for missing or
   inconsistent entries.
3. **Owner** opens an Excel dashboard that pulls the same Sheet through
   Power Query and shows daily production totals, material consumption vs.
   receipts, stock on hand and variances — refreshed in under 30 seconds.

One design decision worth noting: a batch record usually has several
raw-material lines, so the Form logs **one material movement per
submission** (inward or outward), linked by batch number. Stock on hand is
**calculated** (`total received − total issued`), never typed in — typed
stock figures drift within days and make variance reports meaningless.

## What is in this repo

| File | What it is |
| --- | --- |
| `data/sample_production_records.csv` | 252 synthetic rows (14 days, 3 shifts, 3 production lines) in the shape of Form responses: date, shift, supervisor, line, product, batch, material, supplier, movement type, quantities, units produced, defects, downtime, operator. |
| `Code.gs` | Google Apps Script: validates each Form submission as it arrives, writes a clean row, calculates stock on hand, and flags inconsistent entries. |
| `excel/production_dashboard_sample.xlsx` | Working Excel dashboard. `Dashboard` sheet with KPI cards and charts, `Data` sheet with the sample rows, `Summary` sheet with live formulas (units by line, defect rate, received vs. issued, daily trend), and a `How to refresh` sheet with the Power Query steps. |
| `google-form-fields.md` | The Form design: every question, its type, whether it is required, and its validation. |
| `sheets-setup.md` | The Google Sheet design: tabs, validation rules, conditional formatting, the manager view and the formulas behind the KPIs. |
| `generate.py` | The script that generated the sample CSV and the Excel workbook, so the sample can be regenerated or reshaped. |

## How to use it

1. **Try the Excel dashboard first.** Open
   `excel/production_dashboard_sample.xlsx` — the KPIs and charts are live
   formulas over the `Data` sheet, so edit a number and watch the dashboard
   move.
2. **Recreate the Form** from `google-form-fields.md` in Google Forms,
   linked to a new Google Sheet.
3. **Add the script.** In the Sheet, Extensions > Apps Script, paste
   `Code.gs`, and set the `onFormSubmit` installable trigger on form submit.
   Follow `sheets-setup.md` for the tab layout, validation and manager view.
4. **Connect Excel.** Follow the steps on the `How to refresh` sheet in the
   workbook to point Power Query at the published CSV of the summary tab.
5. **Test end to end.** Submit the Form from a phone, check the row appears
   cleaned and flagged (or not) in the Sheet, then refresh the Excel file.

## KPIs in the dashboard

- Daily production totals (units by line and by day)
- Material consumption vs. receipts, and stock on hand per material
- Defect rate overall and by production line
- Downtime minutes by day
- Variance flags: issued quantity above stock on hand, negative quantities,
  missing supplier or batch

## Honest notes and limits

- **All data is synthetic.** Names such as "Supplier Alpha" and products
  such as "Detergent Powder 1kg" are invented placeholders. The totals
  (319,755 units over 14 days, 0.84% defect rate) describe the sample file
  only.
- The Excel workbook uses formulas and pivots over the embedded sample
  data. The live Power Query link to a Google Sheet is documented in
  `How to refresh` but not pre-wired — it needs the client's own Sheet URL.
- The Apps Script is written for the column layout in
  `data/sample_production_records.csv` and has not been run against a live
  Form in this repo; test it with a few real submissions before relying on it.
- No client names, testimonials or results are claimed anywhere in this repo.

## Skills demonstrated

Google Forms, Google Sheets (data validation, ARRAYFORMULA/QUERY, filter
views, conditional formatting), Google Apps Script, Excel (Power Query,
pivot-style formulas, charts, printable dashboard), and production KPI design.
