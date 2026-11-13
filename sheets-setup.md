# Google Sheet setup — manager view, validation and KPIs (sample)

## Tabs

| Tab | Who sees it | Purpose |
| --- | --- | --- |
| `Form Responses 1` | Owner only | Raw Form responses, auto-created. Locked: nobody edits this tab by hand. |
| `Clean_Data` | Owner, Plant Manager (view) | One cleaned row per submission, written by `Code.gs`. Columns match `data/sample_production_records.csv`, plus `stock_on_hand_kg` and `flag`. |
| `Summary` | Owner, Plant Manager (view) | Live KPI tables built by `rebuildSummary()` in `Code.gs`: units by line, defect rate, received vs. issued, daily trend. |
| `Manager_View` | Plant Manager | Read-only filter view over `Clean_Data` (see below). |
| `Alerts` | Owner, Plant Manager | One row per flagged submission, written by `Code.gs`. Review daily, correct the source entry, delete the alert. |

## Data validation (Data > Data validation)

- `Clean_Data` column G (material): dropdown from the six materials in
  `Code.gs` (`setupValidation()` applies this).
- `Clean_Data` column H (supplier): dropdown from the three suppliers.
- `Clean_Data` column I (movement type): `Inward` or `Outward` only.
- Quantity columns (J, K) and production columns (L, M, N): number ≥ 0.
- Column A (date): valid date, not in the future.

## Conditional formatting (Format > Conditional formatting, on `Clean_Data`)

| Rule | Format |
| --- | --- |
| `flag` column is not `OK` | Light red fill, dark red text |
| `qty_issued_kg` > running stock on hand for that material | Light red fill on the quantity cell |
| `defects` > `units_produced` | Light red fill on the defects cell |
| Any required cell blank (date, batch, material, supplier) | Yellow fill |

## Manager view

Data > Filter views > Create new filter view, named `Manager_View`, over
`Clean_Data`. Share the Sheet with the Plant Manager as **Viewer** and send
them the filter-view link, so they can filter by date, shift, batch and
material without touching formulas or other people's filters. Filter views
are per-user: the manager's filters never disturb the owner's view.

## Key formulas (all live — no manual refresh)

Stock on hand per material (used by `Code.gs` and in `Summary`):

```
=SUMIFS(Clean_Data!J:J, Clean_Data!G:G, "Caustic Soda")
 - SUMIFS(Clean_Data!K:K, Clean_Data!G:G, "Caustic Soda")
```

Units produced by line (Inward rows only — each shift/line logs one Inward
and one Outward row carrying the same production figures, so counting both
would double-count):

```
=SUMIFS(Clean_Data!L:L, Clean_Data!D:D, "Line A", Clean_Data!I:I, "Inward")
```

Defect rate by line:

```
=SUMIFS(Clean_Data!M:M, Clean_Data!D:D, "Line A", Clean_Data!I:I, "Inward")
 / SUMIFS(Clean_Data!L:L, Clean_Data!D:D, "Line A", Clean_Data!I:I, "Inward")
```

Daily production trend, one row per date:

```
=SUMIFS(Clean_Data!L:L, Clean_Data!A:A, $A5, Clean_Data!I:I, "Inward")
```

## Excel dashboard link (Power Query)

1. File > Share > Publish to web > `Summary` tab > CSV > Publish.
2. In Excel: Data > Get Data > From Web, paste the published CSV link.
3. Close & Load To the `Data` sheet of
   `excel/production_dashboard_sample.xlsx`, replacing the sample rows.
4. Set the query to refresh on open. Opening the workbook then pulls the
   latest Form responses in under 30 seconds.

Note: the published CSV link is readable by anyone who has it. Publish only
the `Summary` tab, never `Form Responses 1`, and share the Sheet itself only
with the Plant Manager and the Owner.
