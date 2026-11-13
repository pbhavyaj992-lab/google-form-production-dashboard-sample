# Google Form design — Batch Manufacturing Record (sample)

Build this Form in Google Forms and link it to a new Google Sheet. Keep it
short enough to complete on a phone in under two minutes: supervisors fill
it in three times a day, so every extra question is a reason to skip it.

Design rule used throughout: **one material movement per submission**,
linked by batch number. A batch usually consumes several materials, so one
long form per batch would be slow and error-prone on a phone. The supervisor
submits the form once per movement (inward receipt or outward issue).

## Section 1 — Shift and batch

| # | Question | Type | Required | Notes |
| --- | --- | --- | --- | --- |
| 1 | Date | Date | Yes | Defaults to today |
| 2 | Shift | Multiple choice | Yes | Morning (06:00–14:00), Afternoon (14:00–22:00), Night (22:00–06:00) |
| 3 | Supervisor name | Dropdown | Yes | Supervisor 1, Supervisor 2, Supervisor 3 |
| 4 | Production line | Dropdown | Yes | Line A, Line B, Line C |
| 5 | Product | Dropdown | Yes | Options depend on the line; in a plain Form, list all six products and validate in the Sheet |
| 6 | Batch number | Short answer | Yes | Format `BYYMMDD-Ln`, e.g. `B261026-A1`. Response validation: text matches pattern `B[0-9]{6}-[A-C][1-3]` |
| 7 | Operator | Short answer | No | Free text, for traceability |

## Section 2 — Material movement

| # | Question | Type | Required | Notes |
| --- | --- | --- | --- | --- |
| 8 | Material | Dropdown | Yes | Caustic Soda, Soda Ash, SLES, Fragrance, Packaging Film, HDPE Bottles |
| 9 | Supplier | Dropdown | Yes | Supplier Alpha, Supplier Beta, Supplier Gamma |
| 10 | Movement type | Multiple choice | Yes | Inward (receipt) or Outward (issue to production) |
| 11 | Quantity received (kg) | Short answer, number | Only if Inward | Response validation: number ≥ 0. Use "Go to section based on answer" on Q10 so this only appears for Inward |
| 12 | Quantity issued (kg) | Short answer, number | Only if Outward | Response validation: number ≥ 0. Same branching as Q11 |
| 13 | Units produced | Short answer, number | Yes | Number ≥ 0. Asked once per shift/line; repeated on both movement rows of a pair in the sample data, counted once in the dashboard |
| 14 | Defective units | Short answer, number | Yes | Number ≥ 0, must be ≤ units produced (checked in Apps Script — Forms cannot compare two answers) |
| 15 | Downtime (minutes) | Short answer, number | Yes | Number ≥ 0 |

## What the Form deliberately does NOT ask

- **Remaining stock.** Calculated in the Sheet (`total received − total
  issued`), never typed. Typed stock figures drift and corrupt every
  variance report downstream.
- **Anything the supervisor could look up wrong under time pressure**, such
  as a unit price or a supplier code. Keep it to what they can see on the
  shop floor.

## Form settings

- Limit to one response per person: off (a supervisor submits several times
  a day). Instead, collect email addresses so duplicate or mistaken
  submissions can be traced.
- Confirmation message: "Recorded. Check the Sheet if you need to correct a
  quantity — do not submit the same movement twice."
- Supervisors get the Form link only. They are **not** given access to the
  Sheet, so nobody can edit rows directly.
- Test on a phone before go-live: submit one Inward and one Outward
  movement and confirm both rows arrive cleaned in `Clean_Data`.
