/**
 * Google Form-Linked Production Dashboard — SAMPLE Apps Script.
 *
 * What it does:
 *  - Runs on every Google Form submission (installable onFormSubmit trigger).
 *  - Validates the new response row (quantities >= 0, issued <= stock on hand,
 *    required fields present) and writes a cleaned row to the `Clean_Data` tab.
 *  - Calculates stock on hand per material (total received - total issued)
 *    instead of trusting a typed "remaining stock" value.
 *  - Flags inconsistent rows in an `Alerts` tab and colours the flag cell
 *    in `Clean_Data` so the Plant Manager sees problems at a glance.
 *
 * Setup:
 *  1. Create the Form from google-form-fields.md and link it to this Sheet.
 *  2. Create tabs named: Form Responses 1 (created by the Form),
 *     Clean_Data, Summary, Alerts. See sheets-setup.md for the layout.
 *  3. Paste this file in Extensions > Apps Script, then set an installable
 *     trigger: onFormSubmit, event source "From spreadsheet", event "On form
 *     submit".
 *  4. Run `rebuildSummary()` once by hand to populate the Summary tab.
 *
 * All data this script handles in the sample repo is synthetic sample data.
 */

// Column order of the Clean_Data tab (matches data/sample_production_records.csv).
const CLEAN_HEADERS = [
  'date', 'shift', 'supervisor', 'production_line', 'product', 'batch_no',
  'material', 'supplier', 'movement_type', 'qty_received_kg',
  'qty_issued_kg', 'units_produced', 'defects', 'downtime_minutes',
  'operator', 'stock_on_hand_kg', 'flag'
];

const MATERIALS = ['Caustic Soda', 'Soda Ash', 'SLES', 'Fragrance', 'Packaging Film', 'HDPE Bottles'];
const SUPPLIERS = ['Supplier Alpha', 'Supplier Beta', 'Supplier Gamma'];

/** Installable trigger: runs on every Form submission. */
function onFormSubmit(e) {
  const values = e.namedValues; // keys are the Form question titles
  const row = {
    date: first(values['Date']),
    shift: first(values['Shift']),
    supervisor: first(values['Supervisor name']),
    production_line: first(values['Production line']),
    product: first(values['Product']),
    batch_no: first(values['Batch number']),
    material: first(values['Material']),
    supplier: first(values['Supplier']),
    movement_type: first(values['Movement type']),
    qty_received_kg: num(values['Quantity received (kg)']),
    qty_issued_kg: num(values['Quantity issued (kg)']),
    units_produced: num(values['Units produced']),
    defects: num(values['Defective units']),
    downtime_minutes: num(values['Downtime (minutes)']),
    operator: first(values['Operator']),
  };

  const flags = validateRow(row);
  const stock = stockOnHand(row.material) +
    (row.movement_type === 'Inward' ? row.qty_received_kg : 0) -
    (row.movement_type === 'Outward' ? row.qty_issued_kg : 0);
  if (row.movement_type === 'Outward' && row.qty_issued_kg > stockOnHand(row.material)) {
    flags.push('ISSUE EXCEEDS STOCK ON HAND');
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName('Clean_Data');
  ensureHeaders(sheet);
  const out = CLEAN_HEADERS.slice(0, 15).map(h => row[h]);
  out.push(stock, flags.length ? flags.join('; ') : 'OK');
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, out.length).setValues([out]);
  if (flags.length) {
    sheet.getRange(newRow, CLEAN_HEADERS.length)
      .setBackground('#FFC7CE')
      .setFontColor('#9C0006');
    logAlert(row, flags);
  }
}

/** Validation rules for one submission. Returns an array of flag strings (empty = clean). */
function validateRow(r) {
  const flags = [];
  if (!r.date) flags.push('MISSING DATE');
  if (!r.batch_no) flags.push('MISSING BATCH NUMBER');
  if (!r.material || MATERIALS.indexOf(r.material) === -1) flags.push('UNKNOWN MATERIAL');
  if (!r.supplier || SUPPLIERS.indexOf(r.supplier) === -1) flags.push('UNKNOWN SUPPLIER');
  if (r.movement_type === 'Inward' && r.qty_received_kg <= 0) flags.push('INWARD WITH NO QUANTITY RECEIVED');
  if (r.movement_type === 'Outward' && r.qty_issued_kg <= 0) flags.push('OUTWARD WITH NO QUANTITY ISSUED');
  if (r.qty_received_kg < 0 || r.qty_issued_kg < 0) flags.push('NEGATIVE QUANTITY');
  if (r.units_produced < 0 || r.defects < 0 || r.downtime_minutes < 0) flags.push('NEGATIVE PRODUCTION FIGURE');
  if (r.defects > r.units_produced) flags.push('DEFECTS EXCEED UNITS PRODUCED');
  return flags;
}

/** Current stock on hand for a material, from the Clean_Data tab. */
function stockOnHand(material) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Clean_Data');
  const data = sheet.getDataRange().getValues();
  let received = 0, issued = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] !== material) continue; // column G = material
    received += Number(data[i][9]) || 0;   // column J = qty received
    issued += Number(data[i][10]) || 0;    // column K = qty issued
  }
  return received - issued;
}

/** Write the live summary formulas into the Summary tab. Run once by hand, and again if the layout changes. */
function rebuildSummary() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  s.clear();
  s.getRange('A1').setValue('Production summary — live formulas over Clean_Data');
  s.getRange('A3:D3').setValues([['Production line', 'Units produced', 'Defects', 'Defect rate']]);
  const lines = ['Line A', 'Line B', 'Line C'];
  lines.forEach((line, i) => {
    const r = 4 + i;
    s.getRange(r, 1).setValue(line);
    s.getRange(r, 2).setFormula(
      `=SUMIFS(Clean_Data!L:L,Clean_Data!D:D,A${r},Clean_Data!I:I,"Inward")`);
    s.getRange(r, 3).setFormula(
      `=SUMIFS(Clean_Data!M:M,Clean_Data!D:D,A${r},Clean_Data!I:I,"Inward")`);
    s.getRange(r, 4).setFormula(`=IF(B${r}=0,"-",C${r}/B${r})`).setNumberFormat('0.00%');
  });
  const t = 4 + lines.length;
  s.getRange(t, 1).setValue('Total');
  s.getRange(t, 2).setFormula(`=SUM(B4:B${t - 1})`);
  s.getRange(t, 3).setFormula(`=SUM(C4:C${t - 1})`);
  s.getRange(t, 4).setFormula(`=IF(B${t}=0,"-",C${t}/B${t})`).setNumberFormat('0.00%');
}

/** Append a flag to the Alerts tab for the Plant Manager to review. */
function logAlert(row, flags) {
  const ss = SpreadsheetApp.getActive();
  const a = ss.getSheetByName('Alerts') || ss.insertSheet('Alerts');
  if (a.getLastRow() === 0) {
    a.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Date', 'Batch', 'Material', 'Flag']]);
  }
  a.appendRow([new Date(), row.date, row.batch_no, row.material, flags.join('; ')]);
}

/** Add dropdown validation for material and supplier on the Clean_Data tab. */
function setupValidation() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Clean_Data');
  const matRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MATERIALS, true).setAllowInvalid(false).build();
  const supRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SUPPLIERS, true).setAllowInvalid(false).build();
  sheet.getRange('G2:G1000').setDataValidation(matRule);
  sheet.getRange('H2:H1000').setDataValidation(supRule);
}

function first(v) { return v ? String(v[0]).trim() : ''; }
function num(v) { const n = Number(v ? v[0] : 0); return isNaN(n) ? 0 : n; }
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CLEAN_HEADERS.length).setValues([CLEAN_HEADERS]);
    sheet.setFrozenRows(1);
  }
}
