// File 3: Rajeev's provision file. Finds the provision's sheet by name and its amount column by heading
// (never by column letter: the real file's columns differ from the SOP figures). Generalised from the approved
// Security engine's reader.
//
// spec = {
//   sheet: { test: /^\s*hk(?![a-z])/i, text: '"HK <Mon-YY>"' },    which sheet holds this provision
//   amount: { startsWith: 'manpower', text: 'Manpower' },         amount column: first heading cell that starts with this
//   noun: 'housekeeping',                                          used in messages
// }
// Returns { sheet, codeColLetter, amtColLetter, firstRow, lines: [{row, code, name, vendor, sno, state, amount, blank}],
//           blanks, columnTotal } or { needChoice }.
'use strict';
const { InputError, norm, parseMonthYear, colLetter, MON } = require('./common');
const { infoText, infoNumber, infoInteger, sheetRows } = require('./excel');

function readProvisionFile(wb, period, opts, spec) {
  let sheets = wb.worksheets.filter((ws) => spec.sheet.test.test(ws.name));
  if (sheets.length === 0 && wb.worksheets.length === 1) sheets = wb.worksheets.slice();
  if (sheets.length !== 1) {
    if (!opts.rajeevSheet) return { needChoice: { kind: 'rajeevSheet', noun: spec.noun, options: wb.worksheets.map((w) => w.name) } };
    sheets = wb.worksheets.filter((w) => w.name === opts.rajeevSheet);
    if (!sheets.length) throw new InputError([`Provision file: sheet "${opts.rajeevSheet}" not found.`]);
  }
  const ws = sheets[0];
  // The sheet name, if it carries a month, must be the ERP month
  const nm = ws.name.match(new RegExp('(?<![a-z])' + MON + '(?![a-z])[\\s\\-\'’.]*(\\d{2}|\\d{4})(?!\\d)', 'i'));
  if (nm) {
    const p = parseMonthYear(nm[1] + '-' + nm[2]);
    if (p && (p.y !== period.y || p.m !== period.m)) throw new InputError([`Provision file: the sheet used is "${ws.name}", but the ERP extract is for ${period.label}. Upload this month's provision file.`]);
  }
  const g = sheetRows(ws);
  let codeCol = -1, codeRow = -1, nameCol = -1, snoCol = -1, vendorCol = -1, stateCol = -1;
  const amtHits = [];
  for (let r = 1; r <= Math.min(20, g.rowCount) && codeCol < 0; r++) for (let c = 1; c <= g.colCount; c++) {
    if (norm(infoText(g.rows[r][c])) === 'storagecode') { codeCol = c; codeRow = r; break; }
  }
  // The heading block is every row above the first row with a numeric Storage Code (headings can sit in merged
  // rows below "Storage Code"). Headings are looked for only there, never in data rows.
  let headEnd = Math.min(20, g.rowCount);
  if (codeCol > 0) for (let r = codeRow + 1; r <= Math.min(codeRow + 20, g.rowCount); r++) { const n = infoInteger(g.rows[r][codeCol]); if (n.ok && n.v !== null) { headEnd = r - 1; break; } }
  for (let r = 1; r <= headEnd; r++) for (let c = 1; c <= g.colCount; c++) {
    const k = norm(infoText(g.rows[r][c]));
    if (!k) continue;
    if (k.startsWith(spec.amount.startsWith)) amtHits.push({ c, r, text: infoText(g.rows[r][c]) });
    if (k === 'centername' && nameCol < 0) nameCol = c;
    if ((k === 'sno' || k === 'srno') && snoCol < 0) snoCol = c;
    if (k === 'vendorname' && vendorCol < 0) vendorCol = c;
    if (k === 'state' && stateCol < 0) stateCol = c;
  }
  const errs = [];
  if (codeCol < 0) errs.push(`Provision file (sheet "${ws.name}"): no "Storage Code" heading found.` + (spec.sheet.text
    ? ` Upload Rajeev's combined provision file, the one with the ${spec.sheet.text} sheet and a "Storage Code" column, not a separate copy of the ${spec.noun} figures.` : ''));
  const amtCols = [...new Set(amtHits.map((h) => h.c))];
  if (amtCols.length === 0) errs.push(`Provision file (sheet "${ws.name}"): no "${spec.amount.text}" heading found.`);
  if (amtCols.length > 1) errs.push(`Provision file (sheet "${ws.name}"): more than one column is headed "${spec.amount.text}" (columns ${amtCols.map(colLetter).join(', ')}).`);
  if (errs.length) throw new InputError(errs);
  const amtCol = amtCols[0]; const amtRow = amtHits[0].r;
  const first = Math.max(codeRow, amtRow) + 1;
  // A data row has a Storage Code or an amount; a text sub-heading directly under the headings is skipped.
  let columnTotal = 0;
  const lines = []; const blanks = [];
  for (let r = first; r <= g.rowCount; r++) {
    const row = g.rows[r];
    const code = infoInteger(row[codeCol]);
    const amt = infoNumber(row[amtCol]);
    if (amt.ok && amt.v !== null) columnTotal += amt.v;
    if (code.ok && code.v === null && amt.ok && amt.v === null) continue;
    if (r === first && code.ok && code.v === null && row[amtCol].k === 'text') continue;
    if (code.ok && code.v === null) { errs.push(`Provision file row ${r}: there is a ${spec.amount.text} amount but no Storage Code.`); continue; }
    if (!code.ok) { errs.push(`Provision file row ${r}: Storage Code ${code.why}.`); continue; }
    if (!amt.ok) { errs.push(`Provision file row ${r} (Storage Code ${code.v}): ${spec.amount.text} — ${amt.why}.`); continue; }
    if (amt.v === null) blanks.push({ row: r, code: code.v });
    const sno = snoCol > 0 ? infoNumber(row[snoCol]) : { ok: false };
    lines.push({
      row: r, code: code.v, amount: amt.v === null ? 0 : amt.v, blank: amt.v === null,
      name: nameCol > 0 ? infoText(row[nameCol]) : '', vendor: vendorCol > 0 ? infoText(row[vendorCol]) : '', state: stateCol > 0 ? infoText(row[stateCol]) : '',
      sno: sno.ok && sno.v !== null ? sno.v : (snoCol > 0 ? infoText(row[snoCol]) : ''),
    });
  }
  if (errs.length) throw new InputError(errs.slice(0, 50));
  if (!lines.length) throw new InputError([`Provision file (sheet "${ws.name}"): no rows with a Storage Code were found.`]);
  return { sheet: ws.name, codeColLetter: colLetter(codeCol), amtColLetter: colLetter(amtCol), amtHeading: amtHits[0].text, firstRow: first, lines, blanks, columnTotal };
}

module.exports = { readProvisionFile };
