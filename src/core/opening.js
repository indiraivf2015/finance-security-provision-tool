// File 2: last month's statement, which gives this month's opening balances. Copied from the approved Security
// engine; the columns to carry come from the provision's settings.
//
// spec = {
//   columns: [ { key: 'ytd', header: 'YTD FY27', choosable: true },     header matched exactly (after normalising);
//              { key: 'cf',  header: 'FY26 Pro CF' } ],                 choosable = if it appears twice, the user picks one
//   aprilNote: true,                                                     in April, explain the SOP has no year-end rule
// }
// Returns { centres: Map(code -> { code, name, row, v: {key: amount}, blank: {key: bool} }), order, … } or { needChoice }.
'use strict';
const { InputError, norm, parseMonthYear, prevMonth, monthLabel, colLetter, MONTH_ABBR, pad2 } = require('./common');
const { infoText, infoNumber, infoInteger, sheetRows } = require('./excel');

const APRIL = 'This is April — the start of a new financial year. The SOP does not describe the year-end roll-over, so the tool cannot proceed without finance\'s rule.';

function readOpening(wb, period, opts, spec) {
  const cols = spec.columns.map((c) => ({ ...c, k: norm(c.header) }));
  const headerList = cols.map((c) => `"${c.header}"`).join(' and ');
  const found = [];
  for (const ws of wb.worksheets) {
    const g = sheetRows(ws);
    for (let r = 1; r <= Math.min(30, g.rowCount); r++) {
      const keys = []; for (let c = 1; c <= g.colCount; c++) keys[c] = norm(infoText(g.rows[r][c]));
      const cc = keys.indexOf('costcenter') > 0 ? keys.indexOf('costcenter') : keys.indexOf('costcentre');
      if (cc < 1) continue;
      const at = {}; let any = false;
      for (const col of cols) { at[col.key] = []; keys.forEach((k, c) => { if (k === col.k) at[col.key].push(c); }); if (at[col.key].length) any = true; }
      if (!any) continue;
      let nameCol = keys.indexOf('costcentername'); if (nameCol < 1) nameCol = keys.indexOf('costcentrename');
      found.push({ g, hdr: r, cc, nameCol, at });
      break;
    }
  }
  const aprilNote = spec.aprilNote && period.m === 4 ? ' ' + APRIL : '';
  if (!found.length) {
    throw new InputError([`Last month's statement: no sheet has the headings "Cost Center" and ${headerList}. ` +
      `The SOP carries ${headerList} from last month.` + (aprilNote || ' Check that the right file was uploaded.')]);
  }
  let pickSheet = found;
  if (found.length > 1) {
    if (!opts.openingSheet) return { needChoice: { kind: 'openingSheet', options: found.map((f) => f.g.name) } };
    pickSheet = found.filter((f) => f.g.name === opts.openingSheet);
    if (!pickSheet.length) throw new InputError([`Last month's statement: sheet "${opts.openingSheet}" not found.`]);
  }
  const s = pickSheet[0];
  const errs = [];
  for (const col of cols) {
    const n = s.at[col.key].length;
    if (n === 0) errs.push(`Last month's statement (sheet "${s.g.name}"): no "${col.header}" column.` + aprilNote);
    else if (n > 1 && !col.choosable) errs.push(`Last month's statement (sheet "${s.g.name}"): expected one "${col.header}" column, found ${n}.`);
  }
  if (errs.length) throw new InputError(errs);

  // Which month is this statement for?  Must be the month before the ERP period.
  const pm = prevMonth(period.y, period.m);
  let stmtMonth = null, stmtTitle = '';
  for (let r = 1; r < s.hdr && !stmtMonth; r++) for (let c = 1; c <= s.g.colCount && !stmtMonth; c++) {
    const t = infoText(s.g.rows[r][c]);
    const mt = t.match(/month\s+of\s+([A-Za-z]{3,9})[\s\-'’.]*(\d{2,4})\b/i);
    if (mt) { const p = parseMonthYear(mt[1] + '-' + mt[2]); if (p) { stmtMonth = p; stmtTitle = t.trim(); } }
  }
  if (stmtMonth && (stmtMonth.y !== pm.y || stmtMonth.m !== pm.m)) {
    throw new InputError([`Last month's statement is titled "${stmtTitle}", but the ERP extract is for ${period.label}. Upload the statement for ${MONTH_ABBR[pm.m]}-${pad2(pm.y % 100)}.`]);
  }
  if (!stmtMonth && !opts.confirmOpeningMonth) return { needChoice: { kind: 'confirmOpeningMonth', month: `${MONTH_ABBR[pm.m]}-${pad2(pm.y % 100)}`, sheet: s.g.name } };

  // A choosable column that appears more than once: the user says which one is the closing balance
  const totalsOf = (c) => { let t = 0; for (let r = s.hdr + 1; r <= s.g.rowCount; r++) { const n = infoNumber(s.g.rows[r][c]); if (n.ok && n.v !== null && infoInteger(s.g.rows[r][s.cc]).v != null) t += n.v; } return t; };
  const use = {};
  for (const col of cols) {
    use[col.key] = s.at[col.key][0];
    if (s.at[col.key].length > 1) {
      if (!opts.openingYtdCol) return { needChoice: { kind: 'openingYtdCol', header: col.header, sheet: s.g.name, options: s.at[col.key].map((c) => ({ col: colLetter(c), total: totalsOf(c), header: infoText(s.g.rows[s.hdr][c]) })) } };
      const want = s.at[col.key].find((c) => colLetter(c) === opts.openingYtdCol);
      if (!want) throw new InputError([`Last month's statement: column ${opts.openingYtdCol} is not a "${col.header}" column.`]);
      use[col.key] = want;
    }
  }

  const centres = new Map(); const order = [];
  for (let r = s.hdr + 1; r <= s.g.rowCount; r++) {
    const row = s.g.rows[r];
    const code = infoInteger(row[s.cc]);
    const nums = {}; for (const col of cols) nums[col.key] = infoNumber(row[use[col.key]]);
    if (code.ok && code.v === null) {
      if (cols.some((col) => (nums[col.key].ok && nums[col.key].v) || !nums[col.key].ok)) errs.push(`Last month's statement row ${r}: there is an amount but no cost centre code.`);
      continue;
    }
    if (!code.ok) { errs.push(`Last month's statement row ${r}: Cost Center ${code.why}.`); continue; }
    for (const col of cols) if (!nums[col.key].ok) errs.push(`Last month's statement row ${r} (cost centre ${code.v}): ${col.header} — ${nums[col.key].why}.`);
    if (centres.has(code.v)) { errs.push(`Last month's statement: cost centre ${code.v} appears more than once (rows ${centres.get(code.v).row} and ${r}).`); continue; }
    const v = {}, blank = {};
    for (const col of cols) { v[col.key] = nums[col.key].ok ? (nums[col.key].v || 0) : 0; blank[col.key] = nums[col.key].ok && nums[col.key].v === null; }
    centres.set(code.v, { code: code.v, name: s.nameCol > 0 ? infoText(row[s.nameCol]) : '', row: r, v, blank });
    order.push(code.v);
  }
  if (errs.length) throw new InputError(errs.slice(0, 50));
  if (!centres.size) throw new InputError(['Last month\'s statement: no cost-centre rows found under the headings.']);
  const colLetters = {}; for (const col of cols) colLetters[col.key] = colLetter(use[col.key]);
  return { sheet: s.g.name, headerRow: s.hdr, colLetters, title: stmtTitle, monthConfirmedByUser: !stmtMonth, centres, order, columns: cols, label: monthLabel(pm.y, pm.m) };
}

module.exports = { readOpening };
