// Housekeeping Material provision: SOP "Monthly Housekeeping Material Provision Working" v1.0 (Oct 2026), Steps 1-15.
// A simpler statement than Security / Housekeeping Expenses: one booking column, one carry-forward, no GST conversion.
//
// SOP step -> code
//   Step 2      ERP account 51611 only ("GL Code From" and "GL Code To" both 51611)
//   Steps 5-9   as Security (src/core/clean.js). Step 7 keeps "IIVF P2P" and "Payables" (owner's decision, Oct 2026,
//               same wording and treatment as Security). Finance rule: year not in the narration -> current FY.
//   Step 10     pivot rows SEGMENT1 + BALANCING_SEGMENT_DESC, Exp Year as a filter on the current FY: only current-FY
//               bookings reach the statement. Earlier-FY lines stay in Working and are listed in Points to check (owner).
//   Step 11     Rajeev's sheet "HK <Mon-YY>", Materials column only, taken AS IS (no ÷118, no ×1.135).
//               Rows with the same Storage Code (e.g. "Main" + "store boy") are added together.
//   Steps 12-15 [Prev]-Provision = last month's "YTD"; [Month]-Booking; Prov CF = Provision − Booking;
//               [Month]-Provision = Materials; YTD = Prov CF + [Month]-Provision.
//   Amounts: exact values kept; shown without paise (owner's decision, Oct 2026).
//   April: stops (the SOP has no year-end roll-over rule; owner's decision, Oct 2026).
'use strict';
const { createEngine } = require('../../core/engine');
const { EPS, fyLabelWorking, toPaise, fmtINR, MONTH_ABBR } = require('../../core/common');
const { ExcelJS, formatTable, FMT_AMT, HEAD_FILL, BOX } = require('../../core/excel');
const { applyNegativeCarryForward, checkList } = require('../../core/rules');
const { isWrongKind, GST_RECLASS_CREATOR } = require('../../core/clean');
const { writeRaw, writeWorking, writePoints } = require('../../core/sheets');

const VERSION = '1.0.0';
const TOOL = 'Housekeeping Material Provision Tool';

// Negative Prov CF. Owner's decision (Oct 2026): apply finance's negative carry-forward rule, as for Security and
// Housekeeping Expenses: set to 0, taken from the 10 centres with the highest Prov CF, in proportion; column total unchanged.
// This departs from SOP v1.0 Step 14 ("carry it with its minus sign"); finance to update the Material SOP to match.
// null here would mean the SOP as written. Do not change without the owner's decision.
const NEGATIVE_CF_RULE = { topN: 10 };

const monthWord = (y, m) => (m === 9 ? 'Sept' : MONTH_ABBR[m]); // the SOP's headers: "Aug-Provision", "Sept-Booking"
const NEG_COL = [{ sopF: 'cfSop', adjF: 'adj', finF: 'cf', label: 'Prov CF' }];

function labels(period) {
  const pm = period.m === 1 ? { y: period.y - 1, m: 12 } : { y: period.y, m: period.m - 1 };
  const M = monthWord(period.y, period.m), P = monthWord(pm.y, pm.m);
  return { open: `${P}-Provision`, booking: `${M}-Booking`, cf: 'Prov CF', prov: `${M}-Provision`, ytd: 'YTD' };
}

function compute(state, rule) {
  const { period, kept, removed, opening, rajeev, extract } = state;
  const H = labels(period);
  const cur = kept.filter((r) => r.fy === period.cur), earlier = kept.filter((r) => r.fy !== period.cur);

  // Step 10: pivot on current-FY lines (Exp Year filter), by SEGMENT1 + BALANCING_SEGMENT_DESC, in paise
  const pivot = new Map(); const byCode = new Map();
  for (const r of cur) {
    const key = r.seg1 + '\u0000' + r.ccNameRaw.toLowerCase();
    if (!pivot.has(key)) pivot.set(key, { code: r.seg1, name: r.ccNameRaw, amt: 0 });
    pivot.get(key).amt += toPaise(r.bal);
    if (!byCode.has(r.seg1)) byCode.set(r.seg1, { amt: 0, name: r.ccName });
    byCode.get(r.seg1).amt += toPaise(r.bal);
  }
  // Step 11: Materials as is; one statement row per centre
  const provByCode = new Map(); const provCount = new Map();
  for (const l of rajeev.lines) { provByCode.set(l.code, (provByCode.get(l.code) || 0) + l.amount); provCount.set(l.code, (provCount.get(l.code) || 0) + 1); }

  const booked = [...byCode.keys()].sort((a, b) => a - b);
  const codes = booked.concat(opening.order.filter((c) => !byCode.has(c)));
  codes.push(...[...provByCode.keys()].filter((c) => !opening.centres.has(c) && !byCode.has(c)).sort((a, b) => a - b));
  const provName = new Map(); for (const l of rajeev.lines) if (!provName.has(l.code) && l.name) provName.set(l.code, l.name);
  const stmt = codes.map((code) => {
    const o = opening.centres.get(code);
    const open = o ? o.v.ytd : 0;                                       // Step 12
    const booking = byCode.has(code) ? byCode.get(code).amt / 100 : 0;  // Step 13
    const cf = open - booking;                                          // Step 14
    const prov = provByCode.get(code) || 0;                             // Step 15
    const name = (o && o.name) || (byCode.get(code) && byCode.get(code).name) || provName.get(code) || '';
    return { code, name, open, booking, cfSop: cf, adj: 0, cf, prov, ytd: 0, isNew: !o };
  });
  const adjInfo = rule ? applyNegativeCarryForward(stmt, NEG_COL, rule.topN) : null;
  for (const s of stmt) { s.cf = s.cfSop + s.adj; s.ytd = s.cf + s.prov; }

  const sum = (f) => stmt.reduce((a, s) => a + s[f], 0);
  const T = { open: sum('open'), booking: sum('booking'), cf: sum('cf'), prov: sum('prov'), ytd: sum('ytd'), adj: sum('adj') };
  const keptTotalP = kept.reduce((a, r) => a + toPaise(r.bal), 0);
  const curTotalP = cur.reduce((a, r) => a + toPaise(r.bal), 0);
  const removedTotalP = removed.reduce((a, r) => a + toPaise(r.bal), 0);
  const extractTotalP = extract.rows.reduce((a, r) => a + toPaise(r.bal), 0);
  let pivotCurP = 0; for (const p of pivot.values()) pivotCurP += p.amt;
  let openYtd = 0; for (const o of opening.centres.values()) openYtd += o.v.ytd;

  // Section 04 checks + completeness
  const { checks, add, status } = checkList();
  add('04 · check 1', 'Every Working line has an Exp Year tag', 0, kept.filter((r) => r.fy !== period.cur && r.fy !== period.prev).length, 'count0fail');
  const noVendor = kept.filter((r) => !r.vendor).length;
  add('04 · check 1', 'Every Working line has a vendor name in Attribute 3 (lines with no Attribute 3 and no Party Name in the ERP cannot be filled)', 0, noVendor, 'count0');
  add('04 · check 2', 'Working contains only P2P bookings and GST reclass entries by ' + GST_RECLASS_CREATOR, 0, kept.filter(isWrongKind).length, 'count0fail');
  add('04 · check 3', `Pivot grand total = total BALANCE of the Working sheet (${fyLabelWorking(period.cur)} lines; earlier-year lines are filtered out)`, curTotalP / 100, pivotCurP / 100);
  add('04 · check 3', `${H.booking} column total = pivot grand total`, pivotCurP / 100, T.booking);
  add('04 · check 4', `${H.prov} column total = Materials column total in Rajeev's file`, rajeev.columnTotal, T.prov);
  add('04 · check 5', `Totals row: ${H.open} − ${H.booking} = Prov CF`, T.open - T.booking, rule ? T.cf - T.adj : T.cf);
  add('04 · check 5', `Totals row: Prov CF + ${H.prov} = YTD`, T.cf + T.prov, T.ytd);
  add('Completeness', 'ERP extract total = Working total + Removed lines total (no line lost)', extractTotalP / 100, (keptTotalP + removedTotalP) / 100);
  add('Completeness · Step 12', `Opening ${H.open} = last month's YTD total`, openYtd, T.open);
  if (adjInfo) {
    add('Negative CF rule', 'Prov CF adjustments net to zero (column total unchanged)', 0, T.adj);
    add('Negative CF rule', 'No negative Prov CF left after the adjustment', 0, stmt.filter((s) => s.cf < -EPS).length, 'count0fail');
  }

  // Points to check
  const points = [];
  const pt = (code, name, point, amount, check) => points.push({ code, name, point, amount: amount == null ? null : Math.round(amount * 100) / 100, check });
  for (const r of earlier) pt(r.seg1, r.ccName, `${fyLabelWorking(r.fy)} line left out of the statement (only ${fyLabelWorking(period.cur)} bookings are used, SOP Step 10). ERP row ${r.erpRow}: "${r.desc.slice(0, 120)}"`, r.bal, 'For information. Confirm this line belongs to last year.');
  for (const s of stmt) {
    if (!provByCode.has(s.code)) pt(s.code, s.name, `Not in Rajeev's provision file, so "${H.prov}" is 0.` + (Math.abs(s.booking) > EPS ? ` ${fmtINR(s.booking)} was booked for this centre this month.` : ''), Math.abs(s.booking) > EPS ? s.booking : null, 'Check with Rajeev whether this centre should have a material provision this month.');
    if (s.isNew) pt(s.code, s.name, `New centre: not in last month's statement, so its ${H.open} is 0.`, null, 'Confirm this centre is new and has no balance carried from last month.');
  }
  for (const [code, n] of provCount) if (n > 1) pt(code, (stmt.find((s) => s.code === code) || {}).name || '', `Listed ${n} times in Rajeev's provision file; the rows are added together (one statement row per centre).`, provByCode.get(code), 'For information.');
  for (const b of rajeev.blanks) pt(b.code, (stmt.find((s) => s.code === b.code) || {}).name || '', `Rajeev's provision file row ${b.row}: "Materials" is blank, taken as 0.`, null, 'Check whether an amount is missing in Rajeev\'s file.');
  for (const o of opening.centres.values()) if (o.blank.ytd) pt(o.code, o.name, `Last month's statement row ${o.row}: YTD is blank, taken as 0.`, null, 'Check whether last month\'s balance is missing.');
  for (const s of stmt) {
    if (s.cfSop >= -EPS) continue;
    if (!adjInfo) pt(s.code, s.name, `Prov CF is negative (${fmtINR(s.cfSop)}): more booked than provided. Carried with its minus sign (SOP Step 14).`, s.cfSop, 'Review this centre\'s bookings and provision.');
    else if (adjInfo.adj.applied) pt(s.code, s.name, `Prov CF was ${fmtINR(s.cfSop)} (more booked than provided). Set to 0 under the negative carry-forward rule; the amount was taken from the ${adjInfo.adj.recipients.length} centres with the highest Prov CF (see the adjustment table in the "Points to check" sheet).`, s.cfSop, 'Review this centre\'s bookings and provision.');
    else pt(s.code, s.name, `Prov CF is negative (${fmtINR(s.cfSop)}) and could NOT be absorbed: the negative total ${fmtINR(adjInfo.adj.total)} is more than the top ${rule.topN} centres can take.`, s.cfSop, 'The negative carry-forward rule could not be applied this month. Review with finance.');
  }
  const defaulted = kept.filter((r) => r.tagBy === 'Default');
  if (defaulted.length) pt(null, '', `${defaulted.length} Working line(s) had no year in the description and were counted as ${fyLabelWorking(period.cur)} by default (total ${fmtINR(defaulted.reduce((a, r) => a + r.bal, 0))}). They are marked in the Remarks column of the Working sheet.`, defaulted.reduce((a, r) => a + r.bal, 0), 'For information. Check these lines if any of them belong to last year.');
  if (noVendor) pt(null, '', `${noVendor} Working line(s) have no vendor name (Attribute 3 and Party Name are both blank in the ERP extract).`, null, 'For information. These lines are still counted.');
  if (!kept.length) pt(null, '', 'No P2P or GST reclass lines in the ERP extract, so the booking column is 0.', 0, 'Confirm the ERP extract is complete for this month.');

  return { period, labels: H, stmt, T, pivot: [...pivot.values()], checks, status: status(), points, adjInfo, earlier: earlier.length, totals: { keptTotalP, curTotalP, removedTotalP, extractTotalP, openYtd } };
}

async function buildWorkbook(state, R, meta, rule) {
  const { period, kept, extract, rajeev } = state;
  const H = R.labels, L = period.display;
  const wb = new ExcelJS.Workbook();
  wb.creator = TOOL + ' ' + VERSION;
  wb.created = meta && meta.now ? meta.now : new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const wsSum = wb.addWorksheet('Summary');
  const wsPiv = wb.addWorksheet('Pivot');
  const wsRaj = wb.addWorksheet('Provision_Rajeev');
  const wsWork = wb.addWorksheet('Working');
  const wsRaw = wb.addWorksheet('Raw');
  writeRaw(wsRaw, extract);
  writeWorking(wsWork, extract, kept, period);

  /* ---- Pivot (Step 10): Exp Year filter on the current FY ---- */
  wsPiv.getCell('A1').value = 'Exp Year'; wsPiv.getCell('B1').value = fyLabelWorking(period.cur);
  wsPiv.getCell('A1').font = { bold: true };
  wsPiv.getRow(3).values = ['SEGMENT1', 'BALANCING_SEGMENT_DESC', 'Sum of BALANCE'];
  const rows = R.pivot.slice().sort((a, b) => a.code - b.code || (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  let tot = 0;
  rows.forEach((p, i) => { tot += p.amt; wsPiv.getRow(4 + i).values = [p.code, p.name || null, p.amt / 100]; });
  const pGT = 4 + rows.length;
  wsPiv.getRow(pGT).values = ['Grand Total', null, tot / 100];
  formatTable(wsPiv, [3], 1, 3, 4, pGT);
  for (let c = 1; c <= 3; c++) { const cell = wsPiv.getRow(pGT).getCell(c); cell.font = { bold: true }; cell.fill = HEAD_FILL; }
  wsPiv.getColumn('C').numFmt = '#,##0.00'; wsPiv.getColumn('A').width = 12; wsPiv.getColumn('B').width = 28; wsPiv.getColumn('C').width = 16;
  wsPiv.views = [{ state: 'frozen', ySplit: 3 }];
  const pEnd = Math.max(4, pGT - 1);

  /* ---- Provision_Rajeev (Step 11): Materials as is ---- */
  wsRaj.getRow(1).values = ['S.No.', 'Center Name', 'Vendor Name', 'Storage Code', 'State'];
  const rEnd = 3 + rajeev.lines.length;
  wsRaj.getCell('F2').value = { formula: `SUM(F4:F${rEnd})`, result: rajeev.lines.reduce((a, l) => a + l.amount, 0) };
  wsRaj.getRow(3).values = [null, null, null, null, null, 'Materials (per month, with GST)'];
  rajeev.lines.forEach((l, i) => { wsRaj.getRow(4 + i).values = [l.sno === '' ? null : l.sno, l.name || null, l.vendor || null, l.code, l.state || null, l.amount]; });
  wsRaj.getColumn('F').numFmt = '#,##0.00'; wsRaj.getColumn('F').width = 18;
  formatTable(wsRaj, [1], 1, 5, 4, rEnd); formatTable(wsRaj, [3], 6, 6, 4, rEnd);
  { const cell = wsRaj.getCell('F2'); cell.font = { bold: true }; cell.border = BOX; }
  wsRaj.views = [{ state: 'frozen', ySplit: 3 }];
  wsRaj.getColumn('B').width = 26; wsRaj.getColumn('C').width = 18; wsRaj.getColumn('D').width = 10; wsRaj.getColumn('E').width = 16;

  /* ---- Summary (Steps 12-15) ---- */
  wsSum.getCell('A2').value = `Material Provision for the month of ${L}`;
  wsSum.getCell('A2').font = { bold: true, underline: true };
  wsSum.getRow(4).values = [null, 'Cost Center', 'Cost Center Name', H.open, H.booking, H.cf, H.prov, H.ytd];
  const first = 5, last = 4 + R.stmt.length;
  R.stmt.forEach((s, i) => {
    const rn = first + i;
    wsSum.getRow(rn).values = [null, s.code, s.name, s.open,
      { formula: `SUMIFS(Pivot!$C$4:$C$${pEnd},Pivot!$A$4:$A$${pEnd},$B${rn})`, result: s.booking },
      Math.abs(s.adj) > 0 ? s.cf : { formula: `D${rn}-E${rn}`, result: s.cf },
      { formula: `SUMIFS(Provision_Rajeev!$F$4:$F$${rEnd},Provision_Rajeev!$D$4:$D$${rEnd},$B${rn})`, result: s.prov },
      { formula: `F${rn}+G${rn}`, result: s.ytd }];
  });
  const T = { D: R.T.open, E: R.T.booking, F: R.T.cf, G: R.T.prov, H: R.T.ytd };
  for (const [c, v] of Object.entries(T)) { const cell = wsSum.getCell(`${c}3`); cell.value = { formula: `SUBTOTAL(9,${c}${first}:${c}${Math.max(first, last)})`, result: v }; cell.border = BOX; cell.font = { bold: true }; }
  for (const c of ['D', 'E', 'F', 'G', 'H']) { wsSum.getColumn(c).numFmt = FMT_AMT; wsSum.getColumn(c).width = 15; }
  formatTable(wsSum, [4], 2, 8, 5, last);
  wsSum.getRow(4).height = 30;
  wsSum.views = [{ state: 'frozen', ySplit: 4 }];
  wsSum.getColumn('A').width = 3; wsSum.getColumn('B').width = 10; wsSum.getColumn('C').width = 24;
  wsSum.autoFilter = { from: { row: 4, column: 2 }, to: { row: 4, column: 8 } };

  writePoints(wb.addWorksheet('Points to check'), L, R.points, R.stmt, R.adjInfo ? { topN: rule.topN, columns: NEG_COL, adjInfo: R.adjInfo } : null);
  return wb.xlsx.writeBuffer();
}

/** The engine; `negativeCfRule` overrides NEGATIVE_CF_RULE (used by the tests to cover both ways). */
function make(negativeCfRule) {
  const rule = negativeCfRule === undefined ? NEGATIVE_CF_RULE : negativeCfRule;
  return createEngine({
    toolName: TOOL, version: VERSION, sopVersion: 'SOP v1.0 · October 2026',
    extract: { accounts: [51611], accountsText: 'account 51611', glParams: { from: 51611, to: 51611 }, multipleAccountsAllowed: false },
    opening: () => ({ columns: [{ key: 'ytd', header: 'YTD', choosable: true }], aprilNote: false }),
    provisionFile: { sheet: { test: /^\s*hk(?![a-z])/i }, amount: { startsWith: 'materials', text: 'Materials' }, noun: 'housekeeping' },
    stopInApril: true,
    compute: (state) => compute(state, rule),
    buildWorkbook: (state, R, meta) => buildWorkbook(state, R, meta, rule),
    ui: {
      title: 'Housekeeping Material Provision',
      decimals: 0, // shown without paise, as in the SOP (exact values are kept)
      files: {
        extract: { title: 'Oracle ERP extract', desc: 'IIVF Account Analysis report · account 51611 only · one full month, raw download' },
        opening: { title: 'Last month\'s statement', desc: 'Previous month\'s Material Provision statement · its YTD column is the opening' },
        rajeev: { title: 'Provision file', desc: 'This month\'s file from Rajeev · sheet "HK <month>" · Materials column, as it is' },
      },
      outputName: (period) => `Housekeeping Material_${period.display}.xlsx`,
      sheets: 'Summary, Pivot, Provision_Rajeev, Working, Raw, Points to check',
      note: 'The statement is built and the checks have run. Only this financial year\'s bookings are used; lines with no year in the description are counted as this financial year.',
      pointsNote: 'Unusual items found in this month\'s files, for information. The same list is in the "Points to check" sheet of the Excel.',
      keys: (R) => [[R.labels.ytd, R.T.ytd, 'Closing provision · next month\'s opening'], [R.labels.prov, R.T.prov, 'This month\'s provision (Materials)'], [R.labels.booking, R.T.booking, 'Booked this month']],
      totals: (R) => [[R.labels.open, R.T.open], [R.labels.booking, R.T.booking], [R.labels.cf, R.T.cf], [R.labels.prov, R.T.prov], [R.labels.ytd, R.T.ytd]],
    },
  });
}

module.exports = make();
module.exports.withNegativeCfRule = make;
module.exports.NEGATIVE_CF_RULE = NEGATIVE_CF_RULE;
