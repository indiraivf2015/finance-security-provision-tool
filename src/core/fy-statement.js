// "Two-year" provision statement (the Security layout): one row per cost centre with FY{cur}/FY{prev} opening
// provision, bookings, carry-forward, this month's provision as per Rajeev, and YTD. SOP Steps 10-16 + Section 04
// checks. Copied from the approved Security engine (compute + buildWorkbook) and driven by the provision's settings:
//
// cfg = {
//   toolName: 'Housekeeping Expenses Provision Tool',
//   title: (L) => 'Housekeeping Exp – Provision for the month of ' + L,
//   pivotByVendor: true,                  Step 10 rows: BALANCING_SEGMENT_DESC, SEGMENT1, ATTRIBUTE3 (else SEGMENT1, DESC)
//   amountText: 'Manpower',               the column used from Rajeev's file (already read by the provision-file reader)
//   gst: { divisor: 118, basicFactor: 100, finalFactor: 1.135 },   Steps 11-12
//   negativeCarryForward: { topN: 10 } | null,
//   summaryBlock: true,                   Step 17 block (FY{prev} Pro CF, YTD FY{cur}, Total, Every month)
//   amountFormat: FMT_AMT2,
// }
'use strict';
const { EPS, fyLabelWorking, fyLabelHeader, toPaise, fmtINR } = require('./common');
const { ExcelJS, formatTable, FMT_AMT2, HEAD_FILL, THIN, BOX } = require('./excel');
const { applyNegativeCarryForward, checkList } = require('./rules');
const { isWrongKind, GST_RECLASS_CREATOR } = require('./clean');
const { writeRaw, writeWorking, writePoints } = require('./sheets');
const { skippedSummary } = require('./extract');
const { manualRowText } = require('./opening');

const NEG_COLUMNS = (period) => [
  { sopF: 'cfCurSop', adjF: 'adjCur', finF: 'cfCur', label: `${fyLabelHeader(period.cur)} Pro CF` },
  { sopF: 'cfPrevSop', adjF: 'adjPrev', finF: 'cfPrev', label: `${fyLabelHeader(period.prev)} Pro CF` },
];

function compute(state, cfg) {
  const { period, kept, removed, opening, rajeev, extract } = state;

  // Step 10: pivot (sums in paise, exact). Text grouping ignores case, as Excel's pivot does.
  const pivot = new Map();
  const byCode = new Map(); // code -> {cur, prev} paise
  for (const r of kept) {
    const key = cfg.pivotByVendor
      ? r.ccNameRaw.toLowerCase() + '\u0000' + r.seg1 + '\u0000' + r.vendor.toLowerCase()
      : r.seg1 + '\u0000' + r.ccNameRaw.toLowerCase();
    if (!pivot.has(key)) pivot.set(key, { code: r.seg1, name: r.ccNameRaw, vendor: r.vendor, prev: 0, cur: 0 });
    const p = pivot.get(key); const amt = toPaise(r.bal);
    if (r.fy === period.cur) p.cur += amt; else p.prev += amt;
    if (!byCode.has(r.seg1)) byCode.set(r.seg1, { cur: 0, prev: 0, name: r.ccName });
    const b = byCode.get(r.seg1); if (r.fy === period.cur) b.cur += amt; else b.prev += amt;
  }

  // Steps 11-12, row by row, in the SOP's order of operations: Basic = amount ÷ 118 × 100; Final = Basic × 1.135
  for (const l of rajeev.lines) { l.basic = l.amount / cfg.gst.divisor * cfg.gst.basicFactor; l.final = l.basic * cfg.gst.finalFactor; }
  // Step 16: one statement row per centre, so rows with the same code (e.g. Main + OPD) are added together
  const rajByCode = new Map(); const rajCount = new Map();
  for (const l of rajeev.lines) { rajByCode.set(l.code, (rajByCode.get(l.code) || 0) + l.final); rajCount.set(l.code, (rajCount.get(l.code) || 0) + 1); }

  // Statement rows: booked centres by code, then last month's other centres in last month's order, then new centres
  const booked = [...byCode.keys()].sort((a, b) => a - b);
  const codes = booked.concat(opening.order.filter((c) => !byCode.has(c)));
  codes.push(...[...rajByCode.keys()].filter((c) => !opening.centres.has(c) && !byCode.has(c)).sort((a, b) => a - b));
  const rajName = new Map(); for (const l of rajeev.lines) if (!rajName.has(l.code) && l.name) rajName.set(l.code, l.name);
  const stmt = codes.map((code) => {
    const o = opening.centres.get(code);
    const b = byCode.get(code) || { cur: 0, prev: 0 };
    const proCur = o ? o.v.ytd : 0, proPrev = o ? o.v.cf : 0;   // Step 13
    const bkCur = b.cur / 100, bkPrev = b.prev / 100;           // Step 14
    const cfCur = proCur - bkCur, cfPrev = proPrev - bkPrev;    // Step 15
    const raj = rajByCode.get(code) || 0;                       // Step 16
    const name = (o && o.name) || (byCode.get(code) && byCode.get(code).name) || rajName.get(code) || '';
    return { code, name, proCur, proPrev, bkCur, bkPrev, cfCurSop: cfCur, cfPrevSop: cfPrev, adjCur: 0, adjPrev: 0, cfCur, cfPrev, raj, ytd: 0, isNew: !o };
  });

  const adjInfo = cfg.negativeCarryForward ? applyNegativeCarryForward(stmt, NEG_COLUMNS(period), cfg.negativeCarryForward.topN) : null;
  for (const s of stmt) { s.cfCur = s.cfCurSop + s.adjCur; s.cfPrev = s.cfPrevSop + s.adjPrev; s.ytd = s.cfCur + s.raj; }
  // Manual fixed rows (opening.manualRows): after the centres, in last month's order. Carried forward unchanged: no bookings,
  // no Rajeev provision, not part of the negative carry-forward rule; YTD = opening.
  for (const m of opening.manualRows || []) {
    const proCur = m.v.ytd, proPrev = m.v.cf;
    stmt.push({ code: m.code, name: m.name, manual: true, manualRow: m.row, proCur, proPrev, bkCur: 0, bkPrev: 0, cfCurSop: proCur, cfPrevSop: proPrev, adjCur: 0, adjPrev: 0, cfCur: proCur, cfPrev: proPrev, raj: 0, ytd: proCur, isNew: false });
  }

  // Totals
  const sum = (f) => stmt.reduce((a, s) => a + s[f], 0);
  const T = { proCur: sum('proCur'), proPrev: sum('proPrev'), bkCur: sum('bkCur'), bkPrev: sum('bkPrev'), cfCur: sum('cfCur'), cfPrev: sum('cfPrev'), raj: sum('raj'), ytd: sum('ytd'), adjCur: sum('adjCur'), adjPrev: sum('adjPrev') };
  T.total = T.cfPrev + T.ytd; T.everyMonth = T.total / 4; // Step 17 summary block
  const keptTotalP = kept.reduce((a, r) => a + toPaise(r.bal), 0);
  const removedTotalP = removed.reduce((a, r) => a + toPaise(r.bal), 0);
  const extractTotalP = extract.rows.reduce((a, r) => a + toPaise(r.bal), 0);
  const skip = skippedSummary(extract, cfg.accountsText);
  let pivotTotalP = 0; for (const p of pivot.values()) pivotTotalP += p.cur + p.prev;
  const rajIn = rajeev.lines.reduce((a, l) => a + l.amount, 0);
  const rajFinal = rajeev.lines.reduce((a, l) => a + l.final, 0);
  let openYtd = 0, openCf = 0; for (const o of [...opening.centres.values(), ...(opening.manualRows || [])]) { openYtd += o.v.ytd; openCf += o.v.cf; }

  // Section 04 checks + completeness checks
  const C = fyLabelHeader(period.cur), Pv = fyLabelHeader(period.prev);
  const { checks, add, status } = checkList();
  add('04 · check 1', 'Every Working line has an Exp Year tag', 0, kept.filter((r) => r.fy !== period.cur && r.fy !== period.prev).length, 'count0fail');
  const noVendor = kept.filter((r) => !r.vendor).length;
  add('04 · check 1', 'Every Working line has a vendor name in Attribute 3 (lines with no Attribute 3 and no Party Name in the ERP cannot be filled)', 0, noVendor, 'count0');
  add('04 · check 2', 'Working contains only P2P bookings and GST reclass entries by ' + GST_RECLASS_CREATOR, 0, kept.filter(isWrongKind).length, 'count0fail');
  add('04 · check 3', 'Pivot grand total = total BALANCE of the Working sheet', keptTotalP / 100, pivotTotalP / 100);
  add('04 · check 4', `Totals row: ${C} Pro − ${C} Booking = ${C} Pro CF`, T.proCur - T.bkCur, T.cfCur);
  add('04 · check 4', `Totals row: ${Pv} Pro − ${Pv} Booking = ${Pv} Pro CF`, T.proPrev - T.bkPrev, T.cfPrev);
  add('04 · check 5', `Totals row: ${C} Pro CF + ${period.display} Pro as per Rajeev = YTD ${C}`, T.cfCur + T.raj, T.ytd);
  add('04 · check 5', `Rajeev total in the statement = total of the GST conversion (Part D, ${cfg.amountText} only)`, rajFinal, T.raj);
  if (cfg.summaryBlock) add('04 · check 6', `Summary block: ${Pv} Pro CF + YTD ${C} = Total, both lines equal the column totals`, T.cfPrev + T.ytd, T.total);
  if (skip.list.length) add('Completeness', 'ERP extract total = Working total + Removed lines total + other Housekeeping account lines skipped (no line lost)', (extractTotalP + skip.totalP) / 100, (keptTotalP + removedTotalP + skip.totalP) / 100);
  else add('Completeness', 'ERP extract total = Working total + Removed lines total (no line lost)', extractTotalP / 100, (keptTotalP + removedTotalP) / 100);
  add('Completeness', 'Statement bookings (both years) = Working total (every booked centre is in the statement)', keptTotalP / 100, T.bkCur + T.bkPrev);
  add('Completeness', `Provision file: total of the "${cfg.amountText}" column = GST conversion input total`, rajeev.columnTotal, rajIn);
  add('Completeness · Step 13', `Opening ${C} Pro = last month's YTD ${C} total`, openYtd, T.proCur);
  add('Completeness · Step 13', `Opening ${Pv} Pro = last month's ${Pv} Pro CF total`, openCf, T.proPrev);
  if (adjInfo) {
    add('Negative CF rule', `${C} CF adjustments net to zero (column total unchanged)`, 0, T.adjCur);
    add('Negative CF rule', `${Pv} CF adjustments net to zero (column total unchanged)`, 0, T.adjPrev);
    add('Negative CF rule', 'No negative Pro CF left after the adjustment', 0, stmt.filter((s) => s.cfCur < -EPS || s.cfPrev < -EPS).length, 'count0fail');
  }

  // "Points to check": unusual items in plain words. Information only; nothing here changes any figure.
  const points = [];
  const pt = (code, name, point, amount, check) => points.push({ code, name, point, amount: amount == null ? null : Math.round(amount * 100) / 100, check });
  for (const s of stmt) {
    if (s.manual) { pt(s.code, s.name, manualRowText(opening.manualRows.find((m) => m.row === s.manualRow), opening.columns), s.ytd, 'For information. Finance adds this row by hand; it stays until finance removes it from the statement.'); continue; }
    if (!rajByCode.has(s.code)) {
      const bk = s.bkCur + s.bkPrev;
      pt(s.code, s.name, `Not in Rajeev's provision file, so "${period.display} Pro as per Rajeev" is 0.` + (Math.abs(bk) > EPS ? ` ${fmtINR(bk)} was booked for this centre this month.` : ''), Math.abs(bk) > EPS ? bk : null, 'Check with Rajeev whether this centre should have a provision this month.');
    }
    if (s.isNew) pt(s.code, s.name, `New centre: not in last month's statement, so its opening ${C} Pro and ${Pv} Pro are 0.`, null, 'Confirm this centre is new and has no balance carried from last month.');
  }
  for (const [code, n] of rajCount) if (n > 1) pt(code, (stmt.find((s) => s.code === code) || {}).name || '', `Listed ${n} times in Rajeev's provision file; the rows are added together (one statement row per centre).`, rajByCode.get(code), 'For information.');
  for (const b of rajeev.blanks) pt(b.code, (stmt.find((s) => s.code === b.code) || {}).name || '', `Rajeev's provision file row ${b.row}: "${cfg.amountText}" is blank, taken as 0.`, null, 'Check whether an amount is missing in Rajeev\'s file.');
  if (skip.text) pt(null, '', skip.text, skip.totalP / 100, 'For information. Upload the same file to the other Housekeeping tool.');
  for (const o of opening.centres.values()) if (o.blank.ytd || o.blank.cf) pt(o.code, o.name, `Last month's statement row ${o.row}: ${o.blank.ytd ? 'YTD ' + C : ''}${o.blank.ytd && o.blank.cf ? ' and ' : ''}${o.blank.cf ? Pv + ' Pro CF' : ''} is blank, taken as 0.`, null, 'Check whether last month\'s balance is missing.');
  if (adjInfo) {
    for (const s of stmt) {
      if (s.manual) continue;
      for (const { sopF, adjF, label } of NEG_COLUMNS(period)) {
        if (s[sopF] >= -EPS) continue;
        const inf = adjInfo[adjF];
        if (inf.applied) pt(s.code, s.name, `${label} was ${fmtINR(s[sopF])} (more booked than provided). Set to 0 under the negative carry-forward rule; the amount was taken from the ${inf.recipients.length} centres with the highest ${label} (see the adjustment table in the "Points to check" sheet).`, s[sopF], 'Review this centre\'s bookings and provision.');
        else pt(s.code, s.name, `${label} is negative (${fmtINR(s[sopF])}) and could NOT be absorbed: the negative total ${fmtINR(inf.total)} is more than the top ${cfg.negativeCarryForward.topN} centres can take.`, s[sopF], 'The negative carry-forward rule could not be applied this month. Review with finance.');
      }
    }
  }
  const defaulted = kept.filter((r) => r.tagBy === 'Default' && !r.yearWritten);
  for (const r of kept) if (r.tagBy === 'Default' && r.yearWritten) pt(r.seg1, r.ccName, `ERP row ${r.erpRow}: the description writes a year, but the year is unclear (${r.basis}). Counted as ${fyLabelWorking(period.cur)} by default. "${r.desc.slice(0, 120)}"`, r.bal, 'Check which financial year this line belongs to.');
  if (defaulted.length) pt(null, '', `${defaulted.length} Working line(s) had no year in the description and were counted as ${fyLabelWorking(period.cur)} by default (total ${fmtINR(defaulted.reduce((a, r) => a + r.bal, 0))}). They are marked in the Remarks column of the Working sheet.`, defaulted.reduce((a, r) => a + r.bal, 0), 'For information. Check these lines if any of them belong to last year.');
  if (noVendor) pt(null, '', `${noVendor} Working line(s) have no vendor name (Attribute 3 and Party Name are both blank in the ERP extract).`, null, 'For information. These lines are still counted.');
  if (!kept.length) pt(null, '', 'No P2P or GST reclass lines in the ERP extract, so all booking columns are 0.', 0, 'Confirm the ERP extract is complete for this month.');

  return { period, stmt, T, pivot: [...pivot.values()], checks, status: status(), points, adjInfo, totals: { keptTotalP, removedTotalP, extractTotalP, rajIn, rajFinal, openYtd, openCf } };
}

async function buildWorkbook(state, R, meta, cfg, version) {
  const { period, kept, extract, rajeev } = state;
  const wb = new ExcelJS.Workbook();
  wb.creator = cfg.toolName + ' ' + version;
  wb.created = meta && meta.now ? meta.now : new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const L = period.display, C = fyLabelHeader(period.cur), Pv = fyLabelHeader(period.prev);
  const FMT = cfg.amountFormat;
  const wsSum = wb.addWorksheet('Summary');
  const wsPiv = wb.addWorksheet('Pivot');
  const wsRaj = wb.addWorksheet('Provision_Rajeev');
  const wsWork = wb.addWorksheet('Working');
  const wsRaw = wb.addWorksheet('Raw');
  writeRaw(wsRaw, extract);
  writeWorking(wsWork, extract, kept, period);

  /* ---- Pivot (Step 10) ---- */
  let pivAmt; // column letters of the FY{prev} / FY{cur} amounts and of SEGMENT1, for the Summary's SUMIFS
  wsPiv.getCell('A3').value = 'Sum of BALANCE'; wsPiv.getCell('A3').font = { bold: true };
  const byName = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  let pivRows, headers, rowOf;
  if (cfg.pivotByVendor) {
    headers = ['BALANCING_SEGMENT_DESC', 'SEGMENT1', 'ATTRIBUTE3', fyLabelWorking(period.prev), fyLabelWorking(period.cur), 'Grand Total'];
    pivRows = R.pivot.slice().sort((a, b) => byName(a.name.toLowerCase(), b.name.toLowerCase()) || a.code - b.code || byName(a.vendor.toLowerCase(), b.vendor.toLowerCase()));
    rowOf = (p) => [p.name || null, p.code, p.vendor || '(blank)', p.prev ? p.prev / 100 : null, p.cur ? p.cur / 100 : null, (p.prev + p.cur) / 100];
    pivAmt = { code: 'B', prev: 'D', cur: 'E' };
  } else {
    headers = ['SEGMENT1', 'BALANCING_SEGMENT_DESC', fyLabelWorking(period.prev), fyLabelWorking(period.cur), 'Grand Total'];
    pivRows = R.pivot.slice().sort((a, b) => a.code - b.code || byName(a.name, b.name));
    rowOf = (p) => [p.code, p.name || null, p.prev ? p.prev / 100 : null, p.cur ? p.cur / 100 : null, (p.prev + p.cur) / 100];
    pivAmt = { code: 'A', prev: 'C', cur: 'D' };
  }
  const nCol = headers.length;
  wsPiv.getCell(3, nCol - 2).value = 'Exp Year'; wsPiv.getCell(3, nCol - 2).font = { bold: true }; // above the year columns
  wsPiv.getRow(4).values = headers;
  let pPrev = 0, pCur = 0;
  pivRows.forEach((p, i) => { pPrev += p.prev; pCur += p.cur; wsPiv.getRow(5 + i).values = rowOf(p); });
  const pGT = 5 + pivRows.length;
  const gt = new Array(nCol).fill(null); gt[0] = 'Grand Total'; gt[nCol - 3] = pPrev / 100; gt[nCol - 2] = pCur / 100; gt[nCol - 1] = (pPrev + pCur) / 100;
  wsPiv.getRow(pGT).values = gt;
  formatTable(wsPiv, [4], 1, nCol, 5, pGT);
  for (let c = 1; c <= nCol; c++) { const cell = wsPiv.getRow(pGT).getCell(c); cell.font = { bold: true }; cell.fill = HEAD_FILL; }
  for (let c = nCol - 2; c <= nCol; c++) { wsPiv.getColumn(c).numFmt = '#,##0.00'; wsPiv.getColumn(c).width = 14; }
  if (cfg.pivotByVendor) { wsPiv.getColumn('A').width = 28; wsPiv.getColumn('B').width = 11; wsPiv.getColumn('C').width = 40; }
  else { wsPiv.getColumn('B').width = 28; wsPiv.getColumn('A').width = 12; }
  wsPiv.views = [{ state: 'frozen', ySplit: 4 }];
  const pEnd = Math.max(5, pGT - 1);

  /* ---- Provision_Rajeev (Steps 11-12) ---- */
  const g = cfg.gst;
  wsRaj.getRow(1).values = ['S.No.', 'Center Name', 'Vendor Name', 'Storage Code', 'State'];
  const rEnd = 3 + rajeev.lines.length;
  wsRaj.getCell('F2').value = { formula: `SUM(F4:F${rEnd})`, result: R.totals.rajIn };
  wsRaj.getRow(3).values = [null, null, null, null, null, `${cfg.amountText} (per month, with GST)`, 'Basic', `With ${Math.round((g.finalFactor - 1) * 1000) / 10}%GST`, 'Cost Center Name (Summary)'];
  const nameOf = new Map(R.stmt.filter((s) => !s.manual).map((s) => [s.code, s.name]));
  rajeev.lines.forEach((l, i) => {
    const rn = 4 + i;
    wsRaj.getRow(rn).values = [l.sno === '' ? null : l.sno, l.name || null, l.vendor || null, l.code, l.state || null, l.amount,
      { formula: `F${rn}/${g.divisor}*${g.basicFactor}`, result: l.basic },
      { formula: `G${rn}*${g.finalFactor}`, result: l.final },
      { formula: `_xlfn.XLOOKUP(D${rn},Summary!B:B,Summary!C:C)`, result: nameOf.get(l.code) || '' }];
  });
  ['F', 'G', 'H'].forEach((c) => { wsRaj.getColumn(c).numFmt = '#,##0.00'; wsRaj.getColumn(c).width = 16; });
  formatTable(wsRaj, [1], 1, 5, 4, rEnd);
  formatTable(wsRaj, [3], 6, 9, 4, rEnd);
  { const cell = wsRaj.getCell('F2'); cell.font = { bold: true }; cell.border = BOX; }
  wsRaj.views = [{ state: 'frozen', ySplit: 3 }];
  wsRaj.getColumn('B').width = 26; wsRaj.getColumn('C').width = 18; wsRaj.getColumn('D').width = 10; wsRaj.getColumn('E').width = 16; wsRaj.getColumn('I').width = 24;

  /* ---- Summary (Steps 13-16) ---- */
  wsSum.getCell('A2').value = cfg.title(L);
  wsSum.getCell('A2').font = { bold: true, underline: true };
  wsSum.getRow(4).values = [null, 'Cost Center', 'Cost Center Name', `${C} Pro`, `${Pv} Pro`, `${C} Booking`, `${Pv} Booking`, `${C} Pro CF`, `${Pv} Pro CF`, `${L} Pro as per Rajeev`, `YTD ${C}`];
  const first = 5, last = 4 + R.stmt.length;
  const PIV = (c) => `Pivot!$${c}$5:$${c}$${pEnd}`;
  R.stmt.forEach((s, i) => {
    const rn = first + i;
    if (s.manual) { // manual fixed row: plain values, so the centre's bookings and Rajeev provision stay on its normal row
      wsSum.getRow(rn).values = [null, s.code, s.name, s.proCur, s.proPrev, 0, 0, { formula: `D${rn}-F${rn}`, result: s.cfCur }, { formula: `E${rn}-G${rn}`, result: s.cfPrev }, 0, { formula: `H${rn}+J${rn}`, result: s.ytd }];
      return;
    }
    wsSum.getRow(rn).values = [
      null, s.code, s.name, s.proCur, s.proPrev,
      { formula: `SUMIFS(${PIV(pivAmt.cur)},${PIV(pivAmt.code)},$B${rn})`, result: s.bkCur },
      { formula: `SUMIFS(${PIV(pivAmt.prev)},${PIV(pivAmt.code)},$B${rn})`, result: s.bkPrev },
      // Centres changed by the negative carry-forward rule hold the final figure (as finance enters it);
      // the before/after detail is in the adjustment table on the "Points to check" sheet.
      Math.abs(s.adjCur) > 0 ? s.cfCur : { formula: `D${rn}-F${rn}`, result: s.cfCur },
      Math.abs(s.adjPrev) > 0 ? s.cfPrev : { formula: `E${rn}-G${rn}`, result: s.cfPrev },
      { formula: `SUMIFS(Provision_Rajeev!$H$4:$H$${rEnd},Provision_Rajeev!$D$4:$D$${rEnd},$B${rn})`, result: s.raj },
      { formula: `H${rn}+J${rn}`, result: s.ytd },
    ];
  });
  const tot = { D: R.T.proCur, E: R.T.proPrev, F: R.T.bkCur, G: R.T.bkPrev, H: R.T.cfCur, I: R.T.cfPrev, J: R.T.raj, K: R.T.ytd };
  for (const [c, v] of Object.entries(tot)) { const cell = wsSum.getCell(`${c}3`); cell.value = { formula: `SUBTOTAL(9,${c}${first}:${c}${Math.max(first, last)})`, result: v }; cell.border = BOX; cell.font = { bold: true }; }
  for (const c of ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) { wsSum.getColumn(c).numFmt = FMT; wsSum.getColumn(c).width = 15; }
  formatTable(wsSum, [4], 2, 11, 5, last);
  wsSum.getRow(4).height = 30;
  wsSum.views = [{ state: 'frozen', ySplit: 4 }];
  wsSum.getColumn('A').width = 3; wsSum.getColumn('B').width = 10; wsSum.getColumn('C').width = 24;
  wsSum.autoFilter = { from: { row: 4, column: 2 }, to: { row: 4, column: 11 } };

  if (cfg.summaryBlock) {
    // Step 17: summary block beside the statement. Kept below the header row so that next month's reader never
    // sees a second "YTD FY.." heading in the header row.
    const blk = [
      [`${Pv} Pro CF`, { formula: 'I3', result: R.T.cfPrev }],
      [`YTD ${C}`, { formula: 'K3', result: R.T.ytd }],
      ['Total', { formula: 'N5+N6', result: R.T.total }],
      ['Every month', { formula: 'N7/4', result: R.T.everyMonth }],
    ];
    blk.forEach(([label, v], i) => { const r = 5 + i; wsSum.getCell(`M${r}`).value = label; wsSum.getCell(`N${r}`).value = v; wsSum.getCell(`N${r}`).numFmt = FMT; for (const c of ['M', 'N']) wsSum.getCell(`${c}${r}`).border = BOX; });
    wsSum.getCell('O7').value = { formula: 'N7/10^6', result: R.T.total / 1e6 }; wsSum.getCell('O7').numFmt = '0.00" Mn"'; wsSum.getCell('O7').border = BOX;
    for (const c of ['M7', 'N7', 'O7']) { wsSum.getCell(c).font = { bold: true }; wsSum.getCell(c).fill = HEAD_FILL; }
    wsSum.getColumn('M').width = 14; wsSum.getColumn('N').width = 16; wsSum.getColumn('O').width = 10;
  }

  /* ---- Points to check ---- */
  writePoints(wb.addWorksheet('Points to check'), L, R.points, R.stmt,
    R.adjInfo ? { topN: cfg.negativeCarryForward.topN, columns: NEG_COLUMNS(period), adjInfo: R.adjInfo } : null);

  return wb.xlsx.writeBuffer();
}

module.exports = { compute, buildWorkbook, FMT_AMT2, THIN };
