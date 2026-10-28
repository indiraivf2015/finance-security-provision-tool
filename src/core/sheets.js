// Output sheets every provision shares: Raw (the ERP extract as downloaded), Working (Steps 5-9) and
// Points to check (information for the finance reviewer, plus the negative carry-forward table when that rule applies).
// Layouts copied from the approved Security engine.
'use strict';
const { norm, colLetter, fyLabelWorking } = require('./common');
const { cellInfo, origValue, formatTable, FMT_AMT2 } = require('./excel');

/** Raw: the ERP extract exactly as downloaded */
function writeRaw(wsRaw, extract) {
  const src = extract.rawSheet;
  src.eachRow({ includeEmpty: true }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      const i = cellInfo(cell);
      wsRaw.getCell(r, c).value = i.k === 'blank' ? null : origValue(i);
    });
  });
  src.columns && src.columns.forEach((col, i) => { if (col && col.width) wsRaw.getColumn(i + 1).width = col.width; });
  const hdrRow = extract.headerRow, lastCol = Math.max(...extract.headers.map((h) => h.c));
  formatTable(wsRaw, [hdrRow], 1, lastCol, hdrRow + 1, src.rowCount);
  wsRaw.getCell('A1').font = { bold: true, size: 12 };
  // report parameter block above the table (as the ERP prints it)
  for (let r = 1; r < hdrRow; r++) {
    const row = src.getRow(r); let any = false; row.eachCell(() => { any = true; }); if (!any) continue;
    const names = []; row.eachCell((c, ci) => { names.push(ci); });
    if (norm(cellInfo(src.getCell(r, 1)).v) === 'fromdate') formatTable(wsRaw, [r], 1, Math.max(...names), r + 1, r + 1);
  }
  wsRaw.views = [{ state: 'frozen', ySplit: hdrRow }];
}

/** Working (Steps 5-9): the kept lines, "Exp Year" right of LINE_DESCRIPTION, ATTRIBUTE3 filled from PARTY_NAME */
function writeWorking(wsWork, extract, kept, period) {
  const cols = [];
  for (const h of extract.headers) {
    cols.push({ kind: 'orig', h });
    if (h.k === 'gldate') cols.push({ kind: 'glmonth' });
    if (h.k === 'linedescription') cols.push({ kind: 'exp' });
  }
  cols.push({ kind: 'remarks' });
  const wcol = {}; cols.forEach((c, i) => { wcol[c.kind === 'orig' ? c.h.k : c.kind] = i + 1; });
  wsWork.addRow(cols.map((c) => c.kind === 'orig' ? c.h.name : c.kind === 'exp' ? 'Exp Year' : c.kind === 'glmonth' ? 'GL Month' : 'Remarks'));
  const partyLetter = colLetter(wcol.partyname);
  kept.forEach((r, i) => {
    const rowNo = i + 2;
    wsWork.addRow(cols.map((c) => {
      if (c.kind === 'exp') return fyLabelWorking(r.fy);
      if (c.kind === 'remarks') return r.tagBy === 'Default' ? `Actual Booking - ${fyLabelWorking(period.cur)} by default (${r.yearWritten ? 'year in description unclear: check' : 'year not in description'})` : 'Actual Booking';
      if (c.kind === 'glmonth') { const g = r.glDate; return g ? new Date(Date.UTC(g.y, g.m - 1, 1)) : null; }
      const k = c.h.k;
      if (k === 'balance') return r.bal;
      if (k === 'attribute3' && !r.attr3 && r.party) return { formula: partyLetter + rowNo, result: r.party }; // Step 9
      return origValue(r.cells[c.h.c]);
    }));
  });
  if (wcol.glmonth) wsWork.getColumn(wcol.glmonth).numFmt = 'mmm-yy';
  formatTable(wsWork, [1], 1, cols.length, 2, kept.length + 1);
  wsWork.views = [{ state: 'frozen', ySplit: 1 }];
}

/**
 * Points to check (information only; no figure depends on this sheet).
 * neg = null, or { topN, columns: [{ sopF, adjF, finF }], adjInfo } for the negative carry-forward table.
 */
function writePoints(wsPts, label, points, stmt, neg) {
  wsPts.getCell('A1').value = `Points to check — ${label}`;
  wsPts.getCell('A1').font = { bold: true };
  wsPts.getCell('A2').value = 'Unusual items found while preparing the statement, for information.' + (neg ? ' Negative carry-forward adjustments are listed in the table further below.' : '');
  wsPts.getCell('A2').font = { italic: true };
  wsPts.getRow(4).values = ['#', 'Cost Center', 'Cost Center Name', 'Point', 'Amount', 'Suggested check'];
  if (points.length) points.forEach((p, i) => { wsPts.getRow(5 + i).values = [i + 1, p.code, p.name, p.point, p.amount, p.check]; });
  else wsPts.getRow(5).values = [null, null, null, 'No points to check this month.'];
  const pLast = 4 + Math.max(1, points.length);
  formatTable(wsPts, [4], 1, 6, 5, pLast);
  wsPts.getColumn('A').width = 5; wsPts.getColumn('B').width = 11; wsPts.getColumn('C').width = 24;
  wsPts.getColumn('D').width = 70; wsPts.getColumn('E').width = 15; wsPts.getColumn('F').width = 55;
  wsPts.getColumn('E').numFmt = FMT_AMT2;
  for (let r = 5; r <= pLast; r++) for (const c of ['D', 'F']) wsPts.getCell(c + r).alignment = { wrapText: true, vertical: 'top' };
  wsPts.views = [{ state: 'frozen', ySplit: 4 }];

  if (neg) {
    // Negative carry-forward adjustment table (finance rule): before / adjustment / after for every centre changed
    const tTop = pLast + 3;
    wsPts.getCell(`A${tTop}`).value = `Negative carry-forward adjustment — negatives set to 0 and taken from the ${neg.topN} centres with the highest carry-forward in the same column, in proportion`;
    wsPts.getCell(`A${tTop}`).font = { bold: true };
    const tHdr = tTop + 1;
    wsPts.getRow(tHdr).values = ['#', 'Cost Center', 'Cost Center Name', 'What happened', 'Pro CF before', 'Adjustment', 'Pro CF after'];
    let tr = tHdr + 1, n = 0;
    const amtCells = [];
    for (const { sopF, adjF, finF } of neg.columns) {
      const inf = neg.adjInfo[adjF];
      const rows = stmt.filter((x) => Math.abs(x[adjF]) > 0).sort((a, b) => (a[sopF] < 0 ? 0 : 1) - (b[sopF] < 0 ? 0 : 1) || b[sopF] - a[sopF]);
      if (!rows.length) continue;
      const first = tr;
      for (const x of rows) {
        n++;
        const what = x[sopF] < 0 ? `${inf.label} negative — set to 0` : `${inf.label} — share of the negatives taken from this centre`;
        wsPts.getRow(tr).values = [n, x.code, x.name, what, x[sopF], x[adjF], { formula: `E${tr}+F${tr}`, result: x[finF] }];
        amtCells.push(tr); tr++;
      }
      wsPts.getRow(tr).values = [null, null, null, `Total ${inf.label} adjustment (must be 0)`, null, { formula: `SUM(F${first}:F${tr - 1})`, result: rows.reduce((a, x) => a + x[adjF], 0) }, null];
      wsPts.getRow(tr).font = { bold: true }; amtCells.push(tr); tr++;
    }
    if (n === 0) { wsPts.getRow(tr).values = [null, null, null, 'No negative carry-forward this month — nothing adjusted.']; tr++; }
    formatTable(wsPts, [tHdr], 1, 7, tHdr + 1, tr - 1);
    for (const r of amtCells) for (const c of ['E', 'F', 'G']) wsPts.getCell(c + r).numFmt = FMT_AMT2;
    for (let r = tHdr + 1; r < tr; r++) wsPts.getCell('D' + r).alignment = { wrapText: true, vertical: 'top' };
    wsPts.getColumn('G').width = 15;
  }
  if (points.length) wsPts.properties.tabColor = { argb: 'FFF79009' }; // amber sheet tab when there is something to check
}

module.exports = { writeRaw, writeWorking, writePoints };
