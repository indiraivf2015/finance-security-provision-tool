const fs = require('fs'); const ExcelJS = require('exceljs'); const E = require('../../src/security/engine.js');
(async () => {
  const src = new ExcelJS.Workbook(); await src.xlsx.load(fs.readFileSync(process.argv[2]));
  const raw = src.getWorksheet('Raw'); const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet1');
  raw.eachRow({ includeEmpty: true }, (row, r) => { if (r < 6) return; row.eachCell({ includeEmpty: false }, (c, col) => { ws.getCell(r, col).value = c.type === ExcelJS.ValueType.Formula ? c.model.result : c.value; }); });
  const ex = E._internals.readExtract(wb);
  const kept = ex.rows.filter((r) => ['iivf p2p', 'payables'].includes(r.source.trim().toLowerCase()) || r.createdBy.trim().toLowerCase() === 'taxverification2@indiraivf.in');
  const fw = src.getWorksheet('Working'); const hdr = fw.getRow(2).values; const ci = (n) => hdr.indexOf(n);
  const fin = []; for (let r = 3; r <= fw.rowCount; r++) { const row = fw.getRow(r); const v = (n) => { const c = row.getCell(ci(n)); return c.type === ExcelJS.ValueType.Formula ? c.model.result : c.value; }; if (!v('SEGMENT1')) continue; fin.push({ seg: String(v('SEGMENT1')), desc: String(v('LINE_DESCRIPTION') || '').trim(), bal: +v('BALANCE') }); }
  const used = new Set(); const extra = [];
  for (const r of kept) { const i = fin.findIndex((f, k) => !used.has(k) && f.seg === String(r.seg1) && f.desc === r.desc.trim() && Math.abs(f.bal - r.bal) < 0.005); if (i < 0) extra.push(r); else used.add(i); }
  const g = {}; for (const r of extra) { const k = r.source + ' / ' + r.createdBy; g[k] = g[k] || { n: 0, amt: 0 }; g[k].n++; g[k].amt += r.bal; }
  console.log(g);
  for (const r of extra) console.log(r.erpRow, r.seg1, r.ccName, r.bal, r.source, r.createdBy, '|', r.desc.replace(/https?:\/\/\S+/g, '<url>').slice(0, 60));
  // Is the finance Booking/Pivot equal to finance Working?
})();
