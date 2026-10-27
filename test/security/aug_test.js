// Generality test on a different month: Aug-26 raw extract (from finance's Aug file "Raw" sheet)
// Compares the tool's Steps 7-8 filter and Step 6 tags with finance's own Aug-26 Working "FY Bucket".
const fs = require('fs'); const ExcelJS = require('exceljs'); const E = require('../../src/security/engine.js');
(async () => {
  const src = new ExcelJS.Workbook(); await src.xlsx.load(fs.readFileSync(process.argv[2]));
  const raw = src.getWorksheet('Raw');
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet1');
  raw.eachRow({ includeEmpty: true }, (row, r) => { if (r < 6) return; row.eachCell({ includeEmpty: false }, (c, col) => { ws.getCell(r, col).value = c.type === ExcelJS.ValueType.Formula ? c.model.result : c.value; }); });
  const ex = E._internals.readExtract(wb);
  const period = ex.period; console.log('period', period);
  const P2P = ['iivf p2p', 'payables'];
  const kept = ex.rows.filter((r) => P2P.includes(r.source.trim().toLowerCase()) || r.createdBy.trim().toLowerCase() === 'taxverification2@indiraivf.in');
  // finance tags
  const fw = src.getWorksheet('Working'); const hdr = fw.getRow(2).values;
  const ci = (n) => hdr.indexOf(n);
  const fin = [];
  for (let r = 3; r <= fw.rowCount; r++) { const row = fw.getRow(r); const v = (n) => { const c = row.getCell(ci(n)); return c.type === ExcelJS.ValueType.Formula ? c.model.result : c.value; };
    if (!v('SEGMENT1')) continue; fin.push({ seg: String(v('SEGMENT1')), desc: String(v('LINE_DESCRIPTION') || '').trim(), bal: +v('BALANCE'), tag: v('FY Bucket'), src: v('JE_SOURCE_NAME'), by: v('CREATED_BY') }); }
  console.log('tool kept', kept.length, 'finance working lines', fin.length);
  let agree = 0, disagree = 0, unsure = 0, notFound = 0; const used = new Set();
  for (const r of kept) {
    const p = E.proposeFY(r.desc, period);
    const i = fin.findIndex((f, k) => !used.has(k) && f.seg === String(r.seg1) && f.desc === r.desc.trim() && Math.abs(f.bal - r.bal) < 0.005);
    if (i < 0) { notFound++; continue; } used.add(i);
    const ft = fin[i].tag; const tt = p.fy === null ? null : E.fyLabelWorking(p.fy);
    if (!p.sure) { unsure++; console.log('  CONFIRM', r.erpRow, 'suggest', tt, 'finance', ft, '|', p.basis.slice(0, 60), '|', r.desc.replace(/https?:\/\/\S+/g, '<url>').slice(0, 70)); continue; }
    if (tt === ft) agree++; else { disagree++; console.log('  DISAGREE', r.erpRow, 'tool', tt, 'finance', ft, '|', p.basis, '|', r.desc.slice(0, 100)); }
  }
  console.log({ autoTaggedAgree: agree, autoTaggedDisagree: disagree, needsConfirm: unsure, keptNotInFinanceWorking: notFound });
  const finOnly = fin.filter((f, k) => !used.has(k));
  console.log('finance working lines not kept by tool', finOnly.length, finOnly.slice(0, 10).map((f) => [f.seg, f.bal, f.src, f.by, f.tag, f.desc.slice(0, 50)]));
})();
// totals: tool kept vs finance working, by finance tag where matched
(async () => {
  const src = new ExcelJS.Workbook(); await src.xlsx.load(fs.readFileSync(process.argv[2]));
  const raw = src.getWorksheet('Raw'); const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet1');
  raw.eachRow({ includeEmpty: true }, (row, r) => { if (r < 6) return; row.eachCell({ includeEmpty: false }, (c, col) => { ws.getCell(r, col).value = c.type === ExcelJS.ValueType.Formula ? c.model.result : c.value; }); });
  const ex = E._internals.readExtract(wb);
  const kept = ex.rows.filter((r) => ['iivf p2p', 'payables'].includes(r.source.trim().toLowerCase()) || r.createdBy.trim().toLowerCase() === 'taxverification2@indiraivf.in');
  const tot = kept.reduce((a, r) => a + Math.round(r.bal * 100), 0) / 100;
  console.log('\nTOTALS tool kept', tot.toFixed(2), '| compare with finance Aug Booking FY26 + FY27');
  const sources = {}; for (const r of kept) { const k = r.source + ' / ' + r.createdBy; sources[k] = (sources[k] || 0) + 1; } console.log(sources);
  for (const d of ['BEING MARCH MONTH INVOICE PARKED AS PER INVOICE NUMBER SS/26-27/56', 'BEING APRIL MONTH INVOICE PARKED']) { const p = E.proposeFY(d, ex.period); console.log(d.slice(0, 30), '→', p); }
})();
