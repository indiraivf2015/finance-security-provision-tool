// Proves the shared core on real data: Security's Sep-26 files are run through the core with Security's settings,
// and every Summary figure (cost centres, opening, bookings, carry-forward, Rajeev, YTD, totals) must equal
// finance's approved Security output (test-data/golden/security_sep26.xlsx). This does not touch Security itself.
// Usage: node test/core/security_parity.js
'use strict';
const fs = require('fs'); const path = require('path');
const ExcelJS = require('exceljs');
const { createEngine } = require('../../src/core/engine');
const fyStatement = require('../../src/core/fy-statement');
const { fyLabelHeader } = require('../../src/core/common');
const { FMT_AMT } = require('../../src/core/excel');
const { cached } = require('../xlsx_values');

const ROOT = path.join(__dirname, '..', '..');
const TD = (f) => path.join(ROOT, 'test-data', f);
const need = ['extract_Sep-26.xlsx', 'statement_Aug-26.xlsx', 'provision_Sep-26.xlsx', 'golden/security_sep26.xlsx'];
const missing = need.filter((f) => !fs.existsSync(TD(f)));
if (missing.length) { console.log('SKIPPED: sample files missing in test-data/: ' + missing.join(', ')); process.exit(0); }

const cfg = { toolName: 'core (Security settings)', title: (L) => `Security Provision for the month of ${L}`, pivotByVendor: false, amountText: 'Provision with GST',
  gst: { divisor: 118, basicFactor: 100, finalFactor: 1.135 }, negativeCarryForward: { topN: 10 }, summaryBlock: false, amountFormat: FMT_AMT };
const E = createEngine({
  toolName: cfg.toolName, version: 'test', sopVersion: 'Security SOP v1.0',
  extract: { accounts: [51701, 51702, 51703, 51704], accountsText: '51701–51704', glParams: { from: 51701, to: 51704 }, multipleAccountsAllowed: false },
  opening: (period) => ({ columns: [{ key: 'ytd', header: `YTD ${fyLabelHeader(period.cur)}`, choosable: true }, { key: 'cf', header: `${fyLabelHeader(period.prev)} Pro CF` }], aprilNote: true }),
  provisionFile: { sheet: { test: /security/i }, amount: { startsWith: 'provisionwithgst', text: 'Provision with GST' }, noun: 'security' },
  compute: (s) => fyStatement.compute(s, cfg), buildWorkbook: (s, R, m) => fyStatement.buildWorkbook(s, R, m, cfg, 'test'),
});

(async () => {
  const wbs = await E.load({ extract: fs.readFileSync(TD('extract_Sep-26.xlsx')), opening: fs.readFileSync(TD('statement_Aug-26.xlsx')), rajeev: fs.readFileSync(TD('provision_Sep-26.xlsx')) });
  const st = E.analyse(wbs, { openingYtdCol: 'M' });
  if (st.needChoice) throw new Error('needs choice ' + JSON.stringify(st.needChoice));
  const { buffer, result } = await E.build(st, new Map(), { now: new Date(0) });
  fs.mkdirSync(path.join(ROOT, 'out'), { recursive: true }); fs.writeFileSync(path.join(ROOT, 'out/core_security_sep26.xlsx'), Buffer.from(buffer));

  const read = async (f) => { const w = new ExcelJS.Workbook(); await w.xlsx.load(f); return w.getWorksheet('Summary'); };
  const g = await read(fs.readFileSync(TD('golden/security_sep26.xlsx'))); const c = await read(Buffer.from(buffer));
  let cells = 0, bad = 0;
  const rows = Math.max(g.rowCount, c.rowCount);
  for (let r = 2; r <= rows; r++) for (let col = 2; col <= 11; col++) {
    if (r === 2 && col > 2) continue;
    const a = cached(g.getRow(r).getCell(col)), b = cached(c.getRow(r).getCell(col));
    cells++;
    const same = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) < 0.005 : a === b;
    if (!same) { bad++; if (bad <= 20) console.log(`  DIFF Summary!${c.getRow(r).getCell(col).address}: approved ${JSON.stringify(a)} | core ${JSON.stringify(b)}`); }
  }
  console.log(`Summary figures compared: ${cells}, different: ${bad}`);
  const failed = result.checks.filter((x) => x.status === 'FAIL');
  console.log(`core checks: ${result.checks.length - failed.length}/${result.checks.length} pass; status ${result.status}; YTD ${result.T.ytd.toFixed(2)}`);
  console.log(bad || failed.length ? 'CORE PARITY FAILED' : 'CORE PARITY OK: the core reproduces finance\'s approved Security statement');
  process.exit(bad || failed.length ? 1 : 0);
})().catch((e) => { console.error(e.list ? e.list.join('\n') : e.stack); process.exit(1); });
