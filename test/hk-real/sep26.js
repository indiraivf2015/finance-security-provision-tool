// Housekeeping on finance's REAL Sep-26 files (test-data/hk/, git-ignored; skipped when missing):
// the true ERP download (one combined report, accounts 51610–51614), finance's original Aug-26 statements, Rajeev's file.
//   - both tools run on the same combined file, each skipping the other's account and listing it;
//   - Expenses: both Mumbai_Corporate (2003) rows equal finance's Sep-26 statement on every column;
//   - the tool's own Sep-26 Expenses download is read as October's opening with both 2003 rows carried forward.
// Usage: node test/hk-real/sep26.js
'use strict';
const fs = require('fs'); const path = require('path');
const ExcelJS = require('exceljs');
const EXP = require('../../src/provisions/hk-expenses');
const MAT = require('../../src/provisions/hk-material');
const F = require('../fixtures');
const { cached } = require('../xlsx_values');

const TD = (f) => path.join(__dirname, '..', '..', 'test-data', f);
const need = ['hk/ERP_Sep-26_51610-51614.xlsx', 'hk/Housekeeping Exp_Aug-26.xlsx', 'hk/Material_Aug-26.xlsx', 'hk/Housekeeping Exp_Sep26.xlsx', 'provision_Sep-26.xlsx'];
const missing = need.filter((f) => !fs.existsSync(TD(f)));
if (missing.length) { console.log('SKIPPED: finance files missing in test-data/: ' + missing.join(', ')); process.exit(0); }
const near = (a, b) => Math.abs(a - b) < 0.005;
const val = (c) => { const v = cached(c); return typeof v === 'number' ? v : (v == null || v === '' ? 0 : +v); };

async function run(E, opening, opts) {
  const st = E.analyse(await E.load({ extract: fs.readFileSync(TD('hk/ERP_Sep-26_51610-51614.xlsx')), opening, rajeev: fs.readFileSync(TD('provision_Sep-26.xlsx')) }), opts || {});
  if (st.needChoice) throw new Error('needs choice ' + JSON.stringify(st.needChoice));
  const { buffer, result } = await E.build(st, new Map(), { now: new Date(0) });
  return { st, R: result, buffer: Buffer.from(buffer) };
}
const summaryRows = async (buf, code) => { const w = new ExcelJS.Workbook(); await w.xlsx.load(buf); const ws = w.getWorksheet('Summary'); const out = [];
  ws.eachRow((r, i) => { if (i > 4 && +cached(r.getCell('B')) === code) out.push(['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'].map((c) => val(r.getCell(c)))); }); return out; };

const tests = [];
const test = (name, fn) => tests.push({ name, fn });
let sepExp;

test('Expenses: READY on the combined download; 51611 lines skipped and listed', async () => {
  sepExp = await run(EXP, fs.readFileSync(TD('hk/Housekeeping Exp_Aug-26.xlsx')), { confirmOpeningMonth: true });
  const p = sepExp.R.points.find((x) => /combined report\. \d+ line\(s\) of account 51611/.test(x.point));
  return sepExp.R.status !== 'NOT READY' && sepExp.R.checks.every((c) => c.status !== 'FAIL') && p ? true : `status ${sepExp.R.status}, point ${!!p}`;
});
// Columns D..K = FY27 Pro, FY26 Pro, FY27 Booking, FY26 Booking, FY27 Pro CF, FY26 Pro CF, Rajeev, YTD FY27.
// The manual row equals finance's on every column. The normal row too, except that on the true download Kharagpur 1191's
// FY27 CF is negative and the negative carry-forward rule takes a share from 2003's normal row (one of the 10 highest);
// FY27 Pro CF and YTD then differ from finance by exactly that adjustment. The manual row is never part of the rule.
test('Expenses: both Mumbai_Corporate (2003) rows equal finance\'s Sep-26 statement (normal row apart from the negative-rule share)', async () => {
  const fin = await summaryRows(fs.readFileSync(TD('hk/Housekeeping Exp_Sep26.xlsx')), 2003); const tool = await summaryRows(sepExp.buffer, 2003);
  const [n, m] = sepExp.R.stmt.filter((x) => x.code === 2003);
  const adj = [0, 0, 0, 0, n.adjCur, n.adjPrev, 0, n.adjCur];
  const ok = fin.length === 2 && tool.length === 2 && fin[0].every((x, j) => near(x + adj[j], tool[0][j])) && fin[1].every((x, j) => near(x, tool[1][j])) && m.manual && near(m.adjCur + m.adjPrev, 0);
  console.log(`     2003 normal row: negative-rule share FY27 ${n.adjCur.toFixed(2)}, FY26 ${n.adjPrev.toFixed(2)}; manual row identical to finance's`);
  return ok ? true : `finance ${JSON.stringify(fin)} tool ${JSON.stringify(tool)} adj ${adj}`;
});
test('Expenses: the tool\'s own Sep-26 download as October\'s opening carries both 2003 rows forward', async () => {
  const st = EXP.analyse(await EXP.load({ extract: await F.extract({ period: 'Oct-26', from: '2026-10-01', to: '2026-10-31', glFrom: '51610', glTo: '51614', lines: [{ code: 1001, name: 'Udaipur', account: 51610, bal: 10, desc: 'for the month of September 2026' }] }),
    opening: sepExp.buffer, rajeev: await F.hkFile({ sheet: 'HK Oct-26', rows: [{ code: 1001, name: 'Udaipur', manpower: 118000 }, { code: 2003, name: 'Mumbai_Corporate', manpower: 118000 }] }) }), {});
  if (st.needChoice) return 'needs choice ' + st.needChoice.kind;
  const { result: R } = await EXP.build(st, new Map(), { now: new Date(0) });
  const sepRows = sepExp.R.stmt.filter((x) => x.code === 2003); const octRows = R.stmt.filter((x) => x.code === 2003);
  const ok = octRows.length === 2 && near(octRows[0].proCur, sepRows[0].ytd) && near(octRows[0].proPrev, sepRows[0].cfPrev) && octRows[1].manual && near(octRows[1].proCur, 5000000) && near(octRows[1].ytd, 5000000)
    && near(R.T.proCur, sepExp.R.T.ytd) && near(R.T.proPrev, sepExp.R.T.cfPrev);
  return ok ? true : `Oct 2003 rows ${JSON.stringify(octRows.map((x) => [x.proCur, x.proPrev, x.ytd, !!x.manual]))}`;
});
test('Material: READY on the same combined download; other accounts skipped and listed', async () => {
  const { R } = await run(MAT, fs.readFileSync(TD('hk/Material_Aug-26.xlsx')), { confirmOpeningMonth: true });
  const p = R.points.find((x) => /combined report\. .*account 51610/.test(x.point));
  return R.status !== 'NOT READY' && R.checks.every((c) => c.status !== 'FAIL') && p ? true : `status ${R.status}, point ${!!p}`;
});

(async () => {
  let bad = 0;
  for (const t of tests) {
    let r; try { r = await t.fn(); } catch (e) { r = 'ERROR ' + (e.list ? e.list.join(' | ') : e.stack); }
    if (r !== true) bad++;
    console.log(`${r === true ? 'ok  ' : 'FAIL'} ${t.name}${r === true ? '' : ' — ' + r}`);
  }
  console.log(`\n${tests.length - bad} passed, ${bad} failed`);
  process.exit(bad ? 1 : 0);
})();
