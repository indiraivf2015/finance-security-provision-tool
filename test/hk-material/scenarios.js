// Housekeeping Material: SOP rules and problem cases on MADE-UP data (no finance data needed).
// The negative carry-forward setting is tested both ways (SOP as written, and finance's rule).
// Usage: node test/hk-material/scenarios.js
'use strict';
const fs = require('fs'); const path = require('path');
const ExcelJS = require('exceljs');
const SOPWAY = require('../../src/provisions/hk-material').withNegativeCfRule(null);
const RULEWAY = require('../../src/provisions/hk-material').withNegativeCfRule({ topN: 10 });
const F = require('../fixtures');
const { cached } = require('../xlsx_values');

const OUT = path.join(__dirname, '..', '..', 'out'); fs.mkdirSync(OUT, { recursive: true });
const near = (a, b) => Math.abs(a - b) < 0.005;

const CODES = Array.from({ length: 12 }, (_, i) => 8001 + i);
const NAME = (c) => 'Centre ' + c;
const OPEN = (c) => 10000 * (c - 8000);
const OPEN_HDR = ['Cost Center', 'Cost Center Name', 'Jul-Provision', 'Aug-Booking', 'Prov CF', 'Aug-Provision', 'YTD'];
const openRows = (codes) => codes.map((c) => [c, NAME(c), 0, 0, 0, 0, OPEN(c)]);
const L = (code, bal, desc, extra) => ({ code, name: NAME(code), account: 51611, bal, desc, ...(extra || {}) });
const BASE_LINES = [
  L(8001, 2500, 'Housekeeping material for the month of August 2026', { attr3: 'Vendor A' }),
  L(8001, 450, 'GST reclass for the month of August 2026', { source: 'Spreadsheet', by: 'taxverification2@indiraivf.in' }),
  L(8002, 3000, 'Housekeeping material for the month of August 2026', { source: 'Payables' }),
  L(8003, 999, 'Material for the month of March 2026'),            // FY 26: left out of the statement
  L(8004, 700, 'Housekeeping material invoice'),                     // no year: current FY by default
  L(8005, 80000, 'Housekeeping material for the month of August 2026'), // more than opening 50000: negative Prov CF
  L(8006, 111, 'Manual entry for the month of August 2026', { source: 'Spreadsheet', by: 'someone@indiraivf.in' }),
];
const HK_ROWS = (o) => CODES.map((c) => ({ code: c, name: NAME(c), manpower: (o && o.manpower) || 118000, materials: 1000.25 * (c - 8000), rent: 5900 }))
  .concat([{ code: 8002, name: NAME(8002) + ' (store boy)', manpower: 30000, materials: 333.5 }]);

async function files(o) {
  o = o || {};
  return {
    extract: await F.extract({ period: o.period || 'Sep-26', from: o.from || '2026-09-01', to: o.to || '2026-09-30', glFrom: o.glFrom || '51611', glTo: o.glTo || '51611', multiple: o.multiple, lines: o.lines || BASE_LINES }),
    opening: o.openingBuf || await F.statement({ title: o.openTitle || 'Material Provision for the month of Aug-26', headers: o.openHdr || OPEN_HDR, rows: o.openRows || openRows(CODES) }),
    rajeev: o.rajeevBuf || await F.hkFile({ sheet: o.sheet || 'HK Sep-26', rows: o.hkRows || HK_ROWS(), extraSheets: ['Security Sep-26'] }),
  };
}
async function run(o, opts, E) {
  E = E || SOPWAY;
  const st = E.analyse(await E.load(await files(o)), opts || {});
  if (st.needChoice) return { needChoice: st.needChoice };
  const { buffer, result } = await E.build(st, new Map(), { now: new Date(0) });
  return { st, R: result, buffer: Buffer.from(buffer) };
}
async function expectError(o, re) {
  try { await run(o); return 'no error'; } catch (e) { const msg = e.list ? e.list.join(' | ') : e.message; return re.test(msg) ? true : 'wrong message: ' + msg; }
}
const sheetOf = async (buf, name) => { const w = new ExcelJS.Workbook(); await w.xlsx.load(buf); return w.getWorksheet(name); };
const row = (R, c) => R.stmt.find((x) => x.code === c);

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('base month (SOP way): READY, all 10 checks pass', async () => {
  const { R } = await run();
  const fails = R.checks.filter((c) => c.status === 'FAIL').map((c) => c.text);
  return !fails.length && R.checks.length === 10 && R.status !== 'NOT READY' ? true : `${R.status}, ${R.checks.length} checks, failed: ${fails.join('; ')}`;
});
test('Step 2: a 51610 (Expenses) line stops the tool', () => expectError({ lines: BASE_LINES.concat([L(8001, 1, 'x', { account: 51610 })]) }, /51610.*account 51611/));
test('Step 2: report "GL Code From" other than 51611 stops the tool', () => expectError({ glFrom: '51610' }, /GL Code From" is 51610, SOP says 51611/));
test('Step 2: "Multiple Accounts" filled in stops the tool', () => expectError({ multiple: '51611,51612' }, /Multiple Accounts/));
test('Steps 7-8: IIVF P2P, Payables and GST reclass kept; other Spreadsheet removed', async () => {
  const { st } = await run(); return st.removed.length === 1 && st.removed[0].bal === 111 && st.kept.length === 6 ? true : `kept ${st.kept.length} removed ${st.removed.map((r) => r.bal)}`;
});
test('Step 8 (Housekeeping): a taxverification2 line is kept whatever its source says; the same source by anyone else is removed', async () => {
  const tv = 'taxverification2@indiraivf.in';
  const { st } = await run({ lines: BASE_LINES.concat([L(8001, 701, 'GST reclass for the month of August 2026', { source: 'GST Reclass', by: tv }), L(8001, 702, 'GST reclass for the month of August 2026', { source: 'Manual', by: tv }), L(8001, 703, 'Relabelled for the month of August 2026', { source: 'GST Reclass', by: 'someone@indiraivf.in' })]) });
  const k = (b) => st.kept.some((r) => r.bal === b), x = (b) => st.removed.some((r) => r.bal === b);
  return k(701) && k(702) && x(703) && st.kept.filter((r) => r.bal > 700 && r.bal < 703).every((r) => r.keptAs === 'GST reclass (Step 8)') ? true : `kept ${st.kept.map((r) => r.bal)} removed ${st.removed.map((r) => r.bal)}`;
});
test('Step 10: FY 26 line stays in Working, is left out of the statement and listed in Points to check', async () => {
  const { R, st, buffer } = await run();
  const inWorking = st.kept.some((r) => r.bal === 999 && r.fy === 26);
  const pivot = await sheetOf(buffer, 'Pivot'); const pivotHas = R.pivot.some((p) => p.code === 8003);
  const pt = R.points.find((p) => p.code === 8003 && /FY 26 line left out/.test(p.point));
  const filter = cached(pivot.getCell('A1')) === 'Exp Year' && cached(pivot.getCell('B1')) === 'FY 27';
  return inWorking && !pivotHas && pt && near(pt.amount, 999) && near(row(R, 8003).booking, 0) && filter ? true : JSON.stringify({ inWorking, pivotHas, pt, filter });
});
test('SOP check 3 with FY 26 lines present: pivot total = FY 27 Working total, Booking total = pivot total', async () => {
  const { R } = await run(); const c = R.checks.filter((x) => x.sop === '04 · check 3');
  return c.length === 2 && c.every((x) => x.status === 'PASS') && near(c[0].actual, 2500 + 450 + 3000 + 700 + 80000) ? true : JSON.stringify(c);
});
test('Finance rule: line with no year counted as FY 27 and listed', async () => {
  const { R } = await run(); return near(row(R, 8004).booking, 700) && R.points.some((p) => /1 Working line\(s\) had no year/.test(p.point)) ? true : 'not counted';
});
test('Step 11: Materials taken as is (no GST conversion); rows with the same code added together', async () => {
  const { R } = await run(); return near(row(R, 8001).prov, 1000.25) && near(row(R, 8002).prov, 2000.5 + 333.5) ? true : `${row(R, 8001).prov} / ${row(R, 8002).prov}`;
});
test('Manpower and Machine Rent are not used (changing Manpower changes nothing)', async () => {
  const a = await run(); const b = await run({ hkRows: HK_ROWS({ manpower: 999999 }) });
  const sa = await sheetOf(a.buffer, 'Summary'); const sb = await sheetOf(b.buffer, 'Summary');
  for (let r = 1; r <= sa.rowCount; r++) for (let c = 1; c <= 10; c++) if (JSON.stringify(cached(sa.getRow(r).getCell(c))) !== JSON.stringify(cached(sb.getRow(r).getCell(c)))) return `Summary ${sa.getRow(r).getCell(c).address} changed`;
  return true;
});
test('Steps 12-15: opening = last month\'s YTD; Prov CF = Provision − Booking; YTD = Prov CF + Provision', async () => {
  const { R } = await run(); const s = row(R, 8001);
  return near(s.open, OPEN(8001)) && near(s.booking, 2950) && near(s.cf, OPEN(8001) - 2950) && near(s.ytd, s.cf + 1000.25) ? true : JSON.stringify(s);
});
test('SOP way: a negative Prov CF is carried with its minus sign and listed', async () => {
  const { R } = await run(); const s = row(R, 8005);
  return near(s.cf, 50000 - 80000) && near(s.ytd, -30000 + 5001.25) && R.points.some((p) => p.code === 8005 && /minus sign/.test(p.point)) && !R.adjInfo ? true : JSON.stringify(s);
});
test('Rule way: negative Prov CF set to 0, taken from the top 10 in proportion, total unchanged; 12 checks', async () => {
  const { R } = await run(null, null, RULEWAY); const s = row(R, 8005);
  const before = R.stmt.reduce((a, x) => a + x.cfSop, 0), after = R.stmt.reduce((a, x) => a + x.cf, 0);
  const fails = R.checks.filter((c) => c.status === 'FAIL');
  return near(s.cf, 0) && R.adjInfo.adj.recipients.length === 10 && near(before, after) && R.checks.length === 12 && !fails.length ? true : JSON.stringify({ cf: s.cf, before, after, n: R.checks.length, fails });
});
test('The page uses finance\'s negative carry-forward rule (owner\'s decision, Oct 2026)', async () => {
  const PAGE = require('../../src/provisions/hk-material');
  const { R } = await run(null, null, PAGE);
  return PAGE.NEGATIVE_CF_RULE && PAGE.NEGATIVE_CF_RULE.topN === 10 && R.adjInfo && R.adjInfo.adj.applied && near(row(R, 8005).cf, 0) ? true : JSON.stringify(PAGE.NEGATIVE_CF_RULE);
});
test('Statement title, the SOP\'s 7 columns, amounts shown without paise but kept exact', async () => {
  const { buffer } = await run(); const ws = await sheetOf(buffer, 'Summary');
  const t = cached(ws.getCell('A2')); const h = ws.getRow(4).values.slice(2).join('|');
  const exact = ws.getRow(5).getCell(7); // first row's Sept-Provision
  const fmt = ['D', 'E', 'F', 'G', 'H'].every((c) => ws.getColumn(c).numFmt === '#,##0;-#,##0;"-"');
  return t === 'Material Provision for the month of Sept-26' && h === 'Cost Center|Cost Center Name|Aug-Provision|Sept-Booking|Prov CF|Sept-Provision|YTD' && fmt && cached(exact) % 1 !== 0 ? true : `${t} / ${h} / fmt ${fmt} / ${cached(exact)}`;
});
test('Next month: this output is read as next month\'s opening with no question asked', async () => {
  const { buffer, R } = await run();
  const E = SOPWAY;
  const st = E.analyse(await E.load({ extract: await F.extract({ period: 'Oct-26', from: '2026-10-01', to: '2026-10-31', glFrom: '51611', glTo: '51611', lines: [L(8001, 10, 'for the month of September 2026')] }), opening: buffer, rajeev: await F.hkFile({ sheet: 'HK Oct-26', rows: HK_ROWS() }) }), {});
  if (st.needChoice) return 'asked: ' + JSON.stringify(st.needChoice);
  return near(st.opening.centres.get(8005).v.ytd, row(R, 8005).ytd) ? true : 'wrong opening';
});
test('Opening: statement for the wrong month stops the tool', () => expectError({ openTitle: 'Material Provision for the month of Jul-26' }, /titled .*Jul-26/));
test('Opening: two "YTD" columns → the user picks one', async () => {
  const r = await run({ openHdr: OPEN_HDR.concat(['YTD']), openRows: openRows(CODES).map((x) => x.concat([1])) });
  return r.needChoice && r.needChoice.kind === 'openingYtdCol' && r.needChoice.header === 'YTD' ? true : JSON.stringify(r.needChoice || 'no question');
});
test('Opening: a centre on two rows of last month\'s statement is added together, listed, totals unchanged', async () => {
  const rows = openRows(CODES).concat([[CODES[2], NAME(CODES[2]), 0, 0, 0, 0, 5000000]]);
  const a = await run(); const b = await run({ openRows: rows });
  const s = b.R.stmt.find((x) => x.code === CODES[2]); const s0 = a.R.stmt.find((x) => x.code === CODES[2]);
  const p = b.R.points.find((x) => x.code === CODES[2] && /appears 2 times in last month's statement \(₹[\d,]+\.\d\d \+ ₹50,00,000\.00\); rows added together\./.test(x.point));
  return s && near(s.open, s0.open + 5000000) && near(b.R.T.open, a.R.T.open + 5000000) && b.R.stmt.filter((x) => x.code === CODES[2]).length === 1 && p ? true : `open ${s && s.open} vs ${s0.open}, point ${!!p}`;
});
test('Provision file: sheet for another month stops the tool', () => expectError({ sheet: 'HK Aug-26' }, /sheet used is "HK Aug-26"/));
test('April: stops, the SOP has no year-end roll-over rule', () => expectError({ period: 'Apr-27', from: '2027-04-01', to: '2027-04-30', sheet: 'HK Apr-27', openTitle: 'Material Provision for the month of Mar-27', lines: [L(8001, 10, 'for the month of April 2027')] }, /April.*year-end/));
test('Points to check sheet lists the FY 26 line; no adjustment table in the SOP way', async () => {
  const { buffer } = await run(); const ws = await sheetOf(buffer, 'Points to check'); let text = '';
  ws.eachRow((r) => r.eachCell((c) => { text += ' ' + cached(c); }));
  return /FY 26 line left out/.test(text) && !/Negative carry-forward adjustment/.test(text) ? true : 'unexpected content';
});

(async () => {
  let bad = 0;
  for (const t of tests) {
    let r; try { r = await t.fn(); } catch (e) { r = 'ERROR ' + (e.list ? e.list.join(' | ') : e.stack); }
    if (r !== true) bad++;
    console.log(`${r === true ? 'ok  ' : 'FAIL'} ${t.name}${r === true ? '' : ' — ' + r}`);
  }
  fs.writeFileSync(path.join(OUT, 'hk_material_sample.xlsx'), (await run()).buffer);
  console.log(`\n${tests.length - bad} passed, ${bad} failed`);
  process.exit(bad ? 1 : 0);
})();
