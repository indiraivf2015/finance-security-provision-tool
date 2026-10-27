// Housekeeping Expenses: SOP rules and problem cases on MADE-UP data (no finance data needed).
// Usage: node test/hk-expenses/scenarios.js
'use strict';
const fs = require('fs'); const path = require('path');
const ExcelJS = require('exceljs');
const E = require('../../src/provisions/hk-expenses');
const F = require('../fixtures');
const { cached } = require('../xlsx_values');

const OUT = path.join(__dirname, '..', '..', 'out'); fs.mkdirSync(OUT, { recursive: true });
const near = (a, b) => Math.abs(a - b) < 0.005;
const final = (manpower) => manpower / 118 * 100 * 1.135;

// ---------- base month: Sep-26, opening = Aug-26 statement ----------
const CODES = Array.from({ length: 12 }, (_, i) => 9001 + i);
const NAME = (c) => 'Centre ' + c;
const OPEN_YTD = (c) => (c === 9006 ? 600000 : 100000 * (c - 9000));
const OPEN_CF = () => 5000;
const OPEN_HDR = ['Cost Center', 'Cost Center Name', 'FY27 Pro', 'FY26 Pro', 'FY27 Booking', 'FY26 Booking', 'FY27 Pro CF', 'FY26 Pro CF', 'Aug-26 Pro as per Rajeev', 'YTD FY27'];
const openRows = (codes) => codes.map((c) => [c, NAME(c), 0, 0, 0, 0, 0, OPEN_CF(c), 0, OPEN_YTD(c)]);
const L = (code, bal, desc, extra) => ({ code, name: NAME(code), account: 51610, bal, desc, ...(extra || {}) });
const BASE_LINES = [
  L(9001, 50000, 'Housekeeping services for the month of August 2026', { attr3: 'Alpha Services' }),
  L(9001, 1800, 'GST reclass for the month of August 2026', { source: 'Spreadsheet', by: 'taxverification2@indiraivf.in', attr3: 'Alpha Services' }),
  L(9001, 999, 'Manual adjustment for the month of August 2026', { source: 'Spreadsheet', by: 'someone@indiraivf.in' }),
  L(9002, 30000, 'Housekeeping for the month of August 2026', { account: 51612, attr3: 'Alpha Services', source: 'Payables' }),
  L(9002, 20000, 'Housekeeping for the month of Jul-26', { account: 51613, attr3: 'Beta Facility' }),
  L(9002, 555, 'Other source line for the month of August 2026', { source: 'Manual' }),
  L(9003, 4000, 'Housekeeping for the month of March 2026', { account: 51614 }),
  L(9004, 1000, 'Housekeeping invoice', {}),
  L(9005, -7000, 'Reversal of housekeeping for the month of June 2026', {}),
  L(9006, 700000, 'Housekeeping for the month of August 2026', {}),
  L(9007, 2000, 'Housekeeping for the month of August 2026', { attr3: '', party: 'Gamma Facility' }),
];
const HK_ROWS = (opts) => CODES.map((c) => ({ code: c, name: NAME(c), manpower: 118000, materials: (opts && opts.materials) || 5000, rent: (opts && opts.rent) || 5900 }))
  .concat([{ code: 9002, name: NAME(9002) + ' OPD', manpower: 23600, materials: 999, type: 'OPD' }]);

async function files(o) {
  o = o || {};
  return {
    extract: await F.extract({ period: o.period || 'Sep-26', from: o.from || '2026-09-01', to: o.to || '2026-09-30', glFrom: '51610', glTo: '51614', lines: o.lines || BASE_LINES }),
    opening: o.openingBuf || await F.statement({ title: o.openTitle || 'Housekeeping Exp – Provision for the month of Aug-26', headers: o.openHdr || OPEN_HDR, rows: o.openRows || openRows(CODES) }),
    rajeev: o.rajeevBuf || await F.hkFile({ sheet: o.sheet || 'HK Sep-26', rows: o.hkRows || HK_ROWS(), extraSheets: o.extraSheets || ['Security Sep-26'] }),
  };
}
async function run(o, opts) {
  const wbs = await E.load(await files(o));
  const st = E.analyse(wbs, opts || {});
  if (st.needChoice) return { needChoice: st.needChoice };
  const { buffer, result } = await E.build(st, new Map(), { now: new Date(0) });
  return { st, R: result, buffer: Buffer.from(buffer) };
}
async function expectError(o, re) {
  try { await run(o); return 'no error'; } catch (e) { const msg = e.list ? e.list.join(' | ') : e.message; return re.test(msg) ? true : 'wrong message: ' + msg; }
}
const sheetOf = async (buf, name) => { const w = new ExcelJS.Workbook(); await w.xlsx.load(buf); return w.getWorksheet(name); };

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test('base month: READY, all checks pass, 17 checks (SOP 04 checks 1-6, completeness, negative-CF rule)', async () => {
  const { R } = await run();
  const fails = R.checks.filter((c) => c.status === 'FAIL').map((c) => c.text);
  return fails.length === 0 && R.checks.length === 17 && R.status !== 'NOT READY' ? true : `status ${R.status}, ${R.checks.length} checks, failed: ${fails.join('; ')}`;
});
test('Step 2: a 51611 (Material) line stops the tool and names the account', () => expectError({ lines: BASE_LINES.concat([L(9001, 10, 'x for the month of August 2026', { account: 51611 })]) }, /51611.*51610, 51612, 51613 and 51614/));
test('Step 2: a Security account (51701) line stops the tool', () => expectError({ lines: BASE_LINES.concat([L(9001, 10, 'x', { account: 51701 })]) }, /51701/));
test('Steps 7-8: IIVF P2P and Payables kept, GST reclass kept, other Spreadsheet and other sources removed', async () => {
  const { st } = await run();
  const kept = st.kept.map((r) => r.bal).sort((a, b) => a - b).join(','); const removed = st.removed.map((r) => r.bal).sort((a, b) => a - b).join(',');
  return removed === '555,999' && st.kept.length === 9 ? true : `kept ${kept} removed ${removed}`;
});
test('Step 9: blank ATTRIBUTE3 filled with a formula to PARTY_NAME in the Working sheet', async () => {
  const { buffer } = await run();
  const ws = await sheetOf(buffer, 'Working'); const hdr = ws.getRow(1).values; const a3 = hdr.indexOf('ATTRIBUTE3'), pn = hdr.indexOf('PARTY_NAME');
  for (let r = 2; r <= ws.rowCount; r++) { const c = ws.getRow(r).getCell(a3); if (c.formula) return c.formula === ws.getRow(r).getCell(pn).address && cached(c) === 'Gamma Facility' ? true : `formula ${c.formula} result ${cached(c)}`; }
  return 'no formula found';
});
test('Step 10: pivot rows by centre name, code and vendor; statement uses the centre total across vendors', async () => {
  const { R, buffer } = await run();
  const p = R.pivot.filter((x) => x.code === 9002);
  const ws = await sheetOf(buffer, 'Pivot');
  const hdrOk = JSON.stringify(ws.getRow(4).values.slice(1, 4)) === JSON.stringify(['BALANCING_SEGMENT_DESC', 'SEGMENT1', 'ATTRIBUTE3']);
  const s = R.stmt.find((x) => x.code === 9002);
  return p.length === 2 && hdrOk && near(s.bkCur, 50000) ? true : `pivot rows ${p.length}, header ok ${hdrOk}, booking ${s.bkCur}`;
});
test('Step 6 + finance rule: dated lines tagged from the narration; undated line counted as FY 27 and listed', async () => {
  const { R } = await run();
  const g = (c) => R.stmt.find((x) => x.code === c);
  const pt = R.points.some((x) => /1 Working line\(s\) had no year/.test(x.point));
  return near(g(9003).bkPrev, 4000) && near(g(9003).bkCur, 0) && near(g(9004).bkCur, 1000) && pt ? true : `9003 ${g(9003).bkPrev}/${g(9003).bkCur}, 9004 ${g(9004).bkCur}, point ${pt}`;
});
test('Steps 11-12: Manpower ÷118×100×1.135; Main + OPD rows of one code added together', async () => {
  const { R } = await run();
  const g = (c) => R.stmt.find((x) => x.code === c);
  return near(g(9001).raj, final(118000)) && near(g(9002).raj, final(118000) + final(23600)) ? true : `9001 ${g(9001).raj} 9002 ${g(9002).raj}`;
});
test('Materials and Machine Rent are not used anywhere (changing them changes nothing)', async () => {
  const a = await run(); const b = await run({ hkRows: HK_ROWS({ materials: 77777, rent: 88888 }) });
  const sa = await sheetOf(a.buffer, 'Summary'); const sb = await sheetOf(b.buffer, 'Summary');
  for (let r = 1; r <= sa.rowCount; r++) for (let c = 1; c <= 16; c++) if (JSON.stringify(cached(sa.getRow(r).getCell(c))) !== JSON.stringify(cached(sb.getRow(r).getCell(c)))) return `Summary ${sa.getRow(r).getCell(c).address} changed`;
  let text = ''; for (const n of ['Summary', 'Provision_Rajeev']) { const ws = await sheetOf(a.buffer, n); ws.eachRow((row) => row.eachCell((c) => { text += ' ' + String(cached(c)); })); }
  return /material|machine\s*rent/i.test(text) ? 'the words Materials / Machine Rent appear in the statement' : true;
});
test('Steps 13-16: Pro from last month, CF = Pro − Booking, YTD = FY27 Pro CF + Rajeev (reversal raises CF)', async () => {
  const { R } = await run();
  const s = R.stmt.find((x) => x.code === 9005);
  const ok = near(s.proCur, OPEN_YTD(9005)) && near(s.proPrev, 5000) && near(s.cfCurSop, OPEN_YTD(9005) + 7000) && near(s.ytd, s.cfCur + s.raj);
  return ok ? true : JSON.stringify(s);
});
test('Negative carry-forward rule: negative set to 0, taken from the top 10 in proportion, column total unchanged', async () => {
  const { R } = await run();
  const s = R.stmt.find((x) => x.code === 9006);
  const before = R.stmt.reduce((a, x) => a + x.cfCurSop, 0), after = R.stmt.reduce((a, x) => a + x.cfCur, 0);
  const inf = R.adjInfo.adjCur;
  return near(s.cfCur, 0) && s.cfCurSop < 0 && inf.applied && inf.recipients.length === 10 && near(before, after) && R.stmt.every((x) => x.cfCur > -0.005) ? true : JSON.stringify({ s, inf, before, after });
});
test('Negative CF larger than the top 10 can absorb → NOT READY, explained', async () => {
  const lines = BASE_LINES.filter((l) => l.code !== 9006).concat([L(9006, 99000000, 'Housekeeping for the month of August 2026')]);
  const { R } = await run({ lines });
  return R.status === 'NOT READY' && R.points.some((p) => /could NOT be absorbed/.test(p.point)) ? true : R.status;
});
test('Step 17 summary block: FY26 Pro CF + YTD FY27 = Total (and in millions); Every month = Total ÷ 4', async () => {
  const { R, buffer } = await run();
  const ws = await sheetOf(buffer, 'Summary');
  const labels = ['M5', 'M6', 'M7', 'M8'].map((a) => cached(ws.getCell(a))).join('|');
  const ok = labels === 'FY26 Pro CF|YTD FY27|Total|Every month' && near(cached(ws.getCell('N5')), R.T.cfPrev) && near(cached(ws.getCell('N6')), R.T.ytd)
    && near(cached(ws.getCell('N7')), R.T.cfPrev + R.T.ytd) && near(cached(ws.getCell('N8')), (R.T.cfPrev + R.T.ytd) / 4) && near(cached(ws.getCell('O7')), (R.T.cfPrev + R.T.ytd) / 1e6)
    && ws.getCell('N7').formula === 'N5+N6' && ws.getCell('N8').formula === 'N7/4';
  return ok ? true : labels;
});
test('Statement title and the 10 statement columns', async () => {
  const { buffer } = await run();
  const ws = await sheetOf(buffer, 'Summary');
  const t = cached(ws.getCell('A2')); const h = ws.getRow(4).values.slice(2).join('|');
  return t === 'Housekeeping Exp – Provision for the month of Sept-26' && h === 'Cost Center|Cost Center Name|FY27 Pro|FY26 Pro|FY27 Booking|FY26 Booking|FY27 Pro CF|FY26 Pro CF|Sept-26 Pro as per Rajeev|YTD FY27' ? true : `${t} / ${h}`;
});
test('Next month: this output is read as next month\'s opening with no question asked (summary block does not clash)', async () => {
  const { buffer, R } = await run();
  const wbs = await E.load({ extract: await F.extract({ period: 'Oct-26', from: '2026-10-01', to: '2026-10-31', lines: [L(9001, 100, 'for the month of September 2026')] }), opening: buffer, rajeev: await F.hkFile({ sheet: 'HK Oct-26', rows: HK_ROWS() }) });
  const st = E.analyse(wbs, {});
  if (st.needChoice) return 'asked: ' + JSON.stringify(st.needChoice);
  const o = st.opening.centres.get(9002); const s = R.stmt.find((x) => x.code === 9002);
  return near(o.v.ytd, s.ytd) && near(o.v.cf, s.cfPrev) ? true : JSON.stringify(o);
});
test('Provision file: the HK sheet is chosen when the file also has a Security sheet', async () => { const { st } = await run(); return st.rajeev.sheet === 'HK Sep-26' ? true : st.rajeev.sheet; });
test('Provision file: sheet for another month stops the tool', () => expectError({ sheet: 'HK Aug-26' }, /sheet used is "HK Aug-26".*Sep-26/));
test('Provision file: no HK sheet among several → the user is asked which sheet', async () => { const r = await run({ sheet: 'Sheet A', extraSheets: ['Security Sep-26'] }); return r.needChoice && r.needChoice.kind === 'rajeevSheet' ? true : JSON.stringify(r.needChoice || 'no question'); });
test('Provision file: an amount with no Storage Code stops the tool', () => expectError({ hkRows: HK_ROWS().concat([{ code: null, name: 'x', manpower: 100 }]) }, /Manpower amount but no Storage Code/));
test('Provision file: blank Manpower counts as 0 and is listed', async () => {
  const rows = HK_ROWS(); rows[3] = { ...rows[3], manpower: undefined };
  const { R } = await run({ hkRows: rows });
  return near(R.stmt.find((x) => x.code === 9004).raj, 0) && R.points.some((p) => /"Manpower" is blank/.test(p.point)) ? true : 'not listed';
});
test('Opening: statement for the wrong month stops the tool', () => expectError({ openTitle: 'Housekeeping Exp – Provision for the month of Jul-26' }, /titled .*Jul-26.*Aug-26/));
test('Opening: two "YTD FY27" columns → the user picks one', async () => {
  const hdr = OPEN_HDR.concat(['YTD FY27']); const rows = openRows(CODES).map((r) => r.concat([1]));
  const r = await run({ openHdr: hdr, openRows: rows });
  if (!r.needChoice || r.needChoice.kind !== 'openingYtdCol') return 'no question';
  const r2 = await run({ openHdr: hdr, openRows: rows }, { openingYtdCol: r.needChoice.options[1].col });
  return near(r2.R.T.proCur, CODES.length) ? true : 'choice not used: ' + r2.R.T.proCur;
});
test('April: stops, the SOP has no year-end roll-over rule', () => expectError({ period: 'Apr-27', from: '2027-04-01', to: '2027-04-30', sheet: 'HK Apr-27', openTitle: 'Housekeeping Exp – Provision for the month of Mar-27', lines: [L(9001, 10, 'for the month of April 2027')] }, /April/));
test('Report dates: a To Date that is not the month end stops the tool', () => expectError({ to: '2026-09-29' }, /To Date/));

(async () => {
  let bad = 0;
  for (const t of tests) {
    let r; try { r = await t.fn(); } catch (e) { r = 'ERROR ' + (e.list ? e.list.join(' | ') : e.stack); }
    if (r !== true) bad++;
    console.log(`${r === true ? 'ok  ' : 'FAIL'} ${t.name}${r === true ? '' : ' — ' + r}`);
  }
  const { buffer } = await run(); fs.writeFileSync(path.join(OUT, 'hk_expenses_sample.xlsx'), buffer);
  console.log(`\n${tests.length - bad} passed, ${bad} failed`);
  process.exit(bad ? 1 : 0);
})();
