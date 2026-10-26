// Scenario tests: bad inputs must stop with a clear message; good inputs (incl. the tool's own output as next month's opening) must work.
const fs = require('fs'); const ExcelJS = require('exceljs'); const E = require('../src/engine.js');
const S = process.argv[2]; // folder with real files
fs.mkdirSync('out', { recursive: true });
const F = { ex: S + '/extract_Sep-26.xlsx', op: S + '/statement_Aug-26.xlsx', rj: S + '/provision_Sep-26.xlsx' }; // sample files in test-data/ (not in git)
const OWN_SEP = process.argv[3]; // tool's own Sep-26 output
const read = async (p) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(fs.readFileSync(p)); return wb; };
const buf = async (wb) => Buffer.from(await wb.xlsx.writeBuffer());
let pass = 0, fail = 0;
async function expectError(name, files, opts, re) {
  try {
    const wbs = await E.load(files); const st = E.analyse(wbs, opts || {});
    if (st.needChoice) { if (re.test('NEEDCHOICE:' + st.needChoice.kind)) { pass++; console.log('ok  ', name, '→ asks', st.needChoice.kind); return; } }
    fail++; console.log('FAIL', name, '— expected an error, got', st.needChoice ? 'choice ' + st.needChoice.kind : 'success');
  } catch (e) {
    const msg = e.list ? e.list.join(' | ') : e.message;
    if (re.test(msg)) { pass++; console.log('ok  ', name, '→', msg.slice(0, 150)); } else { fail++; console.log('FAIL', name, '— wrong error:', msg.slice(0, 300)); }
  }
}
const base = async () => ({ extract: fs.readFileSync(F.ex), opening: fs.readFileSync(F.op), rajeev: fs.readFileSync(F.rj) });
(async () => {
  const O = { openingYtdCol: 'M' };
  // 1. Wrong month statement (use Sep extract with a statement titled Jul-26)
  { const wb = await read(F.op); wb.getWorksheet('Summary').getCell('A2').value = 'Security Provision for the month of Jul-26'; const f = await base(); f.opening = await buf(wb); await expectError('opening statement for wrong month', f, O, /titled .*Jul-26.*Sep-26/); }
  // 2. Two periods in extract
  { const wb = await read(F.ex); const ws = wb.worksheets[0]; ws.getCell('A12').value = 'Aug-26'; const f = await base(); f.extract = await buf(wb); await expectError('extract with two periods', f, O, /more than one period/); }
  // 3. Account outside range
  { const wb = await read(F.ex); const ws = wb.worksheets[0]; ws.getCell('E12').value = '51801'; const f = await base(); f.extract = await buf(wb); await expectError('account outside 51701–51704', f, O, /51801.*outside/); }
  // 4. Missing column
  { const wb = await read(F.ex); const ws = wb.worksheets[0]; ws.getCell('Q9').value = 'BAL_X'; const f = await base(); f.extract = await buf(wb); await expectError('BALANCE column missing', f, O, /no sheet has the Account Analysis|BALANCE is missing/); }
  // 5. Non-numeric amount
  { const wb = await read(F.ex); const ws = wb.worksheets[0]; ws.getCell('Q15').value = '12,34x'; const f = await base(); f.extract = await buf(wb); await expectError('text in BALANCE', f, O, /row 15: BALANCE text/); }
  // 6. Rajeev sheet for another month
  { const wb = await read(F.rj); wb.getWorksheet('Security Sep-26').name = 'Security Aug-26'; const f = await base(); f.rajeev = await buf(wb); await expectError('provision file for wrong month', f, O, /Security Aug-26.*Sep-26/); }
  // 7. Rajeev amount without code
  { const wb = await read(F.rj); const ws = wb.getWorksheet('Security Sep-26'); ws.getCell('D10').value = null; const f = await base(); f.rajeev = await buf(wb); await expectError('provision amount without storage code', f, O, /row 10: there is a provision amount but no Storage Code/); }
  // 8. Formula with no saved value in opening statement
  { const wb = await read(F.op); const ws = wb.getWorksheet('Summary'); ws.getCell('M5').value = { formula: 'K5+L5' }; const f = await base(); f.opening = await buf(wb); await expectError('opening formula with no saved value', f, O, /row 5.*no saved value/); }
  // 9. Excel error in opening
  { const wb = await read(F.op); const ws = wb.getWorksheet('Summary'); ws.getCell('M6').value = { formula: 'K6+L6', result: { error: '#REF!' } }; const f = await base(); f.opening = await buf(wb); await expectError('opening cell with #REF!', f, O, /row 6.*#REF!/); }
  // 10. Duplicate cost centre in opening
  { const wb = await read(F.op); const ws = wb.getWorksheet('Summary'); ws.getCell('B6').value = 1001; const f = await base(); f.opening = await buf(wb); await expectError('duplicate cost centre in opening', f, O, /1001 appears more than once/); }
  // 11. Two YTD columns → must ask
  { await expectError('two YTD columns → user must choose', await base(), {}, /NEEDCHOICE:openingYtdCol/); }
  // 12. Not an xlsx
  { const f = await base(); f.rajeev = Buffer.from('a,b,c\n1,2,3'); await expectError('csv given instead of xlsx', f, O, /not an \.xlsx file/); }
  // 13. Year-end: April extract with a March statement (no YTD FY28)
  { const wb = await read(F.ex); const ws = wb.worksheets[0]; ws.eachRow((row, r) => { if (r > 9 && row.getCell(1).value) row.getCell(1).value = 'Apr-27'; }); ws.getCell('A5').value = '2027-04-01T00:00:00.000+00:00'; ws.getCell('B5').value = '2027-04-30T00:00:00.000+00:00';
    const wo = await read(F.op); wo.getWorksheet('Summary').getCell('A2').value = 'Security Provision for the month of Mar-27';
    const wr = await read(F.rj); wr.getWorksheet('Security Sep-26').name = 'Security Apr-27';
    await expectError('April (new financial year) — not covered by SOP', { extract: await buf(wb), opening: await buf(wo), rajeev: await buf(wr) }, O, /new financial year/); }
  // 14. Partial-month extract (parameter To Date mid-month)
  { const wb = await read(F.ex); wb.worksheets[0].getCell('B5').value = '2026-09-15T00:00:00.000+00:00'; const f = await base(); f.extract = await buf(wb); await expectError('extract for part of the month', f, O, /To Date.*not the last day/); }
  // 15. Unconfirmed lines block the build
  { const wbs = await E.load(await base()); const st = E.analyse(wbs, O); const R = E.compute(st, new Map());
    const und = st.kept.filter((r) => r.needsConfirm);
    if (und.length && und.every((r) => r.fy === st.period.cur && r.tagBy === 'Default') && st.kept.filter((r) => !r.needsConfirm).every((r) => r.fy === r.autoFy) && R.points.some((p) => /no year in the description/.test(p.point))) { pass++; console.log(`ok   ${und.length} lines with no year in the narration → current FY by default, listed in Points to check`); }
    else { fail++; console.log('FAIL default FY for undated lines'); } }

  // 16. CHAIN: the tool's own Sep-26 output as the opening for a synthetic Oct-26 run
  {
    const wb = await read(F.ex); const ws = wb.worksheets[0];
    ws.eachRow((row, r) => { if (r > 9 && row.getCell(1).value === 'Sep-26') row.getCell(1).value = 'Oct-26'; });
    ws.getCell('A5').value = '2026-10-01T00:00:00.000+00:00'; ws.getCell('B5').value = '2026-10-31T00:00:00.000+00:00';
    const wr = await read(F.rj); wr.getWorksheet('Security Sep-26').name = 'Security Oct-26'; wr.getWorksheet('HK Sep-26').name = 'HK Oct-26';
    const files = { extract: await buf(wb), opening: fs.readFileSync(OWN_SEP), rajeev: await buf(wr) };
    const wbs = await E.load(files); const st = E.analyse(wbs, {});
    if (st.needChoice) { fail++; console.log('FAIL chain asks', st.needChoice); }
    else {
      const choice = new Map(st.kept.filter((r) => r.needsConfirm).map((r) => [r.id, r.suggestFy !== null ? r.suggestFy : st.period.cur])); // reviewer's choice for undated lines
      const R = E.compute(st, choice);
      // Opening must equal our own Sep-26 closing per centre
      const own = await read(OWN_SEP); const os = own.getWorksheet('Summary'); const close = new Map();
      for (let r = 5; r <= os.rowCount; r++) { const code = os.getCell(r, 2).value; if (code == null) continue; close.set(code, { ytd: os.getCell(r, 11).model.result, cf: os.getCell(r, 9).model.result }); }
      let bad = 0; for (const s of R.stmt) { const c = close.get(s.code); if (!c || Math.abs(c.ytd - s.proCur) > 1e-6 || Math.abs(c.cf - s.proPrev) > 1e-6) bad++; }
      const allPass = R.checks.every((c) => c.status !== 'FAIL');
      if (!bad && allPass && R.stmt.length === close.size) { pass++; console.log(`ok   chain: tool's own Sep-26 output read as Oct-26 opening — ${R.stmt.length} centres carried exactly, checks pass, header "${st.period.label}"`); }
      else { fail++; console.log('FAIL chain', { bad, allPass, n: R.stmt.length, m: close.size }); }
    }
  }

  // 17. A month with no bookings at all (every line is a non-reclass spreadsheet journal)
  {
    const wb = await read(F.ex); const ws = wb.worksheets[0];
    ws.eachRow((row, r) => { if (r > 9 && row.getCell(1).value) { row.getCell(18).value = 'Spreadsheet'; row.getCell(37).value = 'SOMEONE'; } });
    const wbs = await E.load({ extract: await buf(wb), opening: fs.readFileSync(F.op), rajeev: fs.readFileSync(F.rj) });
    const st = E.analyse(wbs, O);
    const { buffer, result } = await E.build(st, new Map(), { nowText: 'test' });
    fs.writeFileSync('out/zero_kept.xlsx', Buffer.from(buffer));
    const ok = st.kept.length === 0 && result.T.bkCur === 0 && result.T.bkPrev === 0 && result.checks.every((c) => c.status !== 'FAIL') && result.exceptions.some((e) => e.type === 'No bookings this month');
    if (ok) { pass++; console.log('ok   month with no bookings → statement built, bookings 0, flagged in Exceptions (out/zero_kept.xlsx)'); } else { fail++; console.log('FAIL zero-bookings month', st.kept.length, result.status); }
  }
  // 18. Centre filter in report parameters → stop
  { const wb = await read(F.ex); wb.worksheets[0].getCell('E5').value = '1001'; const f = await base(); f.extract = await buf(wb); await expectError('extract filtered to some centres', f, O, /Center From.*all centres/); }
  // 19. Negative CF rule cannot be applied (negatives larger than the top 10 centres) -> must not be READY
  { const wb = await read(F.op); const ws = wb.getWorksheet('Summary');
    ws.eachRow((row, r) => { if (r > 4 && typeof row.getCell(2).value === 'number') row.getCell(13).value = 0; });
    const f = await base(); f.opening = await buf(wb);
    const wbs = await E.load(f); const st = E.analyse(wbs, O);
    const R = E.compute(st, new Map(st.kept.filter((r) => r.needsConfirm).map((r) => [r.id, st.period.cur])));
    const ok = R.status === 'NOT READY' && R.checks.find((c) => c.id === 16).status === 'FAIL' && R.points.some((p) => /could NOT be absorbed/.test(p.point));
    if (ok) { pass++; console.log('ok   negative CF larger than the top 10 can absorb → NOT READY, explained in Points to check'); } else { fail++; console.log('FAIL negative-CF impossible case', R.status); } }
  // 20. Negative CF rule: column totals unchanged, no negatives left, recipients never pushed below 0 (Sept data)
  { const wbs = await E.load(await base()); const st = E.analyse(wbs, O);
    const R = E.compute(st, new Map(st.kept.filter((r) => r.needsConfirm).map((r) => [r.id, r.suggestFy !== null ? r.suggestFy : st.period.cur])));
    const sum = (f) => R.stmt.reduce((a, s) => a + s[f], 0);
    const ok = Math.abs(sum('cfCur') - sum('cfCurSop')) < 0.005 && Math.abs(sum('cfPrev') - sum('cfPrevSop')) < 0.005 && R.stmt.every((s) => s.cfCur > -0.005 && s.cfPrev > -0.005) && R.stmt.filter((s) => s.adjCur < 0).length === 10;
    if (ok) { pass++; console.log('ok   negative CF rule: totals unchanged, none negative, 10 recipients'); } else { fail++; console.log('FAIL negative-CF rule totals'); } }
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
