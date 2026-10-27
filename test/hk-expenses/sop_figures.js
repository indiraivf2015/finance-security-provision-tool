// Housekeeping Expenses against the figures printed in its SOP (v1.1), using Rajeev's real Sep-26 file.
// The SOP figures are finance data, so they live in test-data/hk-expected.json (git-ignored); skipped if absent.
//  1. Part D: Manpower rows converted (÷118×100×1.135), a Main + OPD centre, and the Sep-26 total, as printed in the SOP
//  2. Part E worked example (two centres): opening, booking, CF, Rajeev and YTD exactly as the SOP
// NOT covered until finance sends the Sep-26 HK ERP extract and Aug-26 statement: the SOP's statement totals.
// Usage: node test/hk-expenses/sop_figures.js
'use strict';
const fs = require('fs'); const path = require('path');
const E = require('../../src/provisions/hk-expenses');
const F = require('../fixtures');

const TD = (f) => path.join(__dirname, '..', '..', 'test-data', f);
if (!fs.existsSync(TD('hk-expected.json'))) { console.log('SKIPPED: test-data/hk-expected.json not present'); process.exit(0); }
const X = JSON.parse(fs.readFileSync(TD('hk-expected.json'), 'utf8')).hkExpenses;
if (!fs.existsSync(TD(X.provisionFile))) { console.log('SKIPPED: test-data/' + X.provisionFile + ' not present'); process.exit(0); }

const r2 = (x) => Math.round(Number(x.toPrecision(15)) * 100) / 100; // as Excel shows it: 2 decimals
let bad = 0;
const ok = (cond, text, detail) => { if (!cond) bad++; console.log(`${cond ? 'ok  ' : 'FAIL'} ${text}${cond ? '' : ' — ' + detail}`); };

(async () => {
  const ex = X.workedExample;
  const lines = ex.map((w) => ({ code: w.code, name: w.name, account: 51610, bal: w.bkCur, desc: 'Housekeeping services for the month of August 2026', attr3: w.vendor }));
  const files = {
    extract: await F.extract({ period: X.period, from: '2026-09-01', to: '2026-09-30', lines }),
    opening: await F.statement({ title: 'Housekeeping Exp – Provision for the month of Aug-26',
      headers: ['Cost Center', 'Cost Center Name', 'FY27 Pro', 'FY26 Pro', 'FY27 Booking', 'FY26 Booking', 'FY27 Pro CF', 'FY26 Pro CF', 'Aug-26 Pro as per Rajeev', 'YTD FY27'],
      rows: ex.map((w) => [w.code, w.name, 0, 0, 0, 0, 0, w.openCf, 0, w.openYtd]) }),
    rajeev: fs.readFileSync(TD(X.provisionFile)),
  };
  const st = E.analyse(await E.load(files), {});
  if (st.needChoice) throw new Error('asked ' + JSON.stringify(st.needChoice));
  const { result: R } = await E.build(st, new Map(), { now: new Date(0) });

  // 1. Part D
  console.log(`Provision file sheet "${st.rajeev.sheet}", Manpower column ${st.rajeev.amtColLetter}, ${st.rajeev.lines.length} rows`);
  for (const w of X.rajeevRows) {
    const l = st.rajeev.lines.find((x) => x.code === w.code && Math.abs(x.amount - w.manpower) < 0.005);
    ok(l && r2(l.basic) === w.basic && r2(l.final) === w.final, `Part D row ${w.code} ${w.manpower}: Basic ${w.basic}, Final ${w.final}`, l ? `${r2(l.basic)} / ${r2(l.final)}` : 'row not found');
  }
  for (const [code, v] of Object.entries(X.rajeevByCode)) { const s = R.stmt.find((x) => x.code === +code); ok(s && r2(s.raj) === v, `Pro as per Rajeev, centre ${code} = ${v}`, s ? r2(s.raj) : 'no row'); }
  ok(r2(R.totals.rajFinal) === X.rajeevFinalTotal, `Pro as per Rajeev total = ${X.rajeevFinalTotal} (SOP Sep-26)`, r2(R.totals.rajFinal));
  ok(r2(R.T.raj) === X.rajeevFinalTotal, 'statement column total = the Part D total', r2(R.T.raj));

  // 2. Part E worked example
  for (const w of ex) {
    const s = R.stmt.find((x) => x.code === w.code);
    const got = { proCur: s.proCur, proPrev: s.proPrev, bkCur: s.bkCur, bkPrev: s.bkPrev, cfCur: s.cfCur, cfPrev: s.cfPrev, raj: s.raj, ytd: s.ytd };
    const want = { proCur: w.openYtd, proPrev: w.openCf, bkCur: w.bkCur, bkPrev: w.bkPrev, cfCur: w.cfCur, cfPrev: w.cfPrev, raj: w.raj, ytd: w.ytd };
    const diff = Object.keys(want).filter((k) => r2(got[k]) !== want[k]);
    ok(!diff.length, `Worked example ${w.code} ${w.name}: FY27 Pro ${w.openYtd}, Booking ${w.bkCur}, FY27 Pro CF ${w.cfCur}, Rajeev ${w.raj}, YTD ${w.ytd}`, diff.map((k) => `${k} ${r2(got[k])} ≠ ${want[k]}`).join(', '));
  }
  console.log('\nNot yet verified (needs the Sep-26 HK ERP extract and Aug-26 HK statement): SOP statement totals ' + JSON.stringify(X.statementTotals_needsRealExtract));
  console.log(`${bad ? 'FAILED' : 'SOP FIGURES OK'}`);
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error(e.list ? e.list.join('\n') : e.stack); process.exit(1); });
