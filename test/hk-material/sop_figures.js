// Housekeeping Material against the figures printed in its SOP, using Rajeev's real Sep-26 file.
// The SOP figures are finance data, so they live in test-data/hk-expected.json (git-ignored); skipped if absent.
// The SOP shows whole rupees, so figures are compared as shown (exact values rounded to the rupee).
//  1. Step 11: the Sep-26 Materials total, as printed in the SOP
//  2. Part E worked example (three centres, one with a negative Prov CF): Aug-Provision, Booking, Prov CF,
//     Sept-Provision, YTD. Run the SOP way (minus sign carried), which is what the SOP's example shows.
// NOT covered until finance sends the Sep-26 51611 extract and Aug-26 Material statement: the statement totals.
// Usage: node test/hk-material/sop_figures.js
'use strict';
const fs = require('fs'); const path = require('path');
const E = require('../../src/provisions/hk-material').withNegativeCfRule(null);
const F = require('../fixtures');

const TD = (f) => path.join(__dirname, '..', '..', 'test-data', f);
if (!fs.existsSync(TD('hk-expected.json'))) { console.log('SKIPPED: test-data/hk-expected.json not present'); process.exit(0); }
const X = JSON.parse(fs.readFileSync(TD('hk-expected.json'), 'utf8')).hkMaterial;
if (!fs.existsSync(TD(X.provisionFile))) { console.log('SKIPPED: test-data/' + X.provisionFile + ' not present'); process.exit(0); }

const rupee = (x) => Math.round(Number(x.toPrecision(15)));
let bad = 0;
const ok = (cond, text, detail) => { if (!cond) bad++; console.log(`${cond ? 'ok  ' : 'FAIL'} ${text}${cond ? '' : ' — ' + detail}`); };

(async () => {
  const ex = X.workedExample;
  const files = {
    extract: await F.extract({ period: 'Sep-26', from: '2026-09-01', to: '2026-09-30', glFrom: '51611', glTo: '51611',
      lines: ex.map((w) => ({ code: w.code, name: w.name, account: 51611, bal: w.booking, desc: 'Housekeeping material for the month of August 2026' })) }),
    opening: await F.statement({ title: 'Material Provision for the month of Aug-26', headers: ['Cost Center', 'Cost Center Name', 'Jul-Provision', 'Aug-Booking', 'Prov CF', 'Aug-Provision', 'YTD'],
      rows: ex.map((w) => [w.code, w.name, 0, 0, 0, 0, w.open]) }),
    rajeev: fs.readFileSync(TD(X.provisionFile)),
  };
  const st = E.analyse(await E.load(files), {});
  if (st.needChoice) throw new Error('asked ' + JSON.stringify(st.needChoice));
  const { result: R } = await E.build(st, new Map(), { now: new Date(0) });

  console.log(`Provision file sheet "${st.rajeev.sheet}", Materials column ${st.rajeev.amtColLetter} ("${st.rajeev.amtHeading}"), ${st.rajeev.lines.length} rows`);
  ok(rupee(R.T.prov) === X.materialsTotal, `Sept-Provision column total = ${X.materialsTotal} (SOP), = Rajeev's Materials total`, `${rupee(R.T.prov)} (exact ${R.T.prov})`);
  for (const w of ex) {
    const s = R.stmt.find((x) => x.code === w.code);
    const got = { open: s.open, booking: s.booking, provCf: s.cf, provision: s.prov, ytd: s.ytd };
    const diff = Object.keys(got).filter((k) => rupee(got[k]) !== w[k]);
    ok(!diff.length, `Worked example ${w.code} ${w.name}: Aug-Provision ${w.open}, Booking ${w.booking}, Prov CF ${w.provCf}, Sept-Provision ${w.provision}, YTD ${w.ytd}`, diff.map((k) => `${k} ${got[k]} ≠ ${w[k]}`).join(', '));
  }
  console.log('\nNot yet verified (needs the Sep-26 51611 ERP extract and Aug-26 Material statement): SOP statement totals ' + JSON.stringify(X.statementTotals_needsRealExtract));
  console.log(`${bad ? 'FAILED' : 'SOP FIGURES OK'}`);
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error(e.list ? e.list.join('\n') : e.stack); process.exit(1); });
