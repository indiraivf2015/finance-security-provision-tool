// The core's year tagger (src/core/tagger.js) is Security's tagger plus the Housekeeping changes of Oct 2026
// (see the header of src/core/tagger.js). This test proves the core differs from Security only where intended:
//   - every narration Security tags itself (sure) gets exactly the same answer from the core;
//   - a narration Security leaves to the default may be tagged by the core, and only in the cases listed in
//     INTENDED below (Security's own test cases) or reported (real narrations).
// 1. Security's own tagger cases (test/security/tagger_test.js).
// 2. Every Line Description in the sample extract(s) in test-data/, if present.
'use strict';
const fs = require('fs'); const path = require('path'); const vm = require('vm');
const SEC = require('../../src/security/engine.js');
const core = require('../../src/core/tagger.js');

const ROOT = path.join(__dirname, '..', '..');
let bad = 0;
const same = (a, b) => a.fy === b.fy && a.sure === b.sure && a.basis === b.basis;

// Security cases the core now tags by the year written in the narration (same year as Security's suggestion)
const INTENDED = new Map([
  ['Being GRN Park form House keeping invoice no SA07DEL0121/2627 dt 07/07/2026', 27],
  ['BEING GRN PARKED FOR THE HOUSEKEEPING SERVICE FOR DEC AT GURUGRAM SEC-83 CENTRE INV NO-BNHR26GN000206 DT-23/02/2026', 26],
  ['Invoice dated May 15 2026', 27],
  ['Being grn created for security guard expenses for the month of may invoice ss/26-27/84 dt 15.06.2026', 27],
  ['Invoice date Sep 25, 2026', 27],
  ['Security for the month of March 2026, invoice dated 05-Apr-2026', 26],
]);

// 1. Security's cases: Security's own checks run unchanged; every narration is also given to the core
const rec = [];
const shim = { ...SEC, proposeFY: (d, p) => { rec.push({ d, p }); return SEC.proposeFY(d, p); } };
const out = [];
const ctx = { require: (m) => (m.endsWith('src/security/engine.js') ? shim : require(m)), console: { log: (...a) => out.push(a.join(' ')) }, process: { exit: (c) => { if (c) bad++; } } };
vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'test/security/tagger_test.js'), 'utf8'), ctx, { filename: 'tagger_test.js' });
let caseDiff = 0;
for (const { d, p } of rec) {
  const a = SEC.proposeFY(d, p), b = core.proposeFY(d, p);
  if (same(a, b)) continue;
  caseDiff++;
  const ok = !a.sure && INTENDED.has(d) && b.sure && b.fy === INTENDED.get(d);
  if (!ok) { bad++; console.log('  FAIL', JSON.stringify(d), 'Security', JSON.stringify(a), 'core', JSON.stringify(b)); }
}
const unused = [...INTENDED.keys()].filter((d) => !rec.some((r) => r.d === d));
if (unused.length) { bad++; console.log('  FAIL: INTENDED lists narrations that are not Security cases:', unused.join(' | ')); }
console.log(`Security tagger cases: ${rec.length}; core differs on ${caseDiff} (intended: ${INTENDED.size})`);

// 2. Real narrations: lines Security tags itself must be identical; other differences are listed
(async () => {
  const files = fs.existsSync(path.join(ROOT, 'test-data')) ? fs.readdirSync(path.join(ROOT, 'test-data')).filter((f) => /^extract_.*\.xlsx$/.test(f)) : [];
  for (const f of files) {
    const wbs = await SEC.load({ extract: fs.readFileSync(path.join(ROOT, 'test-data', f)), opening: fs.readFileSync(path.join(ROOT, 'test-data', f)), rajeev: fs.readFileSync(path.join(ROOT, 'test-data', f)) }).catch(() => null);
    if (!wbs) continue;
    let ex; try { ex = SEC._internals.readExtract(wbs.extract); } catch (e) { console.log(`${f}: not a Security-account extract, skipped`); continue; }
    let n = 0, diff = 0, yearDiff = 0;
    for (const r of ex.rows) {
      n++;
      const a = SEC.proposeFY(r.desc, ex.period), b = core.proposeFY(r.desc, ex.period);
      if (same(a, b)) continue;
      diff++;
      if (a.sure) { bad++; console.log('  FAIL (Security tags this line itself)', f, 'row', r.erpRow, JSON.stringify(a), JSON.stringify(b)); continue; }
      const used = (x) => (x.sure ? x.fy : ex.period.cur);
      if (used(a) !== used(b)) yearDiff++;
      console.log(`  core tags by written year: ${f} row ${r.erpRow} FY ${b.fy}${used(a) !== used(b) ? ` (Security's default FY ${used(a)})` : ''}`);
    }
    console.log(`${f}: ${n} narrations, ${diff} tagged by the core only, ${yearDiff} with a different year`);
  }
  console.log(bad ? 'TAGGER PARITY FAILED' : 'TAGGER PARITY OK');
  process.exit(bad ? 1 : 0);
})();
