// The core's year tagger (src/core/tagger.js) must behave exactly like Security's.
// 1. Runs Security's own tagger test file (test/security/tagger_test.js) with the core tagger in place of Security's.
// 2. Compares both taggers on every Line Description in the sample extract(s) in test-data/, if present.
'use strict';
const fs = require('fs'); const path = require('path'); const vm = require('vm');
const SEC = require('../../src/security/engine.js');
const core = require('../../src/core/tagger.js');
const common = require('../../src/core/common.js');

const ROOT = path.join(__dirname, '..', '..');
let bad = 0;

// 1. Security's 56 cases, core tagger
const src = fs.readFileSync(path.join(ROOT, 'test/security/tagger_test.js'), 'utf8');
const shim = { ...SEC, proposeFY: core.proposeFY, fyOf: common.fyOf };
const out = [];
const ctx = { require: (m) => (m.endsWith('src/security/engine.js') ? shim : require(m)), console: { log: (...a) => out.push(a.join(' ')) }, process: { exit: (c) => { if (c) bad++; } } };
vm.runInNewContext(src, ctx, { filename: 'tagger_test.js (core tagger)' });
console.log('Security tagger cases with the core tagger:', out[out.length - 1]);
out.filter((l) => l.startsWith('FAIL')).forEach((l) => console.log('  ' + l));

// 2. Real narrations: both taggers must give identical answers
(async () => {
  const files = fs.existsSync(path.join(ROOT, 'test-data')) ? fs.readdirSync(path.join(ROOT, 'test-data')).filter((f) => /^extract_.*\.xlsx$/.test(f)) : [];
  for (const f of files) {
    const wbs = await SEC.load({ extract: fs.readFileSync(path.join(ROOT, 'test-data', f)), opening: fs.readFileSync(path.join(ROOT, 'test-data', f)), rajeev: fs.readFileSync(path.join(ROOT, 'test-data', f)) }).catch(() => null);
    if (!wbs) continue;
    let ex; try { ex = SEC._internals.readExtract(wbs.extract); } catch (e) { console.log(`${f}: not a Security-account extract, skipped`); continue; }
    let n = 0, diff = 0;
    for (const r of ex.rows) {
      n++;
      const a = SEC.proposeFY(r.desc, ex.period), b = core.proposeFY(r.desc, ex.period);
      if (JSON.stringify(a) !== JSON.stringify(b)) { diff++; bad++; console.log('  DIFFERENT', f, 'row', r.erpRow, JSON.stringify(a), JSON.stringify(b)); }
    }
    console.log(`${f}: ${n} narrations, ${diff} different`);
  }
  console.log(bad ? 'TAGGER PARITY FAILED' : 'TAGGER PARITY OK');
  process.exit(bad ? 1 : 0);
})();
