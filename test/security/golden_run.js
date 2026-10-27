// Produces the Security output exactly as the page does (no year chosen by hand: the finance rule decides),
// for the golden-output check. Usage: node test/security/golden_run.js <engine.js> <extract> <opening> <rajeev> <out.xlsx> [optionsJSON]
'use strict';
const fs = require('fs'); const path = require('path');
(async () => {
  const [engine, ex, op, rj, out, optJson] = process.argv.slice(2);
  const E = require(path.resolve(engine));
  const wbs = await E.load({ extract: fs.readFileSync(ex), opening: fs.readFileSync(op), rajeev: fs.readFileSync(rj) });
  const st = E.analyse(wbs, optJson ? JSON.parse(optJson) : {});
  if (st.needChoice) { console.error('NEEDS CHOICE', JSON.stringify(st.needChoice)); process.exit(3); }
  const { buffer, result } = await E.build(st, new Map(), { now: new Date(0), nowText: 'golden', names: { extract: path.basename(ex), opening: path.basename(op), rajeev: path.basename(rj) } });
  fs.mkdirSync(path.dirname(out), { recursive: true }); fs.writeFileSync(out, Buffer.from(buffer));
  console.log('STATUS', result.status, '| YTD', result.T.ytd.toFixed(2));
})().catch((e) => { console.error(e.list ? e.list.join('\n') : e.stack); process.exit(2); });
