// Usage: node test/security/run.js <extract.xlsx> <opening.xlsx> <rajeev.xlsx> <out.xlsx> [optionsJSON]
// Accepts every tool suggestion for lines needing confirmation (simulates "Accept all suggestions").
const fs = require('fs');
const SecProv = require('../../src/security/engine.js');
(async () => {
  const [ex, op, rj, out, optJson, humanFy] = process.argv.slice(2); // humanFy: the year a reviewer chose for lines with no suggestion
  const opts = optJson ? JSON.parse(optJson) : {};
  const wbs = await SecProv.load({ extract: fs.readFileSync(ex), opening: fs.readFileSync(op), rajeev: fs.readFileSync(rj) });
  let st;
  try { st = SecProv.analyse(wbs, opts); } catch (e) { console.log('INPUT ERROR:\n' + (e.list ? e.list.join('\n') : e.stack)); process.exit(2); }
  if (st.needChoice) { console.log('NEEDS CHOICE', JSON.stringify(st.needChoice)); process.exit(3); }
  console.log('period', st.period, 'kept', st.kept.length, 'removed', st.removed.length);
  const need = st.kept.filter((r) => r.needsConfirm);
  console.log('needs confirmation:', need.length);
  for (const r of need) console.log('  row', r.erpRow, r.seg1, r.bal, '→', r.suggestFy, '|', r.basis, '|', r.desc.replace(/https?:\/\/\S+/g, '<url>').slice(0, 90));
  const choice = new Map(need.map((r) => [r.id, r.suggestFy !== null ? r.suggestFy : (humanFy ? parseInt(humanFy, 10) : null)]).filter(([, v]) => v !== null));
  if (need.some((r) => r.suggestFy === null)) { console.log('Some lines have no suggestion'); }
  const { buffer, result } = await SecProv.build(st, choice, { nowText: 'test', names: { extract: ex, opening: op, rajeev: rj } });
  fs.mkdirSync(require('path').dirname(out) || '.', { recursive: true });
  fs.writeFileSync(out, Buffer.from(buffer));
  console.log('STATUS', result.status);
  for (const c of result.checks) console.log(' ', c.status.padEnd(6), c.id, c.text.slice(0, 90), '| exp', c.expected, 'act', c.actual);
  const T = result.T; for (const k of Object.keys(T)) console.log('  total', k, T[k].toFixed(2));
  console.log('exceptions', result.exceptions.length, [...new Set(result.exceptions.map((e) => e.type))].join(' / '));
})();
