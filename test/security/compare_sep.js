// Compares the tool's Sep-26 Summary with finance's own Sep-26 statement, centre by centre.
// Usage: node test/security/compare_sep.js <tool_output.xlsx> <finance_statement.xlsx>
// Port of the former compare_sep.py: reads saved (cached) values; reports any difference >= 1 paisa.
'use strict';
const ExcelJS = require('exceljs');
const { rowValues, isNum, pyRound2, pyFloat, pyRepr, pySet } = require('../xlsx_values');

const num0 = (x) => { if (x === null || x === 0 || x === '' || x === false) return 0; const n = Number(x); if (Number.isNaN(n)) throw new Error('not a number: ' + pyRepr(x)); return n; };

(async () => {
  const [out, fin] = process.argv.slice(2);
  const tw = new ExcelJS.Workbook(); await tw.xlsx.readFile(out);
  const ws = tw.getWorksheet('Summary');
  const hdr = rowValues(ws, 4, ws.columnCount);
  const title = rowValues(ws, 2, 1)[0];
  console.log('title:', title === null ? 'None' : String(title)); console.log('header:', pyRepr(hdr));
  const tool = new Map();
  for (let r = 5; r <= ws.rowCount; r++) {
    const v = rowValues(ws, r, ws.columnCount);
    if (v[1] === null) continue;
    tool.set(v[1], v.slice(3, 11).map(num0));
  }
  const fw = new ExcelJS.Workbook(); await fw.xlsx.readFile(fin);
  const fs = fw.getWorksheet('Summary');
  const raj = new Map();
  const wr = fw.getWorksheet('Provision_Rajeev');
  for (let r = 4; r <= wr.rowCount; r++) {
    const v = rowValues(wr, r, Math.max(wr.columnCount, 9));
    if (isNum(v[2]) && isNum(v[8])) { const k = Math.trunc(v[2]); raj.set(k, (raj.get(k) || 0) + v[8]); }
  }
  const finance = new Map();
  for (let r = 5; r <= fs.rowCount; r++) {
    const row = rowValues(fs, r, fs.columnCount);
    if (!isNum(row[1])) continue;
    const c = Math.trunc(row[1]); const v = row.slice(3, 9).map((x) => (isNum(x) ? x : 0));
    v.push(raj.get(c) || 0, v[4] + (raj.get(c) || 0));
    finance.set(c, v);
  }
  const cols = ['FY27 Pro', 'FY26 Pro', 'FY27 Bk', 'FY26 Bk', 'FY27 CF', 'FY26 CF', 'Rajeev', 'YTD'];
  const onlyTool = [...tool.keys()].filter((k) => !finance.has(k)); const onlyFin = [...finance.keys()].filter((k) => !tool.has(k));
  console.log('centres tool', tool.size, 'finance', finance.size, 'only tool', pySet(onlyTool), 'only finance', pySet(onlyFin));
  let diffs = 0;
  for (const [c, fv] of finance) {
    if (!tool.has(c)) throw new Error('KeyError: ' + c);
    const tv = tool.get(c); const n = Math.min(tv.length, fv.length);
    const d = []; for (let i = 0; i < n; i++) d.push(pyRound2(tv[i] - fv[i]));
    if (d.some((x) => Math.abs(x) >= 0.01)) {
      diffs++;
      console.log('DIFF', c, '{' + d.map((x, i) => [cols[i], x]).filter(([, x]) => Math.abs(x) >= 0.01).map(([k, x]) => `'${k}': ${pyFloat(x)}`).join(', ') + '}');
    }
  }
  console.log('centres with any difference >= 1 paisa:', diffs);
})().catch((e) => { console.error(e.stack || e); process.exit(1); });
