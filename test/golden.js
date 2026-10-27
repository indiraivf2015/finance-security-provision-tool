// Golden-output check: a provision's output must be identical to the approved output.
// Usage: node test/golden.js <approved.xlsx> <new.xlsx>
// Strict: every XML part inside the .xlsx (cell values, formulas, saved results, styles, widths, merges,
// sheet order …) must be identical. Only the workbook's created/modified timestamps are ignored.
// When something differs, a cell-by-cell report says what changed. Exit code 0 = identical, 1 = different.
'use strict';
const fs = require('fs');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');
const { cached, isFormula } = require('./xlsx_values');

const IGNORED = /<dcterms:(created|modified)[^>]*>[^<]*<\/dcterms:\1>/g; // timestamps in docProps/core.xml

async function parts(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = new Map();
  for (const name of Object.keys(zip.files).sort()) {
    if (zip.files[name].dir) continue;
    let text = await zip.files[name].async('string');
    if (name === 'docProps/core.xml') text = text.replace(IGNORED, '');
    out.set(name, text);
  }
  return out;
}

const show = (v) => (v === null ? '(blank)' : v instanceof Date ? v.toISOString() : JSON.stringify(v));

function cellFacts(c) {
  return {
    value: show(cached(c)),
    formula: isFormula(c) ? '=' + c.formula : '',
    numFmt: c.numFmt || '',
    font: JSON.stringify(c.font || {}), fill: JSON.stringify(c.fill || {}),
    border: JSON.stringify(c.border || {}), alignment: JSON.stringify(c.alignment || {}),
  };
}

async function cellReport(a, b, limit) {
  const wa = new ExcelJS.Workbook(); await wa.xlsx.readFile(a);
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(b);
  const lines = [];
  const na = wa.worksheets.map((w) => w.name), nb = wb.worksheets.map((w) => w.name);
  if (na.join('|') !== nb.join('|')) lines.push(`Sheets differ: approved [${na.join(', ')}], new [${nb.join(', ')}]`);
  for (const sa of wa.worksheets) {
    const sb = wb.getWorksheet(sa.name); if (!sb) continue;
    const mA = JSON.stringify(Object.keys(sa._merges).sort()), mB = JSON.stringify(Object.keys(sb._merges).sort());
    if (mA !== mB) lines.push(`${sa.name}: merged cells differ`);
    const rows = Math.max(sa.rowCount, sb.rowCount), cols = Math.max(sa.columnCount, sb.columnCount);
    for (let c = 1; c <= cols; c++) if ((sa.getColumn(c).width || null) !== (sb.getColumn(c).width || null)) lines.push(`${sa.name}: column ${sa.getColumn(c).letter} width ${sa.getColumn(c).width} -> ${sb.getColumn(c).width}`);
    for (let r = 1; r <= rows; r++) {
      if ((sa.getRow(r).height || null) !== (sb.getRow(r).height || null)) lines.push(`${sa.name}: row ${r} height ${sa.getRow(r).height} -> ${sb.getRow(r).height}`);
      for (let c = 1; c <= cols; c++) {
        const fa = cellFacts(sa.getRow(r).getCell(c)), fb = cellFacts(sb.getRow(r).getCell(c));
        for (const k of Object.keys(fa)) if (fa[k] !== fb[k]) lines.push(`${sa.name}!${sa.getRow(r).getCell(c).address} ${k}: approved ${fa[k] || '(none)'} | new ${fb[k] || '(none)'}`);
        if (lines.length >= limit) return lines;
      }
    }
  }
  return lines;
}

(async () => {
  const [approved, candidate] = process.argv.slice(2);
  if (!approved || !candidate) { console.error('Usage: node test/golden.js <approved.xlsx> <new.xlsx>'); process.exit(2); }
  const pa = await parts(approved), pb = await parts(candidate);
  const names = [...new Set([...pa.keys(), ...pb.keys()])].sort();
  const diff = names.filter((n) => pa.get(n) !== pb.get(n));
  if (!diff.length) { console.log(`GOLDEN OK: identical to the approved output (${names.length} parts compared; only created/modified time ignored)`); return; }
  console.log('GOLDEN FAILED: output differs from the approved output');
  for (const n of diff) console.log('  part differs:', n, !pa.has(n) ? '(missing in approved)' : !pb.has(n) ? '(missing in new)' : '');
  const lines = await cellReport(approved, candidate, 50);
  for (const l of lines) console.log('  ' + l);
  if (!lines.length) console.log('  (no cell-level difference found; the difference is in workbook structure or settings)');
  process.exit(1);
})().catch((e) => { console.error(e.stack || e); process.exit(2); });
