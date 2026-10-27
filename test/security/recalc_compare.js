// Checks every formula's saved value in the tool's output against the same file recalculated by LibreOffice.
// Usage: node test/security/recalc_compare.js <tool_output.xlsx> <recalculated.xlsx>
// Port of the former recalc_compare.py (XLOOKUP cells are skipped, as before).
'use strict';
const ExcelJS = require('exceljs');
const { cached, isFormula, isNum, pyRepr } = require('../xlsx_values');

(async () => {
  const [mine, recalc] = process.argv.slice(2);
  const a = new ExcelJS.Workbook(); await a.xlsx.readFile(mine);    // formulas and the tool's own cached values
  const b = new ExcelJS.Workbook(); await b.xlsx.readFile(recalc);  // LibreOffice recalculated
  let n = 0, bad = 0;
  for (const ws of a.worksheets) {
    const wb2 = b.getWorksheet(ws.name);
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (c) => {
      if (!isFormula(c)) return;
      const f = '=' + c.formula;
      if (f.toUpperCase().includes('XLOOKUP')) return;
      n++;
      const va = cached(c); const vb = wb2 ? cached(wb2.getCell(c.address)) : null;
      if (isNum(va) && isNum(vb)) {
        if (Math.abs(va - vb) > 0.005) { bad++; console.log('MISMATCH', ws.name, c.address, f.slice(0, 60), va, vb); }
      } else {
        const same = va instanceof Date && vb instanceof Date ? va.getTime() === vb.getTime() : va === vb;
        const zeroish = (x) => x === null || x === 0 || x === false;
        if (!same && !(zeroish(va) && zeroish(vb))) { bad++; console.log('MISMATCH', ws.name, c.address, f.slice(0, 60), pyRepr(va), pyRepr(vb)); }
      }
    }));
  }
  console.log('formula cells checked', n, 'mismatches', bad);
})().catch((e) => { console.error(e.stack || e); process.exit(1); });
