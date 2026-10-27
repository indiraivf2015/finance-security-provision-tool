// Helpers shared by the Node comparison scripts (ported from the former Python/openpyxl versions).
// cached(cell) returns what openpyxl gives with data_only=True: the saved value, or for a formula its saved
// result (null when the file has none). ExcelJS drops a formula result of 0 from cell.value, so read cell.model.result.
'use strict';

function plain(v) {
  if (v === undefined) return null;
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if (v.error) return v.error;                                   // e.g. '#N/A'
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text !== undefined && v.hyperlink !== undefined) return v.text;
  }
  return v;
}

function isFormula(cell) {
  const v = cell.value;
  return !!(v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v));
}

function cached(cell) {
  if (isFormula(cell)) return plain(cell.model.result);
  return plain(cell.value);
}

// Row as a 0-based array of cached values, columns A..maxCol (like openpyxl iter_rows(values_only=True))
function rowValues(ws, r, maxCol) {
  const row = ws.getRow(r); const out = [];
  for (let c = 1; c <= maxCol; c++) out.push(cached(row.getCell(c)));
  return out;
}

const isNum = (x) => typeof x === 'number' && Number.isFinite(x);

// Python round(x, 2): correctly rounded on the exact binary value, ties to even
function pyRound2(x) {
  if (!Number.isFinite(x) || x === 0) return x;
  const neg = x < 0; const s = Math.abs(x).toFixed(80);              // exact decimal expansion of the double
  const [ip, fp] = s.split('.');
  const keep = BigInt(ip + fp.slice(0, 2)); const rest = fp.slice(2);
  const half = '5'.padEnd(rest.length, '0');
  let n = keep;
  if (rest > half || (rest === half && keep % 2n === 1n)) n += 1n;
  const t = n.toString().padStart(3, '0');
  const r = Number(t.slice(0, -2) + '.' + t.slice(-2));
  return neg ? -r : r;
}

// Python-style printing, so the output reads as before
function pyFloat(x) { if (!Number.isFinite(x)) return String(x); const s = String(x); return /[.e]/.test(s) ? s : s + '.0'; }
function pyRepr(v) {
  if (v === null || v === undefined) return 'None';
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (typeof v === 'number') return String(v);
  if (v instanceof Date) return 'datetime(' + v.toISOString() + ')';
  if (typeof v === 'string') return v.includes("'") && !v.includes('"') ? '"' + v + '"' : "'" + v.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
  if (Array.isArray(v)) return '[' + v.map(pyRepr).join(', ') + ']';
  return String(v);
}
function pySet(arr) { return arr.length ? '{' + arr.slice().sort((a, b) => (a > b ? 1 : a < b ? -1 : 0)).map(pyRepr).join(', ') + '}' : 'set()'; }

module.exports = { cached, isFormula, rowValues, isNum, pyRound2, pyFloat, pyRepr, pySet };
