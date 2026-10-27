// Reading Excel cells strictly, loading workbooks, and the shared output styles (header colour FBE2D5, thin borders).
// Copied from the approved Security engine.
'use strict';
const ExcelJS = require('exceljs'); // in the browser build this is the inlined ExcelJS (see src/core/browser/exceljs-global.js)
const { InputError, MAX_FILE_BYTES } = require('./common');

/* ---------- cell reading: distinguishes blank / number / text / error / formula with no saved value ---------- */
const VT = ExcelJS.ValueType;
function cellInfo(cell) {
  if (!cell) return { k: 'blank' };
  const t = cell.type;
  if (t === VT.Null) return { k: 'blank' };
  if (t === VT.Merge) return { k: 'blank' }; // non-master cell of a merged range: no value of its own
  if (t === VT.Formula) {
    const r = cell.model ? cell.model.result : undefined; // read from model: cell.value drops a result of 0
    if (r === undefined || r === null) return { k: 'nores' };
    if (typeof r === 'object' && r && r.error) return { k: 'err', v: r.error };
    if (typeof r === 'number') return { k: 'num', v: r };
    if (typeof r === 'boolean') return { k: 'bool', v: r };
    if (r instanceof Date) return { k: 'date', v: r };
    if (typeof r === 'string') return r.trim() === '' ? { k: 'blank' } : { k: 'text', v: r };
    return { k: 'err', v: 'unreadable formula result' };
  }
  const v = cell.value;
  if (t === VT.Number) return { k: 'num', v };
  if (t === VT.String || t === VT.SharedString) return (String(v).trim() === '') ? { k: 'blank' } : { k: 'text', v: String(v) };
  if (t === VT.RichText) { const s = (v.richText || []).map((p) => p.text).join(''); return s.trim() === '' ? { k: 'blank' } : { k: 'text', v: s }; }
  if (t === VT.Hyperlink) return { k: 'text', v: String(v.text != null ? v.text : v.hyperlink) };
  if (t === VT.Date) return { k: 'date', v };
  if (t === VT.Boolean) return { k: 'bool', v };
  if (t === VT.Error) return { k: 'err', v: v && v.error };
  return { k: 'err', v: 'unsupported cell type' };
}
const infoText = (i) => (i.k === 'text' ? i.v.trim() : i.k === 'num' ? String(i.v) : '');
/** Strict number: Excel number, or text that is a plain decimal number. Returns {ok, v (null for blank), why} */
function infoNumber(i) {
  if (i.k === 'blank') return { ok: true, v: null };
  if (i.k === 'num') return Number.isFinite(i.v) ? { ok: true, v: i.v } : { ok: false, why: 'not a finite number' };
  if (i.k === 'text') {
    const s = i.v.trim();
    if (/^[-+]?\d+(\.\d+)?$/.test(s)) return { ok: true, v: parseFloat(s) };
    return { ok: false, why: `text "${s.slice(0, 40)}" is not a number` };
  }
  if (i.k === 'err') return { ok: false, why: `Excel error ${i.v}` };
  if (i.k === 'nores') return { ok: false, why: 'formula with no saved value (open the file in Excel, let it calculate, save, and upload again)' };
  return { ok: false, why: `unexpected ${i.k} value` };
}
/** Strict whole number (cost-centre / account codes) */
function infoInteger(i) {
  const n = infoNumber(i);
  if (!n.ok || n.v === null) return n;
  if (!Number.isInteger(n.v)) return { ok: false, why: `${n.v} is not a whole number` };
  return n;
}

/** Dense rows: rows[r][c] = cellInfo (1-based) */
function sheetRows(ws) {
  const out = [];
  const maxCol = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const arr = new Array(maxCol + 1);
    for (let c = 1; c <= maxCol; c++) arr[c] = cellInfo(row.getCell(c));
    out[r] = arr;
  }
  return { rows: out, rowCount: ws.rowCount, colCount: maxCol, name: ws.name };
}

async function loadWorkbook(buf, label) {
  const bytes = buf.byteLength != null ? buf.byteLength : buf.length;
  if (bytes > MAX_FILE_BYTES) throw new InputError([`${label}: file is larger than 50 MB.`]);
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  if (!(u8[0] === 0x50 && u8[1] === 0x4b)) throw new InputError([`${label}: this is not an .xlsx file. If it is .xls or .csv, open it in Excel and "Save As" Excel Workbook (.xlsx).`]);
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buf); } catch (e) { throw new InputError([`${label}: the file could not be read as an Excel workbook (${e.message}).`]); }
  return wb;
}

/** A cell's original value, for copying into the output (Raw / Working sheets) */
function origValue(info) {
  if (!info) return null;
  switch (info.k) {
    case 'blank': return null;
    case 'num': case 'bool': case 'date': case 'text': return info.v;
    case 'err': return { error: info.v };
    default: return null;
  }
}

/* ---------- output styles (finance's workbook format) ---------- */
const FMT_AMT = '#,##0;-#,##0;"-"';
const FMT_AMT2 = '#,##0.00;-#,##0.00;"-"';
const THIN = { style: 'thin', color: { argb: 'FFC9CED6' } };
const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBE2D5' } };
const BOX = { top: THIN, bottom: THIN, left: THIN, right: THIN };
function formatTable(ws, headerRows, c1, c2, r1, r2) {
  for (const hr of headerRows) for (let c = c1; c <= c2; c++) {
    const cell = ws.getRow(hr).getCell(c);
    cell.fill = HEAD_FILL; cell.font = { bold: true }; cell.border = BOX;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  for (let r = r1; r <= r2; r++) { const row = ws.getRow(r); for (let c = c1; c <= c2; c++) row.getCell(c).border = BOX; }
}

module.exports = { ExcelJS, cellInfo, infoText, infoNumber, infoInteger, sheetRows, loadWorkbook, origValue, FMT_AMT, FMT_AMT2, THIN, HEAD_FILL, BOX, formatTable };
