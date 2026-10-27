// Builds made-up input files in the real layouts (no finance data), for the provision tests.
//   extract(…)   Oracle "IIVF Account Analysis" download: title, parameter block (rows 3-5), headings in row 9,
//                every value stored as text, as the real download does
//   statement(…) last month's statement: title in A2, headings in row 4
//   hkFile(…)    Rajeev's housekeeping file: sheet "HK <Mon-YY>", headings rows 1-3, data from row 4
'use strict';
const ExcelJS = require('exceljs');

const EXTRACT_COLS = ['PERIOD_NAME', 'BALANCING_SEGMENT_DESC', 'NATURAL_ACCOUNT_DESC', 'SEGMENT1', 'SEGMENT2', 'SEGMENT3', 'SEGMENT4', 'SEGMENT5', 'SEGMENT6', 'SEGMENT7', 'SEGMENT8', 'SEGMENT9', 'GL_DATE', 'LINE_DESCRIPTION', 'ACCOUNTED_DR', 'ACCOUNTED_CR', 'BALANCE', 'JE_SOURCE_NAME', 'GL_BATCH_NAME', 'ACCOUNTING_CLASS_CODE', 'ACCOUNTING_CLASS_NAME', 'ATTRIBUTE1', 'ATTRIBUTE2', 'ATTRIBUTE3', 'ATTRIBUTE4', 'ATTRIBUTE5', 'TRANSACTION_NUMBER', 'TRANSACTION_DATE', 'PARTY_NUMBER', 'PARTY_NAME', 'PARTY_SITE_NAME', 'GL_JE_NAME', 'POSTED_DATE', 'HEADER_DESCRIPTION', 'LAST_UPDATE_DATE', 'CREATION_DATE', 'CREATED_BY'];

/**
 * p = { period: 'Sep-26', from: '2026-09-01', to: '2026-09-30', glFrom: '51610', glTo: '51614', multiple: '',
 *       lines: [{ code, name, account, bal, desc, source = 'IIVF P2P', by = 'ERP.Integrator', attr3 = 'Vendor', party = '' }] }
 */
async function extract(p) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet1');
  ws.getRow(1).values = ['IIVF Account Analysis Report'];
  ws.getRow(3).values = ['Parameter'];
  ws.getRow(4).values = ['From Date', 'To Date', 'GL Code From', 'GL Code To', 'Center From', 'Center To', 'Multiple Accounts'];
  ws.getRow(5).values = [p.from + 'T00:00:00.000+00:00', p.to + 'T00:00:00.000+00:00', p.glFrom || null, p.glTo || null, null, null, p.multiple || null];
  ws.getRow(9).values = EXTRACT_COLS;
  p.lines.forEach((l, i) => {
    const v = {
      PERIOD_NAME: l.period || p.period, BALANCING_SEGMENT_DESC: l.name, NATURAL_ACCOUNT_DESC: 'Housekeeping Expenses', SEGMENT1: String(l.code), SEGMENT2: String(l.account),
      GL_DATE: '15-' + (p.from.slice(5, 7)) + '-' + p.from.slice(0, 4), LINE_DESCRIPTION: l.desc, BALANCE: String(l.bal),
      ACCOUNTED_DR: l.bal > 0 ? String(l.bal) : '0', ACCOUNTED_CR: l.bal < 0 ? String(-l.bal) : '0',
      JE_SOURCE_NAME: l.source || 'IIVF P2P', ATTRIBUTE3: l.attr3 === undefined ? 'Test Vendor Pvt Ltd' : l.attr3, PARTY_NAME: l.party || null, CREATED_BY: l.by || 'ERP.Integrator',
    };
    ws.getRow(10 + i).values = EXTRACT_COLS.map((c) => (v[c] === undefined || v[c] === '' ? null : v[c]));
  });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** p = { title, headers: [...row 4 from column B], rows: [[code, name, ...amounts]] } */
async function statement(p) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Summary');
  ws.getCell('A2').value = p.title;
  ws.getRow(4).values = [null, ...p.headers];
  p.rows.forEach((r, i) => { ws.getRow(5 + i).values = [null, ...r]; });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** p = { sheet: 'HK Sep-26', rows: [{ code, name, vendor, manpower, materials, rent, type }], extraSheets: ['Security Sep-26'] } */
async function hkFile(p) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet(p.sheet);
  ws.getRow(1).values = ['Housekeeping Provision', null, null, null, null, null, null, null, 'Manpower ', 'Materials (Only supplied by Vcare Vendor)', 'Machine Rent ', 'Total Provision with GST'];
  const n = 3 + p.rows.length;
  ws.getRow(2).values = [null, null, null, null, null, null, null, null, { formula: `SUM(I4:I${n})`, result: p.rows.reduce((a, r) => a + (r.manpower || 0), 0) }];
  ws.getRow(3).values = ['S. No. ', 'Center Name ', 'Vendor Name ', 'Storage Code ', 'SQFT', 'State', null, null, 'Per Month Cost with GST', 'Per month cost with GST'];
  p.rows.forEach((r, i) => { ws.getRow(4 + i).values = [i + 1, r.name, r.vendor || 'Test Vendor', r.code, 1000, 'STATE', 'WEST', r.type || 'Main', r.manpower === undefined ? null : r.manpower, r.materials === undefined ? null : r.materials, r.rent || null]; });
  for (const s of p.extraSheets || []) { const x = wb.addWorksheet(s); x.getRow(1).values = ['S.No.', 'Center Name ', 'Vendor Name ', 'Storage Code ']; }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

module.exports = { extract, statement, hkFile, EXTRACT_COLS };
