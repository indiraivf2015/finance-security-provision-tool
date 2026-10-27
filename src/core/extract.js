// File 1: the Oracle ERP "IIVF Account Analysis" extract (SOP Steps 1-4). Copied from the approved Security engine;
// the account rule comes from the provision's settings, so non-contiguous account lists work (e.g. 51610, 51612-51614).
//
// spec = {
//   accounts: [51610, 51612, 51613, 51614],   every line's SEGMENT2 must be one of these
//   accountsText: '51610, 51612, 51613 and 51614',
//   glParams: { from, to } | null,            expected "GL Code From/To" report parameters; null = not checked
//   multipleAccountsAllowed: false,           may the "Multiple Accounts" report parameter be filled in?
// }
'use strict';
const { InputError, norm, parseMonthYear, parseDateText, fyOf, monthLabel, hasMoreThan2Decimals } = require('./common');
const { infoText, infoNumber, infoInteger, sheetRows } = require('./excel');

const EXTRACT_REQUIRED = {
  periodname: 'PERIOD_NAME', segment1: 'SEGMENT1', segment2: 'SEGMENT2', balancingsegmentdesc: 'BALANCING_SEGMENT_DESC',
  linedescription: 'LINE_DESCRIPTION', balance: 'BALANCE', jesourcename: 'JE_SOURCE_NAME', createdby: 'CREATED_BY',
  attribute3: 'ATTRIBUTE3', partyname: 'PARTY_NAME',
};

function readExtract(wb, spec) {
  const errors = [];
  const candidates = [];
  for (const ws of wb.worksheets) {
    const g = sheetRows(ws);
    for (let r = 1; r <= Math.min(60, g.rowCount); r++) {
      const keys = new Set(); for (let c = 1; c <= g.colCount; c++) keys.add(norm(infoText(g.rows[r][c])));
      if (keys.has('periodname') && keys.has('segment1') && keys.has('balance')) { candidates.push({ g, hdr: r, ws }); break; }
    }
  }
  if (candidates.length !== 1) {
    throw new InputError([candidates.length === 0
      ? 'ERP extract: no sheet has the Account Analysis column headings (PERIOD_NAME, SEGMENT1, BALANCE …). Upload the raw "IIVF Account Analysis" download (SOP Step 4).'
      : 'ERP extract: more than one sheet looks like an Account Analysis extract. Upload a file with only the raw download.']);
  }
  const { g, hdr, ws: rawSheet } = candidates[0];
  const headers = []; const col = Object.create(null);
  for (let c = 1; c <= g.colCount; c++) {
    const h = infoText(g.rows[hdr][c]);
    if (!h) continue;
    const k = norm(h);
    if (col[k] !== undefined) { if (EXTRACT_REQUIRED[k]) errors.push(`ERP extract: column ${EXTRACT_REQUIRED[k]} appears more than once.`); continue; }
    col[k] = c; headers.push({ c, name: h.trim(), k });
  }
  for (const [k, label] of Object.entries(EXTRACT_REQUIRED)) if (col[k] === undefined) errors.push(`ERP extract: column ${label} is missing.`);
  if (errors.length) throw new InputError(errors);

  // Optional parameter block (SOP Steps 2–3): "From Date / To Date / GL Code From / GL Code To"
  const params = readExtractParams(g, hdr, spec);

  const rows = [];
  for (let r = hdr + 1; r <= g.rowCount; r++) {
    const row = g.rows[r];
    let any = false; for (let c = 1; c <= g.colCount; c++) if (row[c].k !== 'blank') { any = true; break; }
    if (!any) continue;
    const rec = { erpRow: r, cells: row };
    const per = infoText(row[col.periodname]);
    const seg1 = infoInteger(row[col.segment1]);
    const seg2 = infoInteger(row[col.segment2]);
    const bal = infoNumber(row[col.balance]);
    if (!per) errors.push(`ERP extract row ${r}: PERIOD_NAME is blank.`);
    if (!seg1.ok || seg1.v === null) errors.push(`ERP extract row ${r}: SEGMENT1 (cost centre) ${seg1.ok ? 'is blank' : seg1.why}.`);
    if (!seg2.ok || seg2.v === null) errors.push(`ERP extract row ${r}: SEGMENT2 (account) ${seg2.ok ? 'is blank' : seg2.why}.`);
    if (!bal.ok || bal.v === null) errors.push(`ERP extract row ${r}: BALANCE ${bal.ok ? 'is blank' : bal.why}.`);
    else if (hasMoreThan2Decimals(bal.v)) errors.push(`ERP extract row ${r}: BALANCE ${bal.v} has more than 2 decimals.`);
    rec.period = per; rec.seg1 = seg1.v; rec.seg2 = seg2.v; rec.bal = bal.ok ? bal.v : null;
    rec.desc = infoText(row[col.linedescription]);
    rec.ccName = infoText(row[col.balancingsegmentdesc]);
    rec.ccNameRaw = row[col.balancingsegmentdesc].k === 'text' ? row[col.balancingsegmentdesc].v : (row[col.balancingsegmentdesc].k === 'num' ? String(row[col.balancingsegmentdesc].v) : '');
    rec.source = infoText(row[col.jesourcename]);
    rec.createdBy = infoText(row[col.createdby]);
    rec.attr3 = infoText(row[col.attribute3]);
    if (col.gldate !== undefined) { const gi = row[col.gldate]; rec.glDate = gi.k === 'date' ? { y: gi.v.getUTCFullYear(), m: gi.v.getUTCMonth() + 1, d: gi.v.getUTCDate() } : gi.k === 'text' ? parseDateText(gi.v) : null; }
    rec.party = infoText(row[col.partyname]);
    rows.push(rec);
  }
  if (rows.length === 0) errors.push('ERP extract: there are no data rows under the headings.');
  if (errors.length) throw new InputError(errors.slice(0, 50).concat(errors.length > 50 ? [`… and ${errors.length - 50} more problems.`] : []));

  // One accounting period only (SOP Step 3: first to last day of one month)
  const periods = [...new Set(rows.map((x) => x.period))];
  if (periods.length !== 1) throw new InputError([`ERP extract: it contains more than one period (${periods.join(', ')}). Download one month only (SOP Step 3).`]);
  const pm = parseMonthYear(periods[0]);
  if (!pm) throw new InputError([`ERP extract: PERIOD_NAME "${periods[0]}" is not a month (expected e.g. "Sep-26").`]);

  // Only the SOP's accounts (Step 2)
  const badAcc = [...new Set(rows.filter((x) => !spec.accounts.includes(x.seg2)).map((x) => x.seg2))].sort((a, b) => a - b);
  if (badAcc.length) throw new InputError([`ERP extract: it contains account(s) ${badAcc.join(', ')}. This working uses only ${spec.accountsText} (SOP Step 2). Re-run the report with the correct accounts.`]);

  const period = { label: periods[0].trim(), display: monthLabel(pm.y, pm.m), y: pm.y, m: pm.m, cur: fyOf(pm.y, pm.m), prev: (fyOf(pm.y, pm.m) + 99) % 100 };

  // Parameter block consistency, when present
  const paramErrors = params.problems.slice();
  if (params.from && (params.from.y !== pm.y || params.from.m !== pm.m || params.from.d !== 1)) paramErrors.push(`report "From Date" ${params.fromText} is not the first day of ${period.label}`);
  if (params.to) {
    const last = new Date(Date.UTC(pm.y, pm.m, 0)).getUTCDate();
    if (params.to.y !== pm.y || params.to.m !== pm.m || params.to.d !== last) paramErrors.push(`report "To Date" ${params.toText} is not the last day of ${period.label}`);
  }
  if (spec.glParams) {
    if (params.glFrom !== null && params.glFrom !== spec.glParams.from) paramErrors.push(`report "GL Code From" is ${params.glFrom}, SOP says ${spec.glParams.from}`);
    if (params.glTo !== null && params.glTo !== spec.glParams.to) paramErrors.push(`report "GL Code To" is ${params.glTo}, SOP says ${spec.glParams.to}`);
  }
  if (paramErrors.length) throw new InputError(paramErrors.map((e) => `ERP extract: ${e} (SOP Steps 2–3).`));

  return { sheetName: g.name, headerRow: hdr, headers, col, rows, period, params, rawSheet };
}

function readExtractParams(g, hdr, spec) {
  const out = { from: null, to: null, fromText: '', toText: '', glFrom: null, glTo: null, multipleAccounts: '', problems: [] };
  for (let r = 1; r < hdr && r < g.rowCount; r++) {
    const names = []; for (let c = 1; c <= g.colCount; c++) names[c] = norm(infoText(g.rows[r][c]));
    if (names.indexOf('fromdate') < 1) continue;
    const vals = g.rows[r + 1];
    const pick = (key) => { const i = names.indexOf(key); return i > 0 ? vals[i] : null; };
    const asDate = (info, label) => {
      if (!info || info.k === 'blank') return null;
      let d = null;
      if (info.k === 'date') d = { y: info.v.getUTCFullYear(), m: info.v.getUTCMonth() + 1, d: info.v.getUTCDate() };
      else if (info.k === 'text') d = parseDateText(info.v);
      else if (info.k === 'num' && Number.isInteger(info.v) && info.v > 20000 && info.v < 80000) { const dt = new Date(Date.UTC(1899, 11, 30) + info.v * 86400000); d = { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }; }
      if (!d) out.problems.push(`report "${label}" value "${info.k === 'text' ? info.v : String(info.v)}" could not be read as a date`);
      return d;
    };
    const f = pick('fromdate'), t = pick('todate');
    out.from = asDate(f, 'From Date'); out.fromText = f ? (f.k === 'text' ? f.v : String(f.v)) : '';
    out.to = asDate(t, 'To Date'); out.toText = t ? (t.k === 'text' ? t.v : String(t.v)) : '';
    const gf = pick('glcodefrom'), gt = pick('glcodeto');
    const gfn = gf ? infoInteger(gf) : null, gtn = gt ? infoInteger(gt) : null;
    out.glFrom = gfn && gfn.ok ? gfn.v : null; out.glTo = gtn && gtn.ok ? gtn.v : null;
    if (gf && gf.k !== 'blank' && !(gfn && gfn.ok)) out.problems.push('report "GL Code From" could not be read');
    if (gt && gt.k !== 'blank' && !(gtn && gtn.ok)) out.problems.push('report "GL Code To" could not be read');
    // The SOP extracts all centres; a centre filter would give an incomplete extract
    for (const [key, label] of [['centerfrom', 'Center From'], ['centerto', 'Center To']]) {
      const v = pick(key); if (v && v.k !== 'blank') out.problems.push(`report parameter "${label}" is filled in ("${infoText(v) || v.v}") — the SOP runs the report for all centres`);
    }
    const ma = pick('multipleaccounts');
    if (ma && ma.k !== 'blank') {
      out.multipleAccounts = infoText(ma) || String(ma.v);
      if (!spec.multipleAccountsAllowed) out.problems.push(`report parameter "Multiple Accounts" is filled in ("${out.multipleAccounts}") — the SOP uses only ${spec.accountsText}`);
    }
    break;
  }
  return out;
}

module.exports = { readExtract, EXTRACT_REQUIRED };
