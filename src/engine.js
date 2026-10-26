/*
 * Security Provision Working — calculation engine
 * Implements SOP "Monthly Security Provision Working" v1.0 (October 2026), steps 5–16 and the Section 04 checks.
 *
 * Design rules
 *  - The SOP is the specification. Where the SOP is silent, the tool does NOT guess: it either applies the
 *    behaviour Excel itself would give in the manual process (documented in RULES below) or asks the user.
 *  - Nothing in this file refers to any specific month, centre, vendor or amount.
 *  - All amounts are carried at full precision; nothing is rounded in the calculation.
 *  - Runs unchanged in the browser and in Node (for testing). No network access.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('exceljs'));
  else root.SecProv = factory(root.ExcelJS);
})(typeof self !== 'undefined' ? self : this, function (ExcelJS) {
  'use strict';

  /* ------------------------------------------------------------------ *
   * Fixed parameters taken from the SOP
   * ------------------------------------------------------------------ */
  const SOP = Object.freeze({
    version: 'SOP v1.0 · October 2026',
    accountFrom: 51701,                // Step 2
    accountTo: 51704,                  // Step 2
    p2pSources: ['iivf p2p', 'payables'], // Step 7 — P2P invoice bookings (ERP labels them "IIVF P2P" or "Payables")
    gstReclassCreator: 'taxverification2@indiraivf.in', // Step 8
    gstDivisor: 118,                   // Step 11  Basic = PWG ÷ 118 × 100
    gstBasicFactor: 100,
    finalFactor: 1.135,                // Step 12  Final = Basic × 1.135
    negCfTopN: 10,                     // Finance rule (Oct 2026): negative Pro CF absorbed by the 10 highest centres
  });

  const TOOL_VERSION = '1.1.0';
  const MAX_FILE_BYTES = 50 * 1024 * 1024;
  const EPS = 0.005; // half a paisa — tolerance for comparing floating-point rupee totals

  /* ------------------------------------------------------------------ *
   * Small helpers
   * ------------------------------------------------------------------ */
  class InputError extends Error {
    constructor(list) { super(list.join('\n')); this.list = list; this.name = 'InputError'; }
  }

  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const pad2 = (n) => String(n).padStart(2, '0');
  const MONTH_INDEX = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const MONTH_ABBR = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function monthFromWord(w) {
    const k = String(w).toLowerCase().slice(0, 3);
    return MONTH_INDEX[k] || null;
  }
  /** Indian financial year (Apr–Mar), as the 2-digit year in which it ends: Apr-2026..Mar-2027 → 27 */
  function fyOf(y, m) { return ((m >= 4 ? y + 1 : y) % 100); }
  const fyLabelWorking = (fy) => 'FY ' + pad2(fy);   // SOP Step 6: "FY 26" / "FY 27"
  const fyLabelHeader = (fy) => 'FY' + pad2(fy);     // Statement headers: "FY27 Pro" …
  const ym = (y, m) => y * 12 + (m - 1);
  function prevMonth(y, m) { return m === 1 ? { y: y - 1, m: 12 } : { y, m: m - 1 }; }
  function colLetter(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  const toPaise = (x) => Math.round(x * 100);
  const near = (a, b) => Math.abs(a - b) < EPS;
  const hasMoreThan2Decimals = (x) => Math.abs(x * 100 - Math.round(x * 100)) > 1e-6;

  /* ------------------------------------------------------------------ *
   * Cell reading — distinguishes blank / number / text / error / formula-with-no-saved-value
   * ------------------------------------------------------------------ */
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
  /** Strict number: Excel number, or text that is a plain decimal number. Returns {ok, v(null for blank), why} */
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

  function sheetRows(ws) {
    // Returns dense array of rows; each row = array (1-based col index) of cellInfo
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

  /* ------------------------------------------------------------------ *
   * Period text parsing  ("Sep-26", "Sept-26", "September 2026")
   * ------------------------------------------------------------------ */
  function parseMonthYear(s) {
    const m = String(s).trim().match(/^([A-Za-z]{3,9})[\s\-'’.]*(\d{2}|\d{4})$/);
    if (!m) return null;
    const mo = monthFromWord(m[1]);
    if (!mo) return null;
    let y = parseInt(m[2], 10); if (y < 100) y += 2000;
    return { y, m: mo };
  }

  /* ------------------------------------------------------------------ *
   * Step 6 — Exp Year tagging from the Line Description narration
   * ------------------------------------------------------------------ */
  const MON = '(january|february|march|april|may|june|july|august|september|october|november|december|sept|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)';
  const RE_MONWORD = new RegExp('(?<![a-z])' + MON + '(?![a-z])', 'gi');
  const RE_PERIOD = /period\D{0,20}?(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})\s*(?:to|till|upto|up to|-|–)\s*(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/i;
  // Text immediately before a month that marks it as the SERVICE month: "for the month of", "month of", "for", "service period of" …
  const RE_ANCHOR_BEFORE = /(?:month\s*(?:of)?|\bfor(?:\s+the)?(?:\s+month)?(?:\s+of)?|service\s+(?:month|period)?\s*(?:of|for)?)\s*[:\-]?\s*$/i;
  // Text immediately before a month that marks it as a BILL / document date: "dt 10-", "dated", "invoice date" …
  const RE_DATE_BEFORE = /(?:\bdt|\bdtd|\bdated|\bdate|\bdatd)\.?\s*[:\-]?\s*(?:\d{1,2}(?:st|nd|rd|th)?[\s\-\/.]*)?$/i;
  const RE_MON_MONTH = new RegExp('(?<![a-z])' + MON + '\\s*[\\-\']?\\s*month(?![a-z])', 'i'); // "MARCH MONTH INVOICE"
  const RE_BILLDATE = /(?:\bdt|\bdtd|\bdated|\bdate|\bdatd)\.?\s*[:\-]?\s*(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/i;
  const RE_FYMARK = /(?<![a-z])fy\s*[-:'’]?\s*(20\d{2}|\d{2})(?:\s*[-–\/]\s*(\d{2}))?(?!\d)/i;

  function validDate(d, m, y) {
    if (m < 1 || m > 12 || d < 1) return false;
    const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return d <= dim;
  }

  /** Every month word in the narration, classified. */
  function monthTokens(d, period) {
    const out = [];
    let mm; RE_MONWORD.lastIndex = 0;
    while ((mm = RE_MONWORD.exec(d)) !== null) {
      const start = mm.index, end = start + mm[0].length;
      const mo = monthFromWord(mm[1]);
      const rest = d.slice(end);
      const before = start > 0 ? d[start - 1] : '';
      const pre = d.slice(Math.max(0, start - 45), start);
      const t = { word: mm[1], m: mo, y: null, kind: 'noyear', anchored: false, span: [start, end] };
      let x;
      if ((x = rest.match(/^[\s\-'’.,]*(20\d{2})\s*[-–\/]\s*(\d{2})(?!\d)/)) && (parseInt(x[1], 10) + 1) % 100 === parseInt(x[2], 10)) {
        const fyEnd = parseInt(x[1], 10) + 1; t.y = mo >= 4 ? fyEnd - 1 : fyEnd; t.kind = 'fyform';               // "June 2026-27"
      } else if ((x = rest.match(/^[\s\-'’.,]*(\d{2})\s*[-–]\s*(\d{2})(?!\d)/)) && (parseInt(x[1], 10) + 1) % 100 === parseInt(x[2], 10)) {
        const fyEnd = 2000 + parseInt(x[2], 10); t.y = mo >= 4 ? fyEnd - 1 : fyEnd; t.kind = 'fyform';               // "Mar 26-27"
      } else if ((x = rest.match(/^[\s\-'’.,]*(\d{1,2})(?:st|nd|rd|th)?[\s,]+(20\d{2})(?!\d)/))) {
        t.y = parseInt(x[2], 10); t.kind = 'date';                                                                    // "Sep 25, 2026"
      } else if ((x = rest.match(/^[\s\-'’.,\/]*(20\d{2}|\d{2})(?!\d)/))) {
        let y = parseInt(x[1], 10); if (x[1].length === 2) y += 2000;
        const after = rest.slice(x[0].length, x[0].length + 1);
        if (y >= period.y - 5 && y <= period.y + 1) { t.y = y; t.kind = 'year'; if (after === '/' || after === '_') t.kind = 'ref'; }
      }
      if (/[\/_\\0-9]/.test(before) || (before === '-' && /[a-z0-9]-$/i.test(pre))) { if (t.kind !== 'noyear') t.kind = t.kind === 'date' ? 'date' : 'ref'; t.refLike = true; }
      if (RE_DATE_BEFORE.test(pre) || (before === '-' && /\d-$/.test(pre))) { if (t.y !== null) t.kind = 'date'; t.dateLike = true; }
      t.anchored = !t.refLike && !t.dateLike && RE_ANCHOR_BEFORE.test(pre);
      if (t.y !== null) t.fy = fyOf(t.y, t.m);
      out.push(t);
    }
    return out;
  }

  /**
   * Proposes the financial year for one booking line (SOP Step 6).
   * Returns { fy|null, basis, sure }.
   *  sure=true  → the narration states the service month/period explicitly ("for the month of June 2026",
   *               "month of January 2025-26", "Service Period 01-06-2026 to 30-06-2026") AND nothing else in the
   *               narration points to a different financial year. Tagged automatically.
   *  sure=false → anything else. The user must choose; a suggestion is shown only when the narration itself gives one.
   */
  function proposeFY(desc, period) {
    const d = String(desc || '');
    const pYM = ym(period.y, period.m);
    const allowed = new Set([period.cur, period.prev]);
    const res = (fy, basis, sure) => {
      if (fy !== null && !allowed.has(fy)) {
        return { fy: null, basis: `${basis} → FY ${pad2(fy)}, which is neither ${fyLabelWorking(period.prev)} nor ${fyLabelWorking(period.cur)}`, sure: false };
      }
      return { fy, basis, sure };
    };
    // Finance rule (confirmed by the finance team): narration "Journal Import Created" = current financial year of the posting month
    if (d.trim().replace(/^["'“”]+|["'“”]+$/g, '').trim().toLowerCase() === 'journal import created') {
      return res(period.cur, `Finance rule: "Journal Import Created" is tagged to the current financial year (${fyLabelWorking(period.cur)})`, true);
    }
    const fm = d.match(RE_FYMARK);
    const fyMark = fm ? (fm[2] ? parseInt(fm[2], 10) : parseInt(fm[1], 10) % 100) : null;
    const toks = monthTokens(d, period);

    // Service-month candidates: explicit service period, or a month WITH a year right after "for the month of" / "for" / "month of"
    const cands = [];
    let ambiguous = null;
    const p = d.match(RE_PERIOD);
    if (p) {
      const [d1, m1, y1, d2, m2, y2] = p.slice(1).map((x) => parseInt(x, 10));
      if (validDate(d1, m1, y1) && validDate(d2, m2, y2)) {
        // dd-mm vs mm-dd: if the month-first reading is also a valid pair of dates and gives a different year, it is ambiguous
        const alt = validDate(m1, d1, y1) && validDate(m2, d2, y2);
        if (alt && (fyOf(y1, d1) !== fyOf(y1, m1) || fyOf(y2, d2) !== fyOf(y2, m2))) ambiguous = 'the service-period dates can be read day-first or month-first';
        if (fyOf(y1, m1) !== fyOf(y2, m2)) ambiguous = 'the service period spans two financial years';
        cands.push({ y: y1, m: m1, fy: fyOf(y1, m1), text: `${p[1]}-${p[2]}-${p[3]} to ${p[4]}-${p[5]}-${p[6]}`, period: true, y2, m2 });
      }
    }
    for (const t of toks) if (t.anchored && (t.kind === 'year' || t.kind === 'fyform')) cands.push({ y: t.y, m: t.m, fy: t.fy, text: `${MONTH_ABBR[t.m]}-${t.y}` });

    if (cands.length) {
      const label = [...new Set(cands.map((c) => c.text))].join(', ');
      const fys = new Set(cands.map((c) => c.fy));
      const future = cands.some((c) => ym(c.y, c.m) > pYM || (c.period && ym(c.y2, c.m2) > pYM));
      const otherDated = toks.filter((t) => !(t.anchored && (t.kind === 'year' || t.kind === 'fyform')) && t.y !== null);
      const otherUndated = toks.filter((t) => t.y === null && !cands.some((c) => c.m === t.m));
      let why = null;
      if (future) why = `service month ${label} is later than the posting month`;
      else if (ambiguous) why = ambiguous;
      else if (fys.size > 1) why = `the narration gives service months in different financial years (${label})`;
      else if (otherDated.some((t) => t.fy !== [...fys][0])) why = `another date in the narration (${otherDated.map((t) => t.word + (t.y ? ' ' + t.y : '')).join(', ')}) is in a different financial year`;
      else if (otherUndated.length) why = `the narration also mentions ${otherUndated.map((t) => t.word).join(', ')} without a year`;
      else if (fyMark !== null && fyMark !== [...fys][0]) why = `the narration also says FY ${pad2(fyMark)}`;
      if (!why) return res([...fys][0], `Service ${cands.some((c) => c.period) ? 'period' : 'month'} ${label}`, true);
      // A bill/document date in another year does not cancel the suggestion (e.g. March service billed in April); anything else does
      const sugg = (fys.size === 1 && !future && !ambiguous && !otherDated.some((t) => t.kind !== 'date' && t.fy !== [...fys][0]) && (fyMark === null || fyMark === [...fys][0])) ? [...fys][0] : null;
      return res(sugg, `Please check: ${why}` + (sugg !== null ? ` — suggested ${fyLabelWorking(sugg)}` : ''), false);
    }

    // No explicit service month → the user decides. Offer a suggestion only when the narration itself contains one.
    const dated = toks.filter((t) => t.y !== null);
    const datedFys = new Set(dated.map((t) => t.fy));
    let sugg = null, basis = 'No service month in the narration';
    const nm = d.match(new RegExp('(?:month\\s*(?:of)?|\\bfor(?:\\s+the)?(?:\\s+month)?(?:\\s+of)?)\\s*[:\\-]?\\s*' + MON + '(?![a-z])', 'i')) || d.match(RE_MON_MONTH);
    if (nm) {
      const mo = monthFromWord(nm[1]); const y = mo <= period.m ? period.y : period.y - 1;
      sugg = fyOf(y, mo); basis = `Month "${nm[1]}" without a year — suggested ${MONTH_ABBR[mo]}-${y}, the latest ${MONTH_ABBR[mo]} up to the posting month`;
      if (dated.some((t) => t.fy !== sugg)) { sugg = null; basis = `Month "${nm[1]}" without a year, and another date in a different financial year`; }
    } else if (dated.length) {
      const undatedOther = toks.some((t) => t.y === null && !dated.some((x) => x.m === t.m));
      if (datedFys.size === 1 && !undatedOther && !dated.some((t) => ym(t.y, t.m) > pYM)) { sugg = [...datedFys][0]; basis = `A month is mentioned (${dated.map((t) => t.word + ' ' + t.y).join(', ')}) but not as the service month — suggested from it`; }
      else basis = `Months mentioned (${dated.map((t) => t.word + ' ' + t.y).join(', ')}) but not as the service month`;
    } else {
      const bd = d.match(RE_BILLDATE);
      if (bd) { const [dd, mo, y] = bd.slice(1).map((x) => parseInt(x, 10)); if (validDate(dd, mo, y)) { sugg = fyOf(y, mo); basis = `Only a bill date (${bd[1]}-${bd[2]}-${bd[3]}) — suggested from the bill date`; } }
    }
    if (fyMark !== null) {
      if (sugg === null) { sugg = fyMark; basis += ` — narration says FY ${pad2(fyMark)}`; }
      else if (sugg !== fyMark) { sugg = null; basis += ` — but the narration also says FY ${pad2(fyMark)}`; }
    }
    return res(sugg, basis, false);
  }

  /* ------------------------------------------------------------------ *
   * File 1 — Oracle ERP extract (Steps 1–4 output, then Steps 5–9)
   * ------------------------------------------------------------------ */
  const EXTRACT_REQUIRED = {
    periodname: 'PERIOD_NAME', segment1: 'SEGMENT1', segment2: 'SEGMENT2', balancingsegmentdesc: 'BALANCING_SEGMENT_DESC',
    linedescription: 'LINE_DESCRIPTION', balance: 'BALANCE', jesourcename: 'JE_SOURCE_NAME', createdby: 'CREATED_BY',
    attribute3: 'ATTRIBUTE3', partyname: 'PARTY_NAME',
  };

  function readExtract(wb) {
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
    const params = readExtractParams(g, hdr);

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

    // Accounts within 51701–51704 (SOP Step 2)
    const badAcc = [...new Set(rows.filter((x) => x.seg2 < SOP.accountFrom || x.seg2 > SOP.accountTo).map((x) => x.seg2))];
    if (badAcc.length) throw new InputError([`ERP extract: it contains account(s) ${badAcc.join(', ')}, outside ${SOP.accountFrom}–${SOP.accountTo} (SOP Step 2). Re-run the report with the correct account range.`]);

    const period = { label: periods[0].trim(), display: (pm.m === 9 ? 'Sept' : MONTH_ABBR[pm.m]) + '-' + pad2(pm.y % 100), y: pm.y, m: pm.m, cur: fyOf(pm.y, pm.m), prev: (fyOf(pm.y, pm.m) + 99) % 100 };

    // Parameter block consistency, when present
    const paramErrors = params.problems.slice();
    if (params.from && (params.from.y !== pm.y || params.from.m !== pm.m || params.from.d !== 1)) paramErrors.push(`report "From Date" ${params.fromText} is not the first day of ${period.label}`);
    if (params.to) {
      const last = new Date(Date.UTC(pm.y, pm.m, 0)).getUTCDate();
      if (params.to.y !== pm.y || params.to.m !== pm.m || params.to.d !== last) paramErrors.push(`report "To Date" ${params.toText} is not the last day of ${period.label}`);
    }
    if (params.glFrom !== null && params.glFrom !== SOP.accountFrom) paramErrors.push(`report "GL Code From" is ${params.glFrom}, SOP says ${SOP.accountFrom}`);
    if (params.glTo !== null && params.glTo !== SOP.accountTo) paramErrors.push(`report "GL Code To" is ${params.glTo}, SOP says ${SOP.accountTo}`);
    if (paramErrors.length) throw new InputError(paramErrors.map((e) => `ERP extract: ${e} (SOP Steps 2–3).`));

    return { sheetName: g.name, headerRow: hdr, headers, col, rows, period, params, rawSheet };
  }

  function parseDateText(s) {
    let m = String(s).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = String(s).trim().match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
    if (m) return { y: +m[3], m: +m[2], d: +m[1] };
    m = String(s).trim().match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{2}|\d{4})$/); // 01-SEP-2026
    if (m && monthFromWord(m[2])) { let y = +m[3]; if (y < 100) y += 2000; return { y, m: monthFromWord(m[2]), d: +m[1] }; }
    return null;
  }

  function readExtractParams(g, hdr) {
    const out = { from: null, to: null, fromText: '', toText: '', glFrom: null, glTo: null, problems: [] };
    for (let r = 1; r < hdr - 0 && r < g.rowCount; r++) {
      const names = []; for (let c = 1; c <= g.colCount; c++) names[c] = norm(infoText(g.rows[r][c]));
      const iFrom = names.indexOf('fromdate');
      if (iFrom < 1) continue;
      const vals = g.rows[r + 1];
      const pick = (key) => { const i = names.indexOf(key); return i > 0 ? vals[i] : null; };
      const asDate = (info, label) => {
        if (!info || info.k === 'blank') return null;
        let r = null;
        if (info.k === 'date') r = { y: info.v.getUTCFullYear(), m: info.v.getUTCMonth() + 1, d: info.v.getUTCDate() };
        else if (info.k === 'text') r = parseDateText(info.v);
        else if (info.k === 'num' && Number.isInteger(info.v) && info.v > 20000 && info.v < 80000) { const dt = new Date(Date.UTC(1899, 11, 30) + info.v * 86400000); r = { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() }; }
        if (!r) out.problems.push(`report "${label}" value "${info.k === 'text' ? info.v : String(info.v)}" could not be read as a date`);
        return r;
      };
      const f = pick('fromdate'), t = pick('todate');
      out.from = asDate(f, 'From Date'); out.fromText = f ? (f.k === 'text' ? f.v : String(f.v)) : '';
      out.to = asDate(t, 'To Date'); out.toText = t ? (t.k === 'text' ? t.v : String(t.v)) : '';
      const gf = pick('glcodefrom'), gt = pick('glcodeto');
      const gfn = gf ? infoInteger(gf) : null, gtn = gt ? infoInteger(gt) : null;
      out.glFrom = gfn && gfn.ok ? gfn.v : null; out.glTo = gtn && gtn.ok ? gtn.v : null;
      if (gf && gf.k !== 'blank' && !(gfn && gfn.ok)) out.problems.push('report "GL Code From" could not be read');
      if (gt && gt.k !== 'blank' && !(gtn && gtn.ok)) out.problems.push('report "GL Code To" could not be read');
      // The SOP extracts all centres for the account range; a centre filter or extra account list would give an incomplete extract
      for (const [key, label] of [['centerfrom', 'Center From'], ['centerto', 'Center To'], ['multipleaccounts', 'Multiple Accounts']]) {
        const v = pick(key); if (v && v.k !== 'blank') out.problems.push(`report parameter "${label}" is filled in ("${infoText(v) || v.v}") — the SOP runs the report for all centres and only the account range ${SOP.accountFrom}–${SOP.accountTo}`);
      }
      break;
    }
    return out;
  }

  /* ------------------------------------------------------------------ *
   * File 2 — last month's provision statement (Step 13)
   * ------------------------------------------------------------------ */
  function readOpening(wb, period, opts) {
    const ytdKey = 'ytd' + norm(fyLabelHeader(period.cur));      // e.g. "ytdfy27"
    const cfKey = norm(fyLabelHeader(period.prev)) + 'procf';     // e.g. "fy26procf"
    const found = [];
    for (const ws of wb.worksheets) {
      const g = sheetRows(ws);
      for (let r = 1; r <= Math.min(30, g.rowCount); r++) {
        const keys = []; for (let c = 1; c <= g.colCount; c++) keys[c] = norm(infoText(g.rows[r][c]));
        const cc = keys.indexOf('costcenter') > 0 ? keys.indexOf('costcenter') : keys.indexOf('costcentre');
        if (cc < 1) continue;
        const ytdCols = []; keys.forEach((k, c) => { if (k === ytdKey) ytdCols.push(c); });
        const cfCols = []; keys.forEach((k, c) => { if (k === cfKey) cfCols.push(c); });
        if (!ytdCols.length && !cfCols.length) continue;
        let nameCol = keys.indexOf('costcentername'); if (nameCol < 1) nameCol = keys.indexOf('costcentrename');
        found.push({ g, hdr: r, cc, nameCol, ytdCols, cfCols });
        break;
      }
    }
    if (!found.length) {
      throw new InputError([`Last month's statement: no sheet has the headings "Cost Center", "YTD ${fyLabelHeader(period.cur)}" and "${fyLabelHeader(period.prev)} Pro CF". ` +
        `The SOP (Step 13) carries "YTD ${fyLabelHeader(period.cur)}" and "${fyLabelHeader(period.prev)} Pro CF" from last month. ` +
        (period.m === 4 ? 'This is April — the start of a new financial year. The SOP does not describe the year-end roll-over, so the tool cannot proceed without finance\'s rule.' : 'Check that the right file was uploaded.')]);
    }
    let pickSheet = found;
    if (found.length > 1) {
      if (!opts.openingSheet) return { needChoice: { kind: 'openingSheet', options: found.map((f) => f.g.name) } };
      pickSheet = found.filter((f) => f.g.name === opts.openingSheet);
      if (!pickSheet.length) throw new InputError([`Last month's statement: sheet "${opts.openingSheet}" not found.`]);
    }
    const s = pickSheet[0];
    const errs = [];
    if (s.cfCols.length !== 1) errs.push(`Last month's statement (sheet "${s.g.name}"): expected one "${fyLabelHeader(period.prev)} Pro CF" column, found ${s.cfCols.length}.`);
    if (s.ytdCols.length === 0) errs.push(`Last month's statement (sheet "${s.g.name}"): no "YTD ${fyLabelHeader(period.cur)}" column.` +
      (period.m === 4 ? ' This is April — the start of a new financial year. The SOP does not describe the year-end roll-over, so the tool cannot proceed without finance\'s rule.' : ''));
    if (errs.length) throw new InputError(errs);

    // Which month is this statement for?  Must be the month before the ERP period.
    const pm = prevMonth(period.y, period.m);
    let stmtMonth = null, stmtTitle = '';
    for (let r = 1; r < s.hdr && !stmtMonth; r++) for (let c = 1; c <= s.g.colCount && !stmtMonth; c++) {
      const t = infoText(s.g.rows[r][c]);
      const mt = t.match(/month\s+of\s+([A-Za-z]{3,9})[\s\-'’.]*(\d{2,4})\b/i);
      if (mt) { const p = parseMonthYear(mt[1] + '-' + mt[2]); if (p) { stmtMonth = p; stmtTitle = t.trim(); } }
    }
    if (stmtMonth && (stmtMonth.y !== pm.y || stmtMonth.m !== pm.m)) {
      throw new InputError([`Last month's statement is titled "${stmtTitle}", but the ERP extract is for ${period.label}. Upload the statement for ${MONTH_ABBR[pm.m]}-${pad2(pm.y % 100)}.`]);
    }
    if (!stmtMonth && !opts.confirmOpeningMonth) return { needChoice: { kind: 'confirmOpeningMonth', month: `${MONTH_ABBR[pm.m]}-${pad2(pm.y % 100)}`, sheet: s.g.name } };

    // Which YTD column is the closing balance? If more than one, the user must say.
    let ytdCol = s.ytdCols[0];
    const totalsOf = (c) => { let t = 0; for (let r = s.hdr + 1; r <= s.g.rowCount; r++) { const n = infoNumber(s.g.rows[r][c]); if (n.ok && n.v !== null && infoInteger(s.g.rows[r][s.cc]).v != null) t += n.v; } return t; };
    if (s.ytdCols.length > 1) {
      if (!opts.openingYtdCol) {
        return { needChoice: { kind: 'openingYtdCol', sheet: s.g.name, options: s.ytdCols.map((c) => ({ col: colLetter(c), total: totalsOf(c), header: infoText(s.g.rows[s.hdr][c]) })) } };
      }
      const want = s.ytdCols.find((c) => colLetter(c) === opts.openingYtdCol);
      if (!want) throw new InputError([`Last month's statement: column ${opts.openingYtdCol} is not a YTD ${fyLabelHeader(period.cur)} column.`]);
      ytdCol = want;
    }
    const cfCol = s.cfCols[0];

    const centres = new Map();
    const order = [];
    for (let r = s.hdr + 1; r <= s.g.rowCount; r++) {
      const row = s.g.rows[r];
      const code = infoInteger(row[s.cc]);
      const y = infoNumber(row[ytdCol]);
      const cf = infoNumber(row[cfCol]);
      if (code.ok && code.v === null) {
        const hasAmt = (y.ok && y.v) || (cf.ok && cf.v) || !y.ok || !cf.ok;
        if (hasAmt) errs.push(`Last month's statement row ${r}: there is an amount but no cost centre code.`);
        continue;
      }
      if (!code.ok) { errs.push(`Last month's statement row ${r}: Cost Center ${code.why}.`); continue; }
      if (!y.ok) errs.push(`Last month's statement row ${r} (cost centre ${code.v}): YTD ${fyLabelHeader(period.cur)} — ${y.why}.`);
      if (!cf.ok) errs.push(`Last month's statement row ${r} (cost centre ${code.v}): ${fyLabelHeader(period.prev)} Pro CF — ${cf.why}.`);
      if (centres.has(code.v)) { errs.push(`Last month's statement: cost centre ${code.v} appears more than once (rows ${centres.get(code.v).row} and ${r}).`); continue; }
      const name = s.nameCol > 0 ? infoText(row[s.nameCol]) : '';
      centres.set(code.v, { code: code.v, name, ytd: y.ok ? (y.v || 0) : 0, cf: cf.ok ? (cf.v || 0) : 0, row: r, blankYtd: y.ok && y.v === null, blankCf: cf.ok && cf.v === null });
      order.push(code.v);
    }
    if (errs.length) throw new InputError(errs.slice(0, 50));
    if (!centres.size) throw new InputError(['Last month\'s statement: no cost-centre rows found under the headings.']);
    return {
      sheet: s.g.name, headerRow: s.hdr, ytdColLetter: colLetter(ytdCol), cfColLetter: colLetter(cfCol),
      title: stmtTitle, monthConfirmedByUser: !stmtMonth, centres, order,
    };
  }

  /* ------------------------------------------------------------------ *
   * File 3 — Rajeev's provision file (Steps 11–12)
   * ------------------------------------------------------------------ */
  function readRajeev(wb, period, opts) {
    let sheets = wb.worksheets.filter((ws) => /security/i.test(ws.name));
    if (sheets.length === 0 && wb.worksheets.length === 1) sheets = wb.worksheets.slice();
    if (sheets.length !== 1) {
      if (!opts.rajeevSheet) return { needChoice: { kind: 'rajeevSheet', options: wb.worksheets.map((w) => w.name) } };
      sheets = wb.worksheets.filter((w) => w.name === opts.rajeevSheet);
      if (!sheets.length) throw new InputError([`Provision file: sheet "${opts.rajeevSheet}" not found.`]);
    }
    const ws = sheets[0];
    // The sheet name, if it carries a month, must be the ERP month
    const nm = ws.name.match(new RegExp('(?<![a-z])' + MON + '(?![a-z])[\\s\\-\'’.]*(\\d{2}|\\d{4})(?!\\d)', 'i'));
    if (nm) {
      const p = parseMonthYear(nm[1] + '-' + nm[2]);
      if (p && (p.y !== period.y || p.m !== period.m)) throw new InputError([`Provision file: the sheet used is "${ws.name}", but the ERP extract is for ${period.label}. Upload this month's provision file.`]);
    }
    const g = sheetRows(ws);
    let codeCol = -1, codeRow = -1, amtCol = -1, amtRow = -1, nameCol = -1, snoCol = -1, catCol = -1, zoneCol = -1, stateCol = -1;
    for (let r = 1; r <= Math.min(20, g.rowCount); r++) for (let c = 1; c <= g.colCount; c++) {
      const k = norm(infoText(g.rows[r][c]));
      if (k === 'storagecode' && codeCol < 0) { codeCol = c; codeRow = r; }
      if (k === 'provisionwithgst' && amtCol < 0) { amtCol = c; amtRow = r; }
      if (k === 'centername' && nameCol < 0) nameCol = c;
      if ((k === 'sno' || k === 'srno') && snoCol < 0) snoCol = c;
      if (k === 'category' && catCol < 0) catCol = c;
      if (k === 'zone' && zoneCol < 0) zoneCol = c;
      if (k === 'state' && stateCol < 0) stateCol = c;
    }
    const errs = [];
    if (codeCol < 0) errs.push(`Provision file (sheet "${ws.name}"): no "Storage Code" heading found.`);
    if (amtCol < 0) errs.push(`Provision file (sheet "${ws.name}"): no "Provision with GST" heading found.`);
    if (errs.length) throw new InputError(errs);
    const first = Math.max(codeRow, amtRow) + 1;
    // Independent total of every number in the "Provision with GST" column below the heading (for the completeness check)
    let columnTotal = 0;
    for (let r = first; r <= g.rowCount; r++) { const n = infoNumber(g.rows[r][amtCol]); if (n.ok && n.v !== null) columnTotal += n.v; }
    const lines = []; const notes = [];
    for (let r = first; r <= g.rowCount; r++) {
      const row = g.rows[r];
      const code = infoInteger(row[codeCol]);
      const amt = infoNumber(row[amtCol]);
      if (code.ok && code.v === null && amt.ok && amt.v === null) continue;
      if (code.ok && code.v === null) { errs.push(`Provision file row ${r}: there is a provision amount but no Storage Code.`); continue; }
      if (!code.ok) { errs.push(`Provision file row ${r}: Storage Code ${code.why}.`); continue; }
      if (!amt.ok) { errs.push(`Provision file row ${r} (Storage Code ${code.v}): Provision with GST — ${amt.why}.`); continue; }
      const pwg = amt.v === null ? 0 : amt.v;
      if (amt.v === null) notes.push({ row: r, code: code.v });
      const snoN = snoCol > 0 ? infoNumber(row[snoCol]) : { ok: false };
      lines.push({ row: r, code: code.v, name: nameCol > 0 ? infoText(row[nameCol]) : '', sno: snoN.ok && snoN.v !== null ? snoN.v : (snoCol > 0 ? infoText(row[snoCol]) : ''), pwg, blank: amt.v === null, category: catCol > 0 ? infoText(row[catCol]) : '', zone: zoneCol > 0 ? infoText(row[zoneCol]) : '', state: stateCol > 0 ? infoText(row[stateCol]) : '' });
    }
    if (errs.length) throw new InputError(errs.slice(0, 50));
    if (!lines.length) throw new InputError([`Provision file (sheet "${ws.name}"): no rows with a Storage Code were found.`]);
    // Steps 11–12, row by row, in the same order of operations as the SOP's Excel formulas
    for (const l of lines) { l.basic = l.pwg / SOP.gstDivisor * SOP.gstBasicFactor; l.final = l.basic * SOP.finalFactor; }
    return { sheet: ws.name, codeColLetter: colLetter(codeCol), amtColLetter: colLetter(amtCol), firstRow: first, lines, blanks: notes, columnTotal };
  }

  /* ------------------------------------------------------------------ *
   * Analyse: read the three files, apply Steps 7–9, propose Step 6 tags
   * ------------------------------------------------------------------ */
  async function load(files) {
    const errors = [];
    const out = {};
    for (const [k, label] of [['extract', 'ERP extract'], ['opening', 'Last month\'s statement'], ['rajeev', 'Provision file']]) {
      if (!files[k]) { errors.push(`${label}: no file selected.`); continue; }
      try { out[k] = await loadWorkbook(files[k], label); } catch (e) { if (e instanceof InputError) errors.push(...e.list); else throw e; }
    }
    if (errors.length) throw new InputError(errors);
    return out;
  }

  function analyse(wbs, opts) {
    opts = opts || {};
    const extract = readExtract(wbs.extract);
    const period = extract.period;
    const opening = readOpening(wbs.opening, period, opts);
    if (opening.needChoice) return { needChoice: opening.needChoice, period };
    const rajeev = readRajeev(wbs.rajeev, period, opts);
    if (rajeev.needChoice) return { needChoice: rajeev.needChoice, period };

    // Steps 7–8: keep P2P bookings and GST reclass entries; everything else is removed
    const kept = [], removed = [];
    for (const r of extract.rows) {
      const src = r.source.trim().toLowerCase();
      const by = r.createdBy.trim().toLowerCase();
      if (SOP.p2pSources.includes(src)) { r.keptAs = 'P2P booking (Step 7)'; kept.push(r); }
      else if (src === 'spreadsheet' && by === SOP.gstReclassCreator) { r.keptAs = 'GST reclass (Step 8)'; kept.push(r); }
      else { r.removedWhy = src === 'spreadsheet' ? 'Spreadsheet entry not created by ' + SOP.gstReclassCreator + ' (Step 8)' : `Source "${r.source || '(blank)'}" is not a P2P booking (Step 7)`; removed.push(r); }
    }
    // Step 9: vendor name
    for (const r of kept) {
      if (r.attr3) { r.vendor = r.attr3; r.vendorFrom = 'Attribute 3'; }
      else if (r.party) { r.vendor = r.party; r.vendorFrom = 'Party Name (Step 9)'; }
      else { r.vendor = ''; r.vendorFrom = 'none'; }
    }
    // Step 6: proposals
    kept.forEach((r, i) => {
      r.id = i;
      const p = proposeFY(r.desc, period);
      r.autoFy = p.sure ? p.fy : null;
      r.suggestFy = p.fy;
      r.basis = p.basis;
      r.needsConfirm = !p.sure;
    });
    return { period, extract, opening, rajeev, kept, removed };
  }

  /* ------------------------------------------------------------------ *
   * Build: Steps 10, 13–16 + Section 04 checks + output workbook
   * ------------------------------------------------------------------ */
  function compute(state, tagChoice) {
    const { period, kept, removed, opening, rajeev, extract } = state;
    const choice = tagChoice || new Map();
    for (const r of kept) {
      const c = choice.has(r.id) ? choice.get(r.id) : null;
      if (c !== null && c !== period.cur && c !== period.prev) throw new Error('Invalid financial year chosen for line ' + r.erpRow);
      if (c !== null) { r.fy = c; r.tagBy = r.needsConfirm ? 'User' : (c === r.autoFy ? 'Tool' : 'User (changed)'); }
      else if (!r.needsConfirm) { r.fy = r.autoFy; r.tagBy = 'Tool'; }
      else { r.fy = period.cur; r.tagBy = 'Default'; } // Finance rule (Oct 2026): year not stated in the narration -> current FY
    }

    // Step 10 — pivot: SEGMENT1 × BALANCING_SEGMENT_DESC × Exp Year, sum of BALANCE (in paise, exact)
    const pivot = new Map(); // key code|desc
    const byCode = new Map(); // code -> {cur, prev} paise
    for (const r of kept) {
      const key = r.seg1 + '\u0000' + r.ccNameRaw.toLowerCase(); // Excel text comparison ignores case
      if (!pivot.has(key)) pivot.set(key, { code: r.seg1, name: r.ccNameRaw, prev: 0, cur: 0 });
      const p = pivot.get(key); const amt = toPaise(r.bal);
      if (r.fy === period.cur) p.cur += amt; else p.prev += amt;
      if (!byCode.has(r.seg1)) byCode.set(r.seg1, { cur: 0, prev: 0, name: r.ccName });
      const b = byCode.get(r.seg1); if (r.fy === period.cur) b.cur += amt; else b.prev += amt;
    }
    // Rajeev by code (Step 16 — "[Month] Pro as per Rajeev"; duplicates are added, one row per centre)
    const rajByCode = new Map(); const rajCount = new Map();
    for (const l of rajeev.lines) { rajByCode.set(l.code, (rajByCode.get(l.code) || 0) + l.final); rajCount.set(l.code, (rajCount.get(l.code) || 0) + 1); }

    // Statement rows: last month's centres in last month's order, then any new centres
    const booked = [...byCode.keys()].sort((a, b) => a - b);
    const codes = booked.concat(opening.order.filter((c) => !byCode.has(c)));
    const newCodes = [...rajByCode.keys()].filter((c) => !opening.centres.has(c) && !byCode.has(c)).sort((a, b) => a - b);
    codes.push(...newCodes);
    const rajName = new Map(); for (const l of rajeev.lines) if (!rajName.has(l.code) && l.name) rajName.set(l.code, l.name);
    const stmt = codes.map((code) => {
      const o = opening.centres.get(code);
      const b = byCode.get(code) || { cur: 0, prev: 0 };
      const proCur = o ? o.ytd : 0, proPrev = o ? o.cf : 0;      // Step 13
      const bkCur = b.cur / 100, bkPrev = b.prev / 100;           // Step 14
      const cfCur = proCur - bkCur, cfPrev = proPrev - bkPrev;    // Step 15
      const raj = rajByCode.get(code) || 0;                       // Step 16
      const name = (o && o.name) || (byCode.get(code) && byCode.get(code).name) || rajName.get(code) || '';
      return { code, name, proCur, proPrev, bkCur, bkPrev, cfCurSop: cfCur, cfPrevSop: cfPrev, adjCur: 0, adjPrev: 0, cfCur, cfPrev, raj, ytd: 0, isNew: !o };
    });

    // Finance rule (Oct 2026, added on finance's request; not in SOP v1.0):
    // in each Pro CF column, every negative carry-forward is set to 0 and the total of those negatives is
    // taken from the SOP.negCfTopN (10) centres with the highest carry-forward in the same column, in proportion
    // to their carry-forward. The column total does not change. Shares are rounded to paise; the largest
    // centre takes the rounding remainder so the adjustments net to exactly zero.
    const adjInfo = {};
    for (const [sopF, adjF, finF, label] of [['cfCurSop', 'adjCur', 'cfCur', fyLabelHeader(period.cur)], ['cfPrevSop', 'adjPrev', 'cfPrev', fyLabelHeader(period.prev)]]) {
      const neg = stmt.filter((s) => s[sopF] < -EPS);
      const info = { label, negatives: neg.map((s) => s.code), recipients: [], total: 0, applied: false, impossible: false };
      adjInfo[adjF] = info;
      if (!neg.length) continue;
      const N = neg.reduce((a, s) => a - s[sopF], 0);
      const top = stmt.filter((s) => s[sopF] > EPS).sort((a, b) => (b[sopF] - a[sopF]) || (a.code - b.code)).slice(0, SOP.negCfTopN);
      const S = top.reduce((a, s) => a + s[sopF], 0);
      info.total = N;
      if (!top.length || N > S + EPS) { info.impossible = true; continue; } // cannot be absorbed — reported as a failed check
      for (const s of neg) s[adjF] = -s[sopF];
      let given = 0;
      for (let i = 1; i < top.length; i++) { const share = Math.round(N * top[i][sopF] / S * 100) / 100; top[i][adjF] = -share; given += share; }
      top[0][adjF] = -(N - given);
      info.recipients = top.map((s) => s.code); info.applied = true;
    }
    for (const s of stmt) { s.cfCur = s.cfCurSop + s.adjCur; s.cfPrev = s.cfPrevSop + s.adjPrev; s.ytd = s.cfCur + s.raj; }

    // Totals
    const sum = (f) => stmt.reduce((a, s) => a + s[f], 0);
    const T = { proCur: sum('proCur'), proPrev: sum('proPrev'), bkCur: sum('bkCur'), bkPrev: sum('bkPrev'), cfCur: sum('cfCur'), cfPrev: sum('cfPrev'), raj: sum('raj'), ytd: sum('ytd'), adjCur: sum('adjCur'), adjPrev: sum('adjPrev') };
    const keptTotalP = kept.reduce((a, r) => a + toPaise(r.bal), 0);
    const removedTotalP = removed.reduce((a, r) => a + toPaise(r.bal), 0);
    const extractTotalP = extract.rows.reduce((a, r) => a + toPaise(r.bal), 0);
    let pivotTotalP = 0; for (const p of pivot.values()) pivotTotalP += p.cur + p.prev;
    const rajInPwg = rajeev.lines.reduce((a, l) => a + l.pwg, 0);
    const rajFinal = rajeev.lines.reduce((a, l) => a + l.final, 0);
    let openYtd = 0, openCf = 0; for (const o of opening.centres.values()) { openYtd += o.ytd; openCf += o.cf; }

    // Section 04 checks (+ completeness checks). status: PASS | FAIL | REVIEW
    const checks = [];
    const add = (id, sop, text, expected, actual, kind) => {
      let status;
      if (kind === 'count0') status = actual === 0 ? 'PASS' : 'REVIEW';
      else if (kind === 'count0fail') status = actual === 0 ? 'PASS' : 'FAIL';
      else status = near(expected, actual) ? 'PASS' : 'FAIL';
      checks.push({ id, sop, text, expected, actual, kind, status });
    };
    const noTag = kept.filter((r) => r.fy !== period.cur && r.fy !== period.prev).length;
    const noVendor = kept.filter((r) => !r.vendor).length;
    const wrongKind = kept.filter((r) => !SOP.p2pSources.includes(r.source.trim().toLowerCase()) && !(r.source.trim().toLowerCase() === 'spreadsheet' && r.createdBy.trim().toLowerCase() === SOP.gstReclassCreator)).length;
    add(1, '04 · check 1', 'Every Working line has an Exp Year tag', 0, noTag, 'count0fail');
    add(2, '04 · check 1', 'Every Working line has a vendor name in Attribute 3 (lines with no Attribute 3 and no Party Name in the ERP cannot be filled)', 0, noVendor, 'count0');
    add(3, '04 · check 2', 'Working contains only P2P bookings and GST reclass entries by ' + SOP.gstReclassCreator, 0, wrongKind, 'count0fail');
    add(4, '04 · check 3', 'Pivot grand total = total BALANCE of the Working sheet', keptTotalP / 100, pivotTotalP / 100);
    add(5, '04 · check 4', `Totals row: ${fyLabelHeader(period.cur)} Pro − ${fyLabelHeader(period.cur)} Booking = ${fyLabelHeader(period.cur)} Pro CF`, T.proCur - T.bkCur, T.cfCur);
    add(6, '04 · check 4', `Totals row: ${fyLabelHeader(period.prev)} Pro − ${fyLabelHeader(period.prev)} Booking = ${fyLabelHeader(period.prev)} Pro CF`, T.proPrev - T.bkPrev, T.cfPrev);
    add(7, '04 · check 5', `Totals row: ${fyLabelHeader(period.cur)} Pro CF + ${period.label} Pro as per Rajeev = YTD ${fyLabelHeader(period.cur)}`, T.cfCur + T.raj, T.ytd);
    add(8, '04 · check 5', 'Rajeev total in the statement = total of the GST conversion (Part D)', rajFinal, T.raj);
    add(9, 'Completeness', 'ERP extract total = Working total + Removed lines total (no line lost)', extractTotalP / 100, (keptTotalP + removedTotalP) / 100);
    add(10, 'Completeness', 'Statement bookings (both years) = Working total (every booked centre is in the statement)', keptTotalP / 100, T.bkCur + T.bkPrev);
    add(11, 'Completeness', 'Provision file: total of the "Provision with GST" column = GST conversion input total', rajeev.columnTotal, rajInPwg);
    add(12, 'Completeness · Step 13', `Opening ${fyLabelHeader(period.cur)} Pro = last month's YTD ${fyLabelHeader(period.cur)} total`, openYtd, T.proCur);
    add(13, 'Completeness · Step 13', `Opening ${fyLabelHeader(period.prev)} Pro = last month's ${fyLabelHeader(period.prev)} Pro CF total`, openCf, T.proPrev);
    add(14, 'Negative CF rule', `${fyLabelHeader(period.cur)} CF adjustments net to zero (column total unchanged)`, 0, T.adjCur);
    add(15, 'Negative CF rule', `${fyLabelHeader(period.prev)} CF adjustments net to zero (column total unchanged)`, 0, T.adjPrev);
    add(16, 'Negative CF rule', 'No negative Pro CF left after the adjustment', 0, stmt.filter((s) => s.cfCur < -EPS || s.cfPrev < -EPS).length, 'count0fail');

    const status = checks.some((c) => c.status === 'FAIL') ? 'NOT READY' : (checks.some((c) => c.status === 'REVIEW') ? 'READY — items to review' : 'READY');

    // Exceptions / information for the reviewer
    const exceptions = [];
    for (const r of kept) if (r.tagBy !== 'Tool') exceptions.push({ type: r.tagBy === 'Default' ? 'Exp Year by default' : 'Exp Year set by user', ref: `ERP row ${r.erpRow}`, code: r.seg1, amount: r.bal, detail: `${fyLabelWorking(r.fy)} — ${r.tagBy}. Tool note: ${r.basis}`, text: r.desc });
    for (const r of kept) if (!r.vendor) exceptions.push({ type: 'No vendor name', ref: `ERP row ${r.erpRow}`, code: r.seg1, amount: r.bal, detail: 'Attribute 3 and Party Name are both blank in the ERP extract', text: r.desc });
    for (const s of stmt) if (s.isNew) exceptions.push({ type: 'New centre (not in last month\'s statement)', ref: '', code: s.code, amount: null, detail: 'Opening provision taken as 0', text: s.name });
    for (const [code, n] of rajCount) if (n > 1) exceptions.push({ type: 'Centre listed more than once in provision file', ref: '', code, amount: rajByCode.get(code), detail: `${n} rows added together (one statement row per centre)`, text: '' });
    for (const b of rajeev.blanks) exceptions.push({ type: 'Blank "Provision with GST"', ref: `Provision file row ${b.row}`, code: b.code, amount: 0, detail: 'Blank cell taken as 0 (as Excel does)', text: '' });
    for (const s of stmt) if (opening.centres.has(s.code) && !rajByCode.has(s.code)) exceptions.push({ type: 'No provision this month', ref: '', code: s.code, amount: null, detail: 'Centre is in last month\'s statement but not in the provision file', text: s.name });
    for (const o of opening.centres.values()) if (o.blankYtd || o.blankCf) exceptions.push({ type: 'Blank opening balance in last month\'s statement', ref: `Row ${o.row}`, code: o.code, amount: 0, detail: `${o.blankYtd ? 'YTD ' + fyLabelHeader(period.cur) : ''}${o.blankYtd && o.blankCf ? ' and ' : ''}${o.blankCf ? fyLabelHeader(period.prev) + ' Pro CF' : ''} blank — taken as 0 (as Excel does)`, text: o.name });
    if (!kept.length) exceptions.push({ type: 'No bookings this month', ref: '', code: null, amount: 0, detail: 'No P2P or GST reclass lines in the ERP extract — booking columns are 0', text: '' });
    for (const s of stmt) {
      if (s.cfCurSop < -EPS) exceptions.push({ type: `Negative ${fyLabelHeader(period.cur)} Pro CF`, ref: '', code: s.code, amount: s.cfCurSop, detail: 'Before the negative carry-forward rule', text: s.name });
      if (s.cfPrevSop < -EPS) exceptions.push({ type: `Negative ${fyLabelHeader(period.prev)} Pro CF`, ref: '', code: s.code, amount: s.cfPrevSop, detail: 'Before the negative carry-forward rule', text: s.name });
    }

    // "Points to check" — unusual items only, in plain words, for the finance reviewer.
    // Information only: nothing in this list changes any figure.
    const points = [];
    const C = fyLabelHeader(period.cur), Pv = fyLabelHeader(period.prev);
    const pt = (code, name, point, amount, check) => points.push({ code, name, point, amount: amount == null ? null : Math.round(amount * 100) / 100, check });
    for (const s of stmt) {
      if (!rajByCode.has(s.code)) {
        const bk = s.bkCur + s.bkPrev;
        pt(s.code, s.name, `Not in Rajeev's provision file, so "${period.display} Pro as per Rajeev" is 0.` + (Math.abs(bk) > EPS ? ` ${fmtINR(bk)} was booked for this centre this month.` : ''), Math.abs(bk) > EPS ? bk : null, 'Check with Rajeev whether this centre should have a provision this month.');
      }
      if (s.isNew) pt(s.code, s.name, `New centre: not in last month's statement, so its opening ${C} Pro and ${Pv} Pro are 0.`, null, 'Confirm this centre is new and has no balance carried from last month.');
    }
    for (const b of rajeev.blanks) pt(b.code, (stmt.find((s) => s.code === b.code) || {}).name || '', `Rajeev's provision file row ${b.row}: "Provision with GST" is blank, taken as 0.`, null, 'Check whether an amount is missing in Rajeev\'s file.');
    for (const o of opening.centres.values()) if (o.blankYtd || o.blankCf) pt(o.code, o.name, `Last month's statement row ${o.row}: ${o.blankYtd ? 'YTD ' + C : ''}${o.blankYtd && o.blankCf ? ' and ' : ''}${o.blankCf ? Pv + ' Pro CF' : ''} is blank, taken as 0.`, null, 'Check whether last month\'s balance is missing.');
    for (const s of stmt) {
      for (const [sopF, adjF, lab] of [['cfCurSop', 'adjCur', C], ['cfPrevSop', 'adjPrev', Pv]]) {
        if (s[sopF] >= -EPS) continue;
        const inf = adjInfo[adjF];
        if (inf.applied) pt(s.code, s.name, `${lab} Pro CF was ${fmtINR(s[sopF])} (more booked than provided). Set to 0 under the negative carry-forward rule; the amount was taken from the ${inf.recipients.length} centres with the highest ${lab} Pro CF (see the adjustment table in the "Points to check" sheet).`, s[sopF], 'Review this centre\'s bookings and provision.');
        else pt(s.code, s.name, `${lab} Pro CF is negative (${fmtINR(s[sopF])}) and could NOT be absorbed: the negative total ${fmtINR(inf.total)} is more than the top ${SOP.negCfTopN} centres can take.`, s[sopF], 'The negative carry-forward rule could not be applied this month. Review with finance.');
      }
    }
    const defaulted = kept.filter((r) => r.tagBy === 'Default');
    if (defaulted.length) pt(null, '', `${defaulted.length} Working line(s) had no year in the description and were counted as ${fyLabelWorking(period.cur)} by default (total ${fmtINR(defaulted.reduce((a, r) => a + r.bal, 0))}). They are marked in the Remarks column of the Working sheet.`, defaulted.reduce((a, r) => a + r.bal, 0), 'For information. Check these lines if any of them belong to last year.');
    if (noVendor) pt(null, '', `${noVendor} Working line(s) have no vendor name (Attribute 3 and Party Name are both blank in the ERP extract).`, null, 'For information. These lines are still counted.');
    if (!kept.length) pt(null, '', 'No P2P or GST reclass lines in the ERP extract, so all booking columns are 0.', 0, 'Confirm the ERP extract is complete for this month.');

    return { period, stmt, T, pivot: [...pivot.values()], checks, status, exceptions, points, adjInfo, totals: { keptTotalP, removedTotalP, extractTotalP, rajInPwg, rajFinal, openYtd, openCf } };
  }

  /* ------------------------------------------------------------------ *
   * Output workbook
   * ------------------------------------------------------------------ */
  const FMT_AMT = '#,##0;-#,##0;"-"';
  const FMT_AMT2 = '#,##0.00;-#,##0.00;"-"';
  const HDR_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF1F4' } };
  const THIN = { style: 'thin', color: { argb: 'FFC9CED6' } };
  const qs = (name) => "'" + name.replace(/'/g, "''") + "'";

  function styleHeader(row) {
    row.eachCell((c) => { c.font = { bold: true }; c.fill = HDR_FILL; c.border = { bottom: THIN }; c.alignment = { vertical: 'middle', wrapText: true }; });
  }

  // One header style across all sheets (the header colour finance uses on its Summary), thin borders on every table
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

  async function buildWorkbook(state, R, meta) {
    // Output mirrors finance's own monthly workbook (as in the SOP's figures): Summary, Pivot, Provision_Rajeev, Working, Raw.
    const { period, kept, extract, rajeev } = state;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Security Provision Tool ' + TOOL_VERSION;
    wb.created = meta && meta.now ? meta.now : new Date();
    wb.calcProperties = { fullCalcOnLoad: true };
    const L = period.display;
    const wsSum = wb.addWorksheet('Summary');
    const wsPiv = wb.addWorksheet('Pivot');
    const wsRaj = wb.addWorksheet('Provision_Rajeev');
    const wsWork = wb.addWorksheet('Working');
    const wsRaw = wb.addWorksheet('Raw');

    /* ---- Raw: the ERP extract exactly as downloaded ---- */
    {
      const src = state.extract.rawSheet;
      src.eachRow({ includeEmpty: true }, (row, r) => {
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          const i = cellInfo(cell);
          wsRaw.getCell(r, c).value = i.k === 'blank' ? null : origValue(i);
        });
      });
      src.columns && src.columns.forEach((col, i) => { if (col && col.width) wsRaw.getColumn(i + 1).width = col.width; });
      const hdrRow = extract.headerRow, lastCol = Math.max(...extract.headers.map((h) => h.c));
      formatTable(wsRaw, [hdrRow], 1, lastCol, hdrRow + 1, src.rowCount);
      wsRaw.getCell('A1').font = { bold: true, size: 12 };
      // report parameter block above the table (as the ERP prints it)
      for (let r = 1; r < hdrRow; r++) { const row = src.getRow(r); let any = false; row.eachCell(() => { any = true; }); if (!any) continue;
        const names = []; row.eachCell((c, ci) => { names.push(ci); }); if (norm(cellInfo(src.getCell(r, 1)).v) === 'fromdate') formatTable(wsRaw, [r], 1, Math.max(...names), r + 1, r + 1); }
      wsRaw.views = [{ state: 'frozen', ySplit: hdrRow }];
    }

    /* ---- Working (Steps 5–9) ---- */
    const cols = [];
    for (const h of extract.headers) {
      cols.push({ kind: 'orig', h });
      if (h.k === 'gldate') cols.push({ kind: 'glmonth' });
      if (h.k === 'linedescription') cols.push({ kind: 'exp' });
    }
    cols.push({ kind: 'remarks' });
    const wcol = {}; cols.forEach((c, i) => { wcol[c.kind === 'orig' ? c.h.k : c.kind] = i + 1; });
    wsWork.addRow(cols.map((c) => c.kind === 'orig' ? c.h.name : c.kind === 'exp' ? 'Exp Year' : c.kind === 'glmonth' ? 'GL Month' : 'Remarks'));

    const partyLetter = colLetter(wcol.partyname);
    kept.forEach((r, i) => {
      const rowNo = i + 2;
      wsWork.addRow(cols.map((c) => {
        if (c.kind === 'exp') return fyLabelWorking(r.fy);
        if (c.kind === 'remarks') return r.tagBy === 'Default' ? `Actual Booking - ${fyLabelWorking(period.cur)} by default (year not in description)` : 'Actual Booking';
        if (c.kind === 'glmonth') { const g = r.glDate; return g ? new Date(Date.UTC(g.y, g.m - 1, 1)) : null; }
        const k = c.h.k;
        if (k === 'balance') return r.bal;
        if (k === 'attribute3' && !r.attr3 && r.party) return { formula: partyLetter + rowNo, result: r.party }; // Step 9
        return origValue(r.cells[c.h.c]);
      }));
    });
    if (wcol.glmonth) wsWork.getColumn(wcol.glmonth).numFmt = 'mmm-yy';
    formatTable(wsWork, [1], 1, cols.length, 2, kept.length + 1);
    wsWork.views = [{ state: 'frozen', ySplit: 1 }];

    /* ---- Pivot (Step 10) — same layout as finance's pivot ---- */
    wsPiv.getCell('A3').value = 'Sum of BALANCE'; wsPiv.getCell('C3').value = 'Exp Year';
    wsPiv.getRow(4).values = ['SEGMENT1', 'BALANCING_SEGMENT_DESC', fyLabelWorking(period.prev), fyLabelWorking(period.cur), 'Grand Total'];

    const pivRows = R.pivot.slice().sort((a, b) => a.code - b.code || (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
    let pPrev = 0, pCur = 0;
    pivRows.forEach((p, i) => {
      const rn = 5 + i; pPrev += p.prev; pCur += p.cur;
      wsPiv.getRow(rn).values = [p.code, p.name || null, p.prev ? p.prev / 100 : null, p.cur ? p.cur / 100 : null, (p.prev + p.cur) / 100];
    });
    const pGT = 5 + pivRows.length;
    wsPiv.getRow(pGT).values = ['Grand Total', null, pPrev / 100, pCur / 100, (pPrev + pCur) / 100];
    formatTable(wsPiv, [4], 1, 5, 5, pGT);
    for (let c = 1; c <= 5; c++) { const cell = wsPiv.getRow(pGT).getCell(c); cell.font = { bold: true }; cell.fill = HEAD_FILL; }
    wsPiv.getCell('A3').font = { bold: true }; wsPiv.getCell('C3').font = { bold: true };
    ['C', 'D', 'E'].forEach((c) => { wsPiv.getColumn(c).numFmt = '#,##0.00'; wsPiv.getColumn(c).width = 14; });
    wsPiv.getColumn('B').width = 28; wsPiv.getColumn('A').width = 12;
    const pEnd = Math.max(5, pGT - 1);

    /* ---- Provision_Rajeev (Steps 11–12) — layout as SOP Figure 2 ---- */
    wsRaj.getRow(1).values = ['S.No.', 'Center Name ', 'Storage Code ', 'Category ', 'Zone ', 'State '];

    const rEnd = 3 + rajeev.lines.length;
    wsRaj.getCell('G2').value = { formula: `SUM(G4:G${rEnd})`, result: R.totals.rajInPwg };
    wsRaj.getRow(3).values = [null, null, null, null, null, null, 'Provision  with GST ', 'Basic', 'With 113.5%GST'];

    const nameOf = new Map(R.stmt.map((s) => [s.code, s.name]));
    rajeev.lines.forEach((l, i) => {
      const rn = 4 + i;
      wsRaj.getRow(rn).values = [l.sno === '' ? null : l.sno, l.name || null, l.code, l.category || null, l.zone || null, l.state || null, l.pwg,
        { formula: `G${rn}/${SOP.gstDivisor}*${SOP.gstBasicFactor}`, result: l.basic },
        { formula: `H${rn}*113.5%`, result: l.final },
        { formula: `_xlfn.XLOOKUP(C${rn},Summary!B:B,Summary!C:C)`, result: nameOf.get(l.code) || '' }];
    });
    ['G', 'H', 'I'].forEach((c) => { wsRaj.getColumn(c).numFmt = '#,##0.00'; wsRaj.getColumn(c).width = 14; });
    formatTable(wsRaj, [1], 1, 6, 4, rEnd);            // S.No. … State
    formatTable(wsRaj, [3], 7, 9, 4, rEnd);            // Provision with GST, Basic, With 113.5%GST
    for (let r = 4; r <= rEnd; r++) wsRaj.getRow(r).getCell(10).border = BOX;
    for (const c of ['G2']) { const cell = wsRaj.getCell(c); cell.font = { bold: true }; cell.border = BOX; }
    wsRaj.views = [{ state: 'frozen', ySplit: 3 }];
    wsRaj.getColumn('B').width = 26; wsRaj.getColumn('C').width = 10; wsRaj.getColumn('F').width = 16; wsRaj.getColumn('J').width = 22;

    /* ---- Summary (Steps 13–16) — same layout as finance's statement ---- */
    wsSum.getCell('A2').value = `Security Provision for the month of ${L}`;
    wsSum.getCell('A2').font = { bold: true, underline: true };
    wsSum.getCell('L2').value = { formula: '(I3+K3)/10^6', result: (R.T.cfPrev + R.T.ytd) / 1e6 };
    wsSum.getCell('L2').numFmt = '0.00';
    const hdr = [null, 'Cost Center', 'Cost Center Name', `${fyLabelHeader(period.cur)} Pro`, `${fyLabelHeader(period.prev)} Pro`, `${fyLabelHeader(period.cur)} Booking`, `${fyLabelHeader(period.prev)} Booking`, `${fyLabelHeader(period.cur)} Pro CF`, `${fyLabelHeader(period.prev)} Pro CF`, `${L} Pro as per Rajeev`, `YTD ${fyLabelHeader(period.cur)}`];
    wsSum.getRow(4).values = hdr;

    const first = 5, last = 4 + R.stmt.length;
    const PIV = (c) => `Pivot!$${c}$5:$${c}$${pEnd}`;
    R.stmt.forEach((s, i) => {
      const rn = first + i;
      wsSum.getRow(rn).values = [
        null, s.code, s.name, s.proCur, s.proPrev,
        { formula: `SUMIFS(${PIV('D')},${PIV('A')},$B${rn})`, result: s.bkCur },
        { formula: `SUMIFS(${PIV('C')},${PIV('A')},$B${rn})`, result: s.bkPrev },
        // Centres changed by the negative carry-forward rule hold the final figure (as finance enters it);
        // the before/after detail is in the adjustment table on the "Points to check" sheet.
        Math.abs(s.adjCur) > 0 ? s.cfCur : { formula: `D${rn}-F${rn}`, result: s.cfCur },
        Math.abs(s.adjPrev) > 0 ? s.cfPrev : { formula: `E${rn}-G${rn}`, result: s.cfPrev },
        { formula: `SUMIFS(Provision_Rajeev!$I$4:$I$${rEnd},Provision_Rajeev!$C$4:$C$${rEnd},$B${rn})`, result: s.raj },
        { formula: `H${rn}+J${rn}`, result: s.ytd },
      ];

    });
    const tot = { D: R.T.proCur, E: R.T.proPrev, F: R.T.bkCur, G: R.T.bkPrev, H: R.T.cfCur, I: R.T.cfPrev, J: R.T.raj, K: R.T.ytd };
    for (const [c, v] of Object.entries(tot)) { const cell = wsSum.getCell(`${c}3`); cell.value = { formula: `SUBTOTAL(9,${c}${first}:${c}${Math.max(first, last)})`, result: v }; cell.border = { top: THIN, bottom: THIN, left: THIN, right: THIN }; }
    for (const c of ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) { wsSum.getColumn(c).numFmt = FMT_AMT; wsSum.getColumn(c).width = 14; }
    formatTable(wsSum, [4], 2, 11, 5, last);
    for (const c of ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) wsSum.getCell(c + '3').font = { bold: true };
    wsSum.getRow(4).height = 30;
    wsSum.views = [{ state: 'frozen', ySplit: 4 }];
    wsSum.getColumn('A').width = 3; wsSum.getColumn('B').width = 10; wsSum.getColumn('C').width = 24;
    wsSum.autoFilter = { from: { row: 4, column: 2 }, to: { row: 4, column: 11 } };

    /* ---- Points to check (information only; no figure depends on this sheet) ---- */
    const wsPts = wb.addWorksheet('Points to check');
    wsPts.getCell('A1').value = `Points to check — ${L}`;
    wsPts.getCell('A1').font = { bold: true };
    wsPts.getCell('A2').value = 'Unusual items found while preparing the statement, for information. Negative carry-forward adjustments are listed in the table further below.';
    wsPts.getCell('A2').font = { italic: true };
    wsPts.getRow(4).values = ['#', 'Cost Center', 'Cost Center Name', 'Point', 'Amount', 'Suggested check'];
    if (R.points.length) {
      R.points.forEach((p, i) => { wsPts.getRow(5 + i).values = [i + 1, p.code, p.name, p.point, p.amount, p.check]; });
    } else {
      wsPts.getRow(5).values = [null, null, null, 'No points to check this month.'];
    }
    const pLast = 4 + Math.max(1, R.points.length);
    formatTable(wsPts, [4], 1, 6, 5, pLast);
    wsPts.getColumn('A').width = 5; wsPts.getColumn('B').width = 11; wsPts.getColumn('C').width = 24;
    wsPts.getColumn('D').width = 70; wsPts.getColumn('E').width = 15; wsPts.getColumn('F').width = 55;
    wsPts.getColumn('E').numFmt = FMT_AMT2;
    for (let r = 5; r <= pLast; r++) for (const c of ['D', 'F']) wsPts.getCell(c + r).alignment = { wrapText: true, vertical: 'top' };
    wsPts.views = [{ state: 'frozen', ySplit: 4 }];

    // Negative carry-forward adjustment table (finance rule): before / adjustment / after for every centre changed
    const tTop = pLast + 3;
    wsPts.getCell(`A${tTop}`).value = `Negative carry-forward adjustment — negatives set to 0 and taken from the ${SOP.negCfTopN} centres with the highest carry-forward in the same column, in proportion`;
    wsPts.getCell(`A${tTop}`).font = { bold: true };
    const tHdr = tTop + 1;
    wsPts.getRow(tHdr).values = ['#', 'Cost Center', 'Cost Center Name', 'What happened', 'Pro CF before', 'Adjustment', 'Pro CF after'];
    let tr = tHdr + 1, n = 0;
    const amtCells = [];
    for (const [sopF, adjF, finF] of [['cfCurSop', 'adjCur', 'cfCur'], ['cfPrevSop', 'adjPrev', 'cfPrev']]) {
      const inf = R.adjInfo[adjF];
      const rows = R.stmt.filter((x) => Math.abs(x[adjF]) > 0).sort((a, b) => (a[sopF] < 0 ? 0 : 1) - (b[sopF] < 0 ? 0 : 1) || b[sopF] - a[sopF]);
      if (!rows.length) continue;
      const first = tr;
      for (const x of rows) {
        n++;
        const what = x[sopF] < 0 ? `${inf.label} Pro CF negative — set to 0` : `${inf.label} Pro CF — share of the negatives taken from this centre`;
        wsPts.getRow(tr).values = [n, x.code, x.name, what, x[sopF], x[adjF], { formula: `E${tr}+F${tr}`, result: x[finF] }];
        amtCells.push(tr); tr++;
      }
      wsPts.getRow(tr).values = [null, null, null, `Total ${inf.label} adjustment (must be 0)`, null, { formula: `SUM(F${first}:F${tr - 1})`, result: rows.reduce((a, x) => a + x[adjF], 0) }, null];
      wsPts.getRow(tr).font = { bold: true }; amtCells.push(tr); tr++;
    }
    if (n === 0) { wsPts.getRow(tr).values = [null, null, null, 'No negative carry-forward this month — nothing adjusted.']; tr++; }
    formatTable(wsPts, [tHdr], 1, 7, tHdr + 1, tr - 1);
    for (const r of amtCells) for (const c of ['E', 'F', 'G']) wsPts.getCell(c + r).numFmt = FMT_AMT2;
    for (let r = tHdr + 1; r < tr; r++) wsPts.getCell('D' + r).alignment = { wrapText: true, vertical: 'top' };
    wsPts.getColumn('G').width = 15;
    if (R.points.length) wsPts.properties.tabColor = { argb: 'FFF79009' }; // amber sheet tab when there is something to check

    return wb.xlsx.writeBuffer();
  }

  function origValue(info) {
    if (!info) return null;
    switch (info.k) {
      case 'blank': return null;
      case 'num': case 'bool': case 'date': return info.v;
      case 'text': return info.v;
      case 'err': return { error: info.v };
      default: return null;
    }
  }

  function fmtINR(x) {
    const neg = x < 0; x = Math.abs(x);
    const [i, d] = x.toFixed(2).split('.');
    let s = i.length > 3 ? i.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + i.slice(-3) : i;
    return (neg ? '−' : '') + '₹' + s + '.' + d;
  }

  function rulesText(period) {
    const c = fyLabelWorking(period.cur), p = fyLabelWorking(period.prev);
    return [
      ['Steps 1–4', `Input is the raw IIVF Account Analysis download. The tool stops if the file holds more than one period, or accounts outside ${SOP.accountFrom}–${SOP.accountTo}, or (when the report parameters are in the file) a date range other than the full month.`],
      ['Step 5', 'Working sheet = the extract with an "Exp Year" column inserted immediately to the right of LINE_DESCRIPTION. Three audit columns are added at the far right (tool note, who set the tag, ERP row).'],
      ['Step 6', `Each line is tagged ${p} or ${c} from the service month in the narration. Tagged automatically only when the narration states it: a service period (dd-mm-yyyy to dd-mm-yyyy), a month with a financial year ("June 2026-27"), or a month with a year ("June 2026", "Jun-26"). Month-year text inside reference codes (e.g. "/MAY-26/") is ignored. Every other line — no date, a month without a year, only a bill date, a service month later than the posting month, months in two financial years, or a year outside ${p}/${c} — is shown to the user with a suggestion (where one exists) and must be confirmed by the user before the statement is produced. The Working sheet records which lines the user set.`],
      ['Step 7', `Kept: lines whose JE_SOURCE_NAME is a P2P invoice booking — "IIVF P2P" or "Payables".`],
      ['Step 8', `Kept: Spreadsheet journal lines created by ${SOP.gstReclassCreator} (GST reclass). Every other line is removed. Removed lines are listed in the "Removed lines" sheet for audit; they are not used anywhere.`],
      ['Step 9', 'Where ATTRIBUTE3 is blank, it is filled with a cell reference to PARTY_NAME on the same row. If both are blank the line is listed for review (the ERP has no vendor name to use).'],
      ['Step 10', 'Pivot: rows SEGMENT1 + BALANCING_SEGMENT_DESC, columns Exp Year, values Sum of BALANCE. Negative amounts keep their sign.'],
      ['Steps 11–12', `Every row of the provision file: Basic = Provision with GST ÷ ${SOP.gstDivisor} × ${SOP.gstBasicFactor}; Final = Basic × ${SOP.finalFactor}. A blank "Provision with GST" cell counts as 0 (as Excel does) and is listed in Exceptions. A centre that appears on more than one row is added together, because the statement has one row per centre.`],
      ['Step 13', `FY Pro opening = last month's "YTD ${fyLabelHeader(period.cur)}"; ${fyLabelHeader(period.prev)} Pro = last month's "${fyLabelHeader(period.prev)} Pro CF", matched on cost-centre code. Statement must be for the month before the ERP period. If last month's statement has more than one "YTD" column, the user chooses which is the closing balance. A centre not in last month's statement starts at 0 and is listed in Exceptions.`],
      ['Step 14', 'Bookings = SUMIFS over the Working sheet by cost-centre code and Exp Year (equal to the pivot, matched on code).'],
      ['Step 15', `${fyLabelHeader(period.cur)} Pro CF = ${fyLabelHeader(period.cur)} Pro − ${fyLabelHeader(period.cur)} Booking; ${fyLabelHeader(period.prev)} Pro CF = ${fyLabelHeader(period.prev)} Pro − ${fyLabelHeader(period.prev)} Booking. Applied as written; no adjustment of any result.`],
      ['Step 16', `"${period.label} Pro as per Rajeev" = final provision from the GST Conversion sheet (SUMIFS by code). YTD ${fyLabelHeader(period.cur)} = ${fyLabelHeader(period.cur)} Pro CF + ${period.label} Pro as per Rajeev.`],
      ['Statement layout', 'Title, totals row above the header, one row per cost centre: last month\'s centres in last month\'s order, then new centres by code.'],
      ['Section 04', 'All five checks are on the "Checks" sheet with live formulas, plus completeness checks (no line lost, every booked centre included, provision totals carried, opening totals carried).'],
      ['Not in the SOP', 'The tool makes no other adjustment: no manual additions, no overrides of carry-forwards. Any such adjustment is a decision for finance outside this tool.'],
    ];
  }

  async function build(state, tagChoice, meta) {
    const R = compute(state, tagChoice);
    const buffer = await buildWorkbook(state, R, meta);
    return { buffer, result: R };
  }

  return { SOP, TOOL_VERSION, InputError, load, analyse, compute, build, proposeFY, fyOf, fyLabelWorking, fyLabelHeader, fmtINR, _internals: { cellInfo, infoNumber, parseMonthYear, readExtract, readRajeev } };
});
