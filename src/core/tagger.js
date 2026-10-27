// Step 6: proposes each line's financial year ("Exp Year") from its Line Description narration.
// Copied line for line from the approved Security engine (src/security/engine.js, monthTokens + proposeFY);
// test/core/tagger_parity.js runs the Security tagger cases against both copies.
'use strict';
const { MON, MONTH_ABBR, monthFromWord, fyOf, fyLabelWorking, ym, pad2, validDate } = require('./common');

const RE_MONWORD = new RegExp('(?<![a-z])' + MON + '(?![a-z])', 'gi');
const RE_PERIOD = /period\D{0,20}?(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})\s*(?:to|till|upto|up to|-|–)\s*(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/i;
// Text immediately before a month that marks it as the SERVICE month: "for the month of", "month of", "for", "service period of" …
const RE_ANCHOR_BEFORE = /(?:month\s*(?:of)?|\bfor(?:\s+the)?(?:\s+month)?(?:\s+of)?|service\s+(?:month|period)?\s*(?:of|for)?)\s*[:\-]?\s*$/i;
// Text immediately before a month that marks it as a BILL / document date: "dt 10-", "dated", "invoice date" …
const RE_DATE_BEFORE = /(?:\bdt|\bdtd|\bdated|\bdate|\bdatd)\.?\s*[:\-]?\s*(?:\d{1,2}(?:st|nd|rd|th)?[\s\-\/.]*)?$/i;
const RE_MON_MONTH = new RegExp('(?<![a-z])' + MON + '\\s*[\\-\']?\\s*month(?![a-z])', 'i'); // "MARCH MONTH INVOICE"
const RE_BILLDATE = /(?:\bdt|\bdtd|\bdated|\bdate|\bdatd)\.?\s*[:\-]?\s*(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/i;
const RE_FYMARK = /(?<![a-z])fy\s*[-:'’]?\s*(20\d{2}|\d{2})(?:\s*[-–\/]\s*(\d{2}))?(?!\d)/i;

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

module.exports = { proposeFY, monthTokens };
