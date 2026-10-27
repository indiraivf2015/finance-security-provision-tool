// Shared helpers: input errors, text normalising, months and Indian financial years, money formatting.
// Copied from the approved Security engine (src/security/engine.js), which stays untouched.
'use strict';

const EPS = 0.005; // half a paisa: tolerance for comparing floating-point rupee totals
const MAX_FILE_BYTES = 50 * 1024 * 1024;

class InputError extends Error {
  constructor(list) { super(list.join('\n')); this.list = list; this.name = 'InputError'; }
}

const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
const pad2 = (n) => String(n).padStart(2, '0');
const MONTH_INDEX = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const MONTH_ABBR = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MON = '(january|february|march|april|may|june|july|august|september|october|november|december|sept|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)';

function monthFromWord(w) { return MONTH_INDEX[String(w).toLowerCase().slice(0, 3)] || null; }
/** Indian financial year (Apr–Mar), as the 2-digit year in which it ends: Apr-2026..Mar-2027 → 27 */
function fyOf(y, m) { return ((m >= 4 ? y + 1 : y) % 100); }
const fyLabelWorking = (fy) => 'FY ' + pad2(fy);   // SOP Step 6: "FY 26" / "FY 27"
const fyLabelHeader = (fy) => 'FY' + pad2(fy);     // Statement headers: "FY27 Pro" …
const ym = (y, m) => y * 12 + (m - 1);
function prevMonth(y, m) { return m === 1 ? { y: y - 1, m: 12 } : { y, m: m - 1 }; }
/** "Sept-26" style label used in titles and headers (finance writes September as "Sept") */
const monthLabel = (y, m) => (m === 9 ? 'Sept' : MONTH_ABBR[m]) + '-' + pad2(y % 100);
function colLetter(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
const toPaise = (x) => Math.round(x * 100);
const near = (a, b) => Math.abs(a - b) < EPS;
const hasMoreThan2Decimals = (x) => Math.abs(x * 100 - Math.round(x * 100)) > 1e-6;

/** "Sep-26", "Sept-26", "September 2026" → {y, m} */
function parseMonthYear(s) {
  const m = String(s).trim().match(/^([A-Za-z]{3,9})[\s\-'’.]*(\d{2}|\d{4})$/);
  if (!m) return null;
  const mo = monthFromWord(m[1]);
  if (!mo) return null;
  let y = parseInt(m[2], 10); if (y < 100) y += 2000;
  return { y, m: mo };
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

function validDate(d, m, y) {
  if (m < 1 || m > 12 || d < 1) return false;
  return d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** ₹ in Indian grouping: 12,34,567.89. Rounds like Excel's display (15 significant digits first, so
 *  x.785 shows as .79 as in Excel, not .78 from a binary value just below .785). */
function fmtINR(x, decimals) {
  const dp = decimals === undefined ? 2 : decimals;
  const neg = x < 0; x = Math.abs(Number(x.toPrecision(15)));
  const [i, d] = x.toFixed(dp).split('.');
  const s = i.length > 3 ? i.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + i.slice(-3) : i;
  return (neg ? '−' : '') + '₹' + s + (dp ? '.' + d : '');
}

module.exports = {
  EPS, MAX_FILE_BYTES, InputError, norm, pad2, MONTH_INDEX, MONTH_ABBR, MON, monthFromWord, fyOf, fyLabelWorking, fyLabelHeader,
  ym, prevMonth, monthLabel, colLetter, toPaise, near, hasMoreThan2Decimals, parseMonthYear, parseDateText, validDate, fmtINR,
};
