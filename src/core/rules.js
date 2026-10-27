// Finance rules (Oct 2026) shared by provisions, and the checks list. Copied from the approved Security engine.
'use strict';
const { EPS, near } = require('./common');

/**
 * Negative carry-forward rule (finance, Oct 2026; not in the SOPs): in each carry-forward column separately, every
 * negative is set to 0 and the total of those negatives is taken from the topN centres with the highest carry-forward
 * in the same column, in proportion to their carry-forward. The column total does not change. Shares are rounded to
 * paise; the largest centre takes the rounding remainder so the adjustments net to exactly zero.
 * columns: [{ sopF, adjF, label }] — sopF = field before the rule, adjF = field that receives the adjustment.
 * Returns adjInfo keyed by adjF: { label, negatives, recipients, total, applied, impossible }.
 */
function applyNegativeCarryForward(stmt, columns, topN) {
  const adjInfo = {};
  for (const { sopF, adjF, label } of columns) {
    for (const s of stmt) s[adjF] = s[adjF] || 0;
    const neg = stmt.filter((s) => s[sopF] < -EPS);
    const info = { label, negatives: neg.map((s) => s.code), recipients: [], total: 0, applied: false, impossible: false };
    adjInfo[adjF] = info;
    if (!neg.length) continue;
    const N = neg.reduce((a, s) => a - s[sopF], 0);
    const top = stmt.filter((s) => s[sopF] > EPS).sort((a, b) => (b[sopF] - a[sopF]) || (a.code - b.code)).slice(0, topN);
    const S = top.reduce((a, s) => a + s[sopF], 0);
    info.total = N;
    if (!top.length || N > S + EPS) { info.impossible = true; continue; } // cannot be absorbed: reported as a failed check
    for (const s of neg) s[adjF] = -s[sopF];
    let given = 0;
    for (let i = 1; i < top.length; i++) { const share = Math.round(N * top[i][sopF] / S * 100) / 100; top[i][adjF] = -share; given += share; }
    top[0][adjF] = -(N - given);
    info.recipients = top.map((s) => s.code); info.applied = true;
  }
  return adjInfo;
}

/** Checks list. status: PASS | FAIL | REVIEW. Overall: NOT READY if any FAIL. */
function checkList() {
  const checks = [];
  const add = (sop, text, expected, actual, kind) => {
    let status;
    if (kind === 'count0') status = actual === 0 ? 'PASS' : 'REVIEW';
    else if (kind === 'count0fail') status = actual === 0 ? 'PASS' : 'FAIL';
    else status = near(expected, actual) ? 'PASS' : 'FAIL';
    checks.push({ id: checks.length + 1, sop, text, expected, actual, kind, status });
  };
  const status = () => (checks.some((c) => c.status === 'FAIL') ? 'NOT READY' : (checks.some((c) => c.status === 'REVIEW') ? 'READY — items to review' : 'READY'));
  return { checks, add, status };
}

module.exports = { applyNegativeCarryForward, checkList };
