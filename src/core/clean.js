// Steps 6-9 on the extract lines: keep P2P bookings and GST reclass entries, fill the vendor name, tag the year.
// Copied from the approved Security engine (analyse + the tag part of compute).
'use strict';
const { proposeFY } = require('./tagger');

// SOP Step 7 says "keep every row whose source is Payables". In the ERP extract the P2P invoice bookings appear as
// "IIVF P2P" or "Payables", and finance's own working keeps both (owner's decision, Oct 2026). Step 8: GST reclass.
// Housekeeping change (owner's instruction, Oct 2026; Security unchanged): a line created by the GST reclass user is kept
// whatever its JE_SOURCE_NAME says ("Spreadsheet" in the ERP download, but e.g. "GST Reclass" once relabelled).
const P2P_SOURCES = ['iivf p2p', 'payables'];
const GST_RECLASS_CREATOR = 'taxverification2@indiraivf.in';

/** Steps 7-9 + Step 6 proposals. Returns { kept, removed }; extract rows are annotated in place. */
function cleanLines(rows, period) {
  const kept = [], removed = [];
  for (const r of rows) {
    const src = r.source.trim().toLowerCase();
    const by = r.createdBy.trim().toLowerCase();
    if (P2P_SOURCES.includes(src)) { r.keptAs = 'P2P booking (Step 7)'; kept.push(r); }
    else if (by === GST_RECLASS_CREATOR) { r.keptAs = 'GST reclass (Step 8)'; kept.push(r); }
    else { r.removedWhy = src === 'spreadsheet' ? 'Spreadsheet entry not created by ' + GST_RECLASS_CREATOR + ' (Step 8)' : `Source "${r.source || '(blank)'}" is not a P2P booking (Step 7)`; removed.push(r); }
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
    r.yearWritten = !!p.yearWritten;
  });
  return { kept, removed };
}

/** Final Exp Year per line. Finance rule (Oct 2026): year not stated in the narration -> current FY. */
function applyTags(kept, period, tagChoice) {
  const choice = tagChoice || new Map();
  for (const r of kept) {
    const c = choice.has(r.id) ? choice.get(r.id) : null;
    if (c !== null && c !== period.cur && c !== period.prev) throw new Error('Invalid financial year chosen for line ' + r.erpRow);
    if (c !== null) { r.fy = c; r.tagBy = r.needsConfirm ? 'User' : (c === r.autoFy ? 'Tool' : 'User (changed)'); }
    else if (!r.needsConfirm) { r.fy = r.autoFy; r.tagBy = 'Tool'; }
    else { r.fy = period.cur; r.tagBy = 'Default'; }
  }
}

const isWrongKind = (r) => !P2P_SOURCES.includes(r.source.trim().toLowerCase()) && r.createdBy.trim().toLowerCase() !== GST_RECLASS_CREATOR;

module.exports = { cleanLines, applyTags, isWrongKind, P2P_SOURCES, GST_RECLASS_CREATOR };
