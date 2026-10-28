// Housekeeping year tagging (src/core/tagger.js, changes of Oct 2026): a year written in the narration is used;
// "current FY by default" only when no year is written at all. Each case: [narration, expected fy (null = default), sure, yearWritten]
// Usage: node test/core/tagger_hk.js
'use strict';
const { proposeFY } = require('../../src/core/tagger');
const { fyOf } = require('../../src/core/common');
const P = (label, y, m) => ({ label, y, m, cur: fyOf(y, m), prev: (fyOf(y, m) + 99) % 100 });
const SEP26 = P('Sep-26', 2026, 9);

const cases = [
  // The 10 Sep-26 Housekeeping Material lines (finance's own tags were FY 26; the "May 2025" line finance tagged FY 27)
  ['Being towards Mysore centre house keeping material invoice booked for the month of Nov_25. against Invoice no BNKA25SBR000127 Dt 02/01/2026.', 26, true, true],
  ['Being towards Mysore centre house keeping material invoice booked for the month of Aug_25. against invoice no BNKA25BR000216 Dt 22/01/2026.', 26, true, true],
  ['bEING HOUSEKEEPING MATARIAL PROVIDED BY DEC-25', 26, true, true],
  ['BEING HOUSE KEEPING MATARIAL PROVIDED BY SMC OCT-25', 26, true, true],
  ['baeing GRN done for housekeeping materail received from smc in july 25 (Cn for excess amount is attached)', 26, true, true],
  ['invoice no BNMH25SPN000122 invoice date 02-01/2026 smc material chares opd baramari centre', 26, true, true],
  ['BEING INVOICE BOOKED AGAINST HK MATERIAL FOR THE MONTH OF OF AUGUST 2025 PLEASE APPROVE', 26, true, true],
  ['Being Invoice booked for the purchase of house keeping material for the month of DECEMBER5 2025 Invoice no BNKA26SBR000093', 26, true, true],
  ['Being Invoice booked for the purchase of house keeping material for the month of May 2025 Invoice No BNKASBN000273/25 on 19.09.2026', 26, true, true],
  ['BULDHANA CENTER HOUSEKEPING MATERIAL DECAMBER 2025', 26, true, true],
  // Same kinds of text, current-year months
  ['Housekeeping material provided by AUG-26', 27, true, true],
  ['Housekeeping material for the month of Jul_26', 27, true, true],
  ['for the month of Septmber 2025', 26, true, true],
  // No year written at all -> default (not tagged, no year written)
  ['Housekeeping Material |  Invoice URL : https://vendorportal.indiraivf.in/x', null, false, false],
  ['Housekeeping Personnel 8 Hr (31 days)', null, false, false],
  ['BEING APRIL MONTH INVOICE PARKED', 27, false, false],
  ['Housekeeping material for the month of DEC', 26, false, false],
  ['Housekeeping service VC/MH/MAY-26/21', 27, false, false], // invoice-number month: a suggestion only, not a written year
  // A year is written but the years disagree -> default, marked yearWritten
  ['Security for the month of March 2026 and April 2026', null, false, true],
  ['Housekeeping for the month of March 2025', null, false, true],
  ['Housekeeping for the month of March, bill dated 05-04-2026', null, false, true],
  ['Invoice of SMC Housekeeping for the month of December 2026', null, false, true], // a month later than the posting month
  ['Housekeeping material DECAMBER 2026', null, false, true],
  // Words that only look like a misspelt month are not months
  ['Agustya Enterprises housekeeping bill', null, false, false],
];

let fail = 0;
for (const [d, fy, sure, written] of cases) {
  const p = proposeFY(d, SEP26);
  if (p.fy !== fy || p.sure !== sure || p.yearWritten !== written) { fail++; console.log(`FAIL "${d}"\n     expected FY ${fy} sure ${sure} yearWritten ${written}; got FY ${p.fy} sure ${p.sure} yearWritten ${p.yearWritten} (${p.basis})`); }
}
console.log(`${cases.length - fail}/${cases.length} Housekeeping tagger cases pass`);
process.exit(fail ? 1 : 0);
