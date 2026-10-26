// Unit tests for SOP Step 6 tagging. Each case: [narration, period, expected fy (or null), expected sure]
const E = require('../src/engine.js');
const P = (label, y, m) => ({ label, y, m, cur: E.fyOf(y, m), prev: (E.fyOf(y, m) + 99) % 100 });
const SEP26 = P('Sep-26', 2026, 9), APR27 = P('Apr-27', 2027, 4), MAR27 = P('Mar-27', 2027, 3);
const cases = [
  ['BEING BILL NO. SA07DEL0125/2627 DT 07-07-2026 FOR HOUSEKEEPING SERVICES FOR THE MONTH OF JUNE 26 BOOKED', SEP26, 27, true],
  ['Provision for Material Expenses for the month of Aug-26', SEP26, 27, true],
  ['Being the GRN done Service Period- 01-06-2026 to 30-06-2026 Inv. No- SA18ASM0006/2627 Dt.- 09/07/2026', SEP26, 27, true],
  ['BEING INVOICE VC/MH/MAY-26/21 Housekeeping personnel 8 hr for the month of APR-2026', SEP26, 27, true],
  ['Being GRN Park form House keeping invoice no SA07DEL0121/2627 dt 07/07/2026', SEP26, 27, false],
  ['Housekeeping Material |  Invoice URL : https://vendorportal.indiraivf.in/x?docIndex=D9mygwJaY1yrQ4ZjuYaoBzZb2a6YQf9HQ6kRW', SEP26, null, false],
  ['BEING INVOICE BOOKED AGAINST HOUSEKEEPING MANPOWER FOR THE MONTH OF NOVEMBER 2025', SEP26, 26, true],
  ['BEING INVOICE BOOKED HOUSKEEPING MANPOWER SERVICES FOR THE MONTH OF JULY 2025 PLEASE APPROVE', SEP26, 26, true],
  ['Invoice of SMC Housekeeping for the month of December 2026', SEP26, null, false],
  ['BEING GRN PARKED FOR THE HOUSEKEEPING SERVICE FOR DEC AT GURUGRAM SEC-83 CENTRE INV NO-BNHR26GN000206 DT-23/02/2026', SEP26, 26, false],
  ['housekeeping expense for the month of feb-2026', SEP26, 26, true],
  ['Housekeeping Invoice for the Month Of DEC2025', SEP26, 26, true],
  ['Journal Import Created', SEP26, 27, true],
  ['"Journal Import Created"', APR27, 28, true],
  ['Journal Import Created for March 2026', SEP26, 26, true],
  ['SS4 SECURITY INVOICE NO.SS/25-26/1262 SECURITY GUARD SERVICES MONTH OF JANUARY 2025-26 (8 HOURS AND 31 DAYS)', SEP26, 26, true],
  ['SS4 SECURITY INVOICE NO.SS/26-27/380 (TOTAL WORKING DAYS=25 MONTH OF JUNE 2026-27)', SEP26, 27, true],
  ['Being security services for April 2026 to June 2026', SEP26, 27, true],
  ['Security for the month of March 2026 and April 2026', SEP26, null, false],
  ['Being amount for Jan-26', SEP26, 26, true],
  ['Marketing expenses decided', SEP26, null, false],
  ['Invoice dated May 15 2026', SEP26, 27, false],
  ['for the month of March 2027', APR27, 27, true],
  ['for the month of April 2027', APR27, 28, true],
  ['for the month of March 2025', SEP26, null, false],
  ['Service Period 15-03-2026 to 14-04-2026', SEP26, null, false],
  ['Bill for Sept-26', SEP26, 27, true],
  ['Bill for Mar-27', SEP26, null, false],
  ['BEING MARCH MONTH INVOICE PARKED AS PER INVOICE NUMBER SS/26-27/56', SEP26, 26, false],
  ['BEING APRIL MONTH INVOICE PARKED', SEP26, 27, false],
  ['Being grn created for security guard expenses for the month of may invoice ss/26-27/84 dt 15.06.2026', SEP26, 27, false],
  ['DMS Security staff salary monthly salary bill for Jul 26 Invoice No. MP/26-27/639 Invoice Date 01-08-2026.', SEP26, 27, true],
  ['Being GRN parked against Security Guard for Month of August 2026.', SEP26, 27, true],
  ['INVOICE NO-SS/26-27/141 DATE- 25/072026  JUNE NOMTH BILL COPY', SEP26, null, false],
  ['Being Securitty service amount for the month of April-26 with Invoice No-SS/26-27/46 Dated 12.05.2026', SEP26, 27, true],
  ['Being grn created for security guard expenses for the month of april26 invoice no. SS/26-27/51 Date 12-05-2026', SEP26, 27, true],
  ['Being GRN for the month of December 2026', MAR27, 27, true],
  ['Being GRN for the month of Jan 2027', MAR27, 27, true],
  ['Being GRN for the month of April 2027', MAR27, null, false],
  ['Saksham Services_SS/26-27/123_GR/26-27/23017', SEP26, null, false],
  ['R R Enterprises_RRE/2026-27/IVF/25_GR/26-27/22114', SEP26, null, false],
  ['', SEP26, null, false],
  // Reviewer's adversarial narrations — none may be auto-tagged
  ['Security charges for March, bill dated 10-Apr-2026', P('May-26', 2026, 5), null, false],
  ['Security charges Mar-Apr 2026', P('May-26', 2026, 5), null, false],
  ['Guard bill Feb to Apr-26', P('May-26', 2026, 5), null, false],
  ['INV-APR-26-0012 security for March month', P('May-26', 2026, 5), null, false],
  ['Invoice date Sep 25, 2026', SEP26, 27, false],
  ['Security services Mar 26-27 arrears', SEP26, null, false],
  ['Service period 04/01/2026 to 04/02/2026', P('May-26', 2026, 5), null, false],
  ['FY 2025-26 bill for the month of April', SEP26, null, false],
  ['Security for the month of June 2026, invoice dated 05-Jul-2026', SEP26, 27, true],
  ['Security for the month of March 2026, invoice dated 05-Apr-2026', SEP26, 26, false],
  ['for the month of June 2026 FY 2026-27', SEP26, 27, true],
  ['for the month of June 2026 FY 2025-26', SEP26, null, false],
  ['Feb-26 to Apr-26 security bill for the month of Apr-26', SEP26, null, false],
  ['for the month of April 2026 & May 2026', SEP26, 27, true],
];
let fail = 0;
for (const [d, p, fy, sure] of cases) {
  const r = E.proposeFY(d, p);
  const ok = r.fy === fy && r.sure === sure;
  if (!ok) { fail++; console.log('FAIL', JSON.stringify(d).slice(0, 90), p.label, 'expected', fy, sure, 'got', r.fy, r.sure, '|', r.basis); }
}
console.log(`${cases.length - fail}/${cases.length} tagging cases pass`);
process.exit(fail ? 1 : 0);
