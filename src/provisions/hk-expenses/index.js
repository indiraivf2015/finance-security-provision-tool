// Housekeeping Expenses provision: SOP "Monthly Housekeeping Provision Working" v1.1 (Oct 2026), Steps 1-17.
// Same statement layout as Security (src/core/fy-statement.js); what differs is set here.
//
// SOP step -> setting
//   Step 2      ERP accounts 51610, 51612, 51613, 51614 (51611 is the Material working, not this one)
//   Steps 5-9   as Security (src/core/clean.js). Step 7: the SOP says "keep … source Payables"; in the ERP extract
//               P2P invoice bookings appear as "IIVF P2P" or "Payables" and both are kept, exactly as for Security
//               (owner's decision, Oct 2026: the Security SOP has the same wording and finance's working keeps both).
//   Step 10     pivot rows BALANCING_SEGMENT_DESC + SEGMENT1 + ATTRIBUTE3 (vendor), columns Exp Year;
//               the statement takes each centre's total across vendors (Step 14)
//   Steps 11-12 Rajeev's sheet "HK <Mon-YY>", Manpower column only ("Per Month Cost with GST"),
//               ÷ 118 × 100, then × 1.135. Materials and Machine Rent are not read at all.
//               Rows with the same Storage Code (e.g. Main + OPD) are added together.
//   Step 13     opening = last month's "YTD FY{cur}" and "FY{prev} Pro CF"
//   Step 17     summary block: FY{prev} Pro CF, YTD FY{cur}, Total (also in millions), Every month = Total ÷ 4
//   Finance rules (Oct 2026, confirmed by finance for Housekeeping Expenses): year not in the narration -> current FY;
//               negative carry-forward -> 0, taken from the 10 highest centres in proportion.
//
// NOT YET VERIFIED (awaiting finance's Sep-26 HK extract): how the report's parameter block shows the four accounts
// ("GL Code From/To" or "Multiple Accounts"). Until then the parameters are not checked; every line's account is.
'use strict';
const { createEngine } = require('../../core/engine');
const fyStatement = require('../../core/fy-statement');
const { fyLabelHeader } = require('../../core/common');
const { FMT_AMT2 } = require('../../core/excel');

const VERSION = '1.0.0';

const statement = {
  toolName: 'Housekeeping Expenses Provision Tool',
  title: (L) => `Housekeeping Exp – Provision for the month of ${L}`,
  pivotByVendor: true,
  amountText: 'Manpower',
  gst: { divisor: 118, basicFactor: 100, finalFactor: 1.135 },
  negativeCarryForward: { topN: 10 },
  summaryBlock: true,
  amountFormat: FMT_AMT2,
};

module.exports = createEngine({
  toolName: statement.toolName,
  version: VERSION,
  sopVersion: 'SOP v1.1 · October 2026',
  extract: {
    accounts: [51610, 51612, 51613, 51614],
    accountsText: 'accounts 51610, 51612, 51613 and 51614',
    glParams: null,                 // not checked until a real extract shows how the report lists the four accounts
    multipleAccountsAllowed: true,  // the SOP selects four separate accounts
  },
  opening: (period) => ({
    columns: [{ key: 'ytd', header: `YTD ${fyLabelHeader(period.cur)}`, choosable: true }, { key: 'cf', header: `${fyLabelHeader(period.prev)} Pro CF` }],
    aprilNote: true,
    mergeDuplicates: true, // owner's decision (Oct 2026): a centre on several rows is added together and listed on "Points to check"
  }),
  provisionFile: { sheet: { test: /^\s*hk(?![a-z])/i, text: '"HK <Mon-YY>"' },amount: { startsWith: 'manpower', text: 'Manpower' }, noun: 'housekeeping' },
  stopInApril: false, // as Security: in April last month's statement has no "YTD FY{new}" column, and the reader stops with the April note
  compute: (state) => fyStatement.compute(state, statement),
  buildWorkbook: (state, R, meta) => fyStatement.buildWorkbook(state, R, meta, statement, VERSION),
  ui: {
    title: 'Housekeeping Expenses Provision',
    files: {
      extract: { title: 'Oracle ERP extract', desc: 'IIVF Account Analysis report · accounts 51610, 51612, 51613, 51614 · one full month, raw download' },
      opening: { title: 'Last month\'s statement', desc: 'Previous month\'s Housekeeping Exp provision statement · gives the opening balances' },
      rajeev: { title: 'Provision file', desc: 'This month\'s file from Rajeev · sheet "HK <month>" · Manpower column only' },
    },
    outputName: (period) => `Housekeeping Exp_${period.display}.xlsx`,
    sheets: 'Summary, Pivot, Provision_Rajeev, Working, Raw, Points to check',
    note: 'The statement is built and the checks have run. Lines with no year in the description are counted as the current financial year.',
    pointsNote: 'Unusual items found in this month\'s files, for information. The same list, plus the negative carry-forward adjustment table, is in the "Points to check" sheet of the Excel.',
    keys: (R) => {
      const C = fyLabelHeader(R.period.cur), P = fyLabelHeader(R.period.prev);
      return [
        ['Total provision', R.T.total, `${P} Pro CF + YTD ${C} · ${(R.T.total / 1e6).toFixed(2)} Mn`],
        ['YTD ' + C, R.T.ytd, 'Closing provision'],
        [P + ' Pro CF', R.T.cfPrev, 'Carried to next month\'s opening'],
        ['Every month', R.T.everyMonth, 'Total ÷ 4'],
      ];
    },
    totals: (R) => {
      const C = fyLabelHeader(R.period.cur), P = fyLabelHeader(R.period.prev);
      return [[C + ' Pro', R.T.proCur], [P + ' Pro', R.T.proPrev], [C + ' Booking', R.T.bkCur], [P + ' Booking', R.T.bkPrev], [C + ' Pro CF', R.T.cfCur], [P + ' Pro CF', R.T.cfPrev], [R.period.display + ' Pro as per Rajeev', R.T.raj], ['YTD ' + C, R.T.ytd]];
    },
  },
});
