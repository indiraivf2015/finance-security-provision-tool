# CLAUDE.md: Security Provision Tool

Read this before changing anything. This is a **finance tool**: figures go into the ERP. Correctness beats features.

## What it is
Automates Indira IVF finance's monthly **Security Provision Working** (SOP v1.0, Oct 2026, 16 steps).
Inputs (3 xlsx): Oracle ERP "IIVF Account Analysis" extract (accounts 51701-51704, one month), last month's
statement, Rajeev's provision file. Output: statement workbook. Runs fully in the browser (offline, no backend).

## Owner's standing instructions (do not break these)
1. **Follow the SOP exactly.** Do not add your own judgement. If something is unclear, ask; do not guess.
2. **Mirror finance's workbook format** where the SOP is silent (sheet names, columns, header colour FBE2D5, borders).
3. **Must work for any month.** Never hard-code anything specific to the sample data.
4. The owner is not a finance or Excel specialist. Explain changes in plain words.
5. **Discuss before implementing** when a request changes rules or numbers.

## Architecture
- `src/engine.js`: all logic (UMD, runs in Node and the browser). Key functions: `load`, `analyse` (read + filter + tag),
  `compute` (pivot, statement, checks, finance rules, points), `build` (Excel), `proposeFY` (year from narration).
- `src/ui.html`: 2-step wizard (Upload -> Download). Placeholders `/*__EXCELJS__*/` and `/*__ENGINE__*/` are filled by `build.js`.
- `index.html`: the built file the servers serve. Rebuild with `npm run build`; never edit it by hand.
- Security: CSP `connect-src 'none'`, no unsafe-eval, DOM built with textContent only (no innerHTML).

## Rules implemented (SOP step -> code)
- Steps 1-4: extract must be one period, full month (From/To date params), accounts 51701-51704, no centre filter.
- Step 5-6: "Exp Year" column. `proposeFY` tags a line only when the narration states month+year (or FY form like
  "June 2026-27", or a service period). "Journal Import Created" = current FY.
- Steps 7-8: keep JE source "IIVF P2P"/"Payables" and Spreadsheet lines created by taxverification2@indiraivf.in (GST reclass).
- Step 9: blank ATTRIBUTE3 filled with a formula to PARTY_NAME.
- Step 10: pivot by SEGMENT1 x Exp Year (sums in paise, exact).
- Steps 11-12: Rajeev final = Provision with GST / 118 x 100 x 1.135.
- Step 13: opening = last month's "YTD FY{cur}" and "FY{prev} Pro CF". If two YTD columns, the user picks one.
- Steps 14-16: CF = Pro - Booking; YTD = FY{cur} Pro CF + Rajeev.
- April (new FY) stops with a message: the SOP has no year-end roll-over rule yet.

## Finance rules added Oct 2026 (NOT in SOP v1.0; ask finance to add them as v1.1)
1. **Year not stated in the narration -> current FY** (FY of the posting month; FY28 from Apr-2027). No user step.
   Includes "month but no year" lines (finance confirmed FY27 for all). Lines are marked in Working "Remarks".
2. **Negative carry-forward:** in each of FY{cur} Pro CF and FY{prev} Pro CF separately, negatives are set to 0 and their
   total is taken from the 10 centres with the highest CF in that column, in proportion. Column total unchanged.
   Summary shows the final figure (plain number for changed centres, formula `=D-F` for others).
   Before/adjustment/after table is on the "Points to check" sheet. Constant: `SOP.negCfTopN = 10`.
   If negatives exceed what the top 10 can absorb -> NOT READY (check 16 fails).

## Output workbook
Sheets: Summary, Pivot, Provision_Rajeev, Working, Raw, Points to check (amber tab when non-empty).
Summary: title A2 "Security Provision for the month of {Sept-26}", L2 =(I3+K3)/10^6, SUBTOTAL totals row 3, header row 4.
16 checks (SOP Section 04 checks 1-5, completeness, negative-CF rule). Status READY / NOT READY.

## Gotchas
- ExcelJS `cell.value` drops a formula result of 0: read `cell.model.result`. A formula with no saved value is an error, never 0.
- Amounts in rupees are floats from the opening file; compare with EPS = 0.005. Bookings are summed in paise.
- LibreOffice reuses cached values unless forced: use `test/lo_recalc_profile.xcu` (OOXMLRecalcMode=0). Prove it with a planted wrong value.
- The next month's opening reader matches headers exactly ("YTD FY27", "FY26 Pro CF"). Never add another column whose
  normalised header equals those.

## Testing (run all before any deploy)
`npm run test:tagger` (56), `npm run test:scenarios` (20), `npm run test:sep` (compare with finance), `npm run test:e2e`.
Sample files go in `test-data/` (git-ignored; see test-data/README.md). Never commit finance data.

## Validation history
- Sep-26 and Jun-26 runs were compared centre by centre with finance's statements; every difference was traced
  to a known cause (manual edits in finance's file, or Payables lines finance excluded). Details are kept offline.

## Open questions with finance
1. Exclude Payables lines created by an individual user? (Seen in Sep-26 and Jun-26.)
2. Some centres are missing from Rajeev's file every month.
3. Year-end roll-over rule for April.
4. Is "top 10, in proportion" the confirmed split for negatives?
5. Add both Oct-2026 rules to the SOP (v1.1).

## Deploy
See `deploy-ec2/DEPLOY_EC2.md`. UAT https://uatsecprov.indiraivf.in, production secprov.indiraivf.in.
Workflow: change -> `npm run build` -> all tests -> commit + push -> `git pull` on UAT -> finance checks -> `git pull` on production.
