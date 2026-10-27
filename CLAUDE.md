# CLAUDE.md: Finance Provision Tools

Read this before changing anything. These are **finance tools**: figures go into the ERP. Correctness beats features.

## What it is
Automates Indira IVF finance's monthly provision workings, one page per provision, opened from a start page.
Live: **Security** (SOP v1.0). Being added: **Housekeeping Expenses**, **Housekeeping Material**; about 8-10 more
distinct workings follow from Nov-26 (finance listed 44 expense heads). Each page reads 3 xlsx files in the browser
(ERP "IIVF Account Analysis" extract, last month's statement, Rajeev's provision file) and produces the statement
workbook. No backend: finance data never leaves the user's computer.

## Owner's standing instructions (do not break these)
1. **Follow the SOP exactly.** Do not add your own judgement. If something is unclear, ask; do not guess.
2. **Mirror finance's workbook format** where the SOP is silent (sheet names, columns, header colour FBE2D5, borders).
3. **Must work for any month.** Never hard-code anything specific to the sample data.
4. The owner is not a finance or Excel specialist. Explain changes in plain words.
5. **Discuss before implementing** when a request changes rules or numbers.
6. **Security is frozen.** `src/security/` is never edited and never uses the shared core. It may move onto the core
   only with the owner's explicit approval AND a passing golden-output test. Never run `npm run approve:security` unasked.
7. Each provision is ONE self-contained HTML file (ExcelJS and all code inlined): no external scripts, no network calls.

## Architecture
- Site root (what Nginx serves; all built by `npm run build`, never edited by hand):
  `index.html` start page, `security.html`, later `<provision-id>.html`; `SHA256SUMS` lists their hashes.
- `src/start/index.html`: start page template. Cards: Security (described in build.js) + one per `src/provisions/<id>/meta.json`.
- `src/security/`: FROZEN copy of the approved tool. `engine.js` (all logic, UMD), `ui.html` (2-step wizard; placeholders
  `/*__EXCELJS__*/`, `/*__ENGINE__*/`), `APPROVED.sha256` (the lock).
- `src/core/`: shared code for Housekeeping and later provisions (CommonJS; copied from Security's engine, never imported
  by it). `common` (months, FY, ₹ format), `excel` (strict cell reading, styles), `tagger` (Step 6, copied line for line),
  `extract` (Steps 1-4; accounts as a list), `clean` (Steps 7-9 + tags + "no year -> current FY"), `opening` (Step 13,
  columns from settings), `provision-file` (Rajeev's file: sheet by name, column by heading), `rules` (negative-CF rule,
  checks), `sheets` (Raw, Working, Points to check), `fy-statement` (the Security-style FY{cur}/FY{prev} statement,
  Steps 10-17), `engine` (createEngine(def): load/analyse/build), `ui/` (page.html + app.js: the shared screens).
- `src/provisions/<id>/index.js` + `meta.json`: one provision = settings + what differs; build.js bundles it with esbuild-wasm
  (`require('exceljs')` -> the inlined ExcelJS) into `<id>.html`. Adding a provision = a new folder; the start page card is automatic.
- Security of the pages: CSP `connect-src 'none'`, no unsafe-eval, DOM built with textContent only (no innerHTML).

## Housekeeping Expenses (src/provisions/hk-expenses, SOP v1.1 Oct 2026)
- Accounts 51610, 51612, 51613, 51614 only (51611 = Material); any other account stops the tool.
- Steps 5-9 as Security; Step 7 keeps "IIVF P2P" and "Payables" (owner's decision Oct 2026, same as Security).
- Pivot rows BALANCING_SEGMENT_DESC + SEGMENT1 + ATTRIBUTE3; statement takes each centre's total across vendors.
- Rajeev: sheet "HK <Mon-YY>", column found by heading "Manpower" (column I in the real file; the SOP figure says H).
  ÷118×100×1.135; same Storage Code rows added. Materials and Machine Rent are never read.
- Statement "Housekeeping Exp – Provision for the month of {Sept-26}", same 10 columns as Security, amounts with 2 decimals.
  Step 17 block in M5:O8 (below the header row so next month's reader never sees a second "YTD FY.." heading).
- Finance rules: no year -> current FY; negative-CF rule (top 10). 17 checks.
- NOT YET VERIFIED (no real HK extract yet): the report's GL parameter block for four non-contiguous accounts (not checked
  until seen), and the SOP's statement totals. Verified: every Part D figure and both Part E worked examples (test-data/hk-expected.json).

## Housekeeping Material (src/provisions/hk-material, SOP v1.0 Oct 2026; own simpler statement)
- Account 51611 only; report "GL Code From/To" must both be 51611; "Multiple Accounts" must be blank.
- Steps 5-9 as Security (Step 7 keeps "IIVF P2P" and "Payables"); no year -> current FY.
- Only current-FY bookings reach the statement (pivot filter Exp Year = FY{cur}); earlier-FY lines stay in Working and are
  each listed in Points to check (owner). SOP check 3 read as: pivot total = current-FY Working total, Booking total = pivot total.
- Rajeev: sheet "HK <Mon-YY>", column by heading "Materials", taken AS IS (no GST conversion); same code rows added.
- Statement "Material Provision for the month of {Sept-26}": Cost Center | Cost Center Name | {Aug}-Provision | {Sept}-Booking |
  Prov CF | {Sept}-Provision | YTD. Opening = last month's "YTD". Exact values kept, shown without paise (owner).
- Stops in April (owner: no year-end rule in the SOP).
- **Negative Prov CF: `NEGATIVE_CF_RULE` in index.js, AWAITING THE OWNER'S DECISION.** null = SOP Step 14 (carry the minus
  sign, listed in Points to check); { topN: 10 } = finance's rule. Both ways are tested. 10 checks (12 with the rule).
- Verified: Materials total and the three worked examples (one with a negative Prov CF). Not yet: SOP statement totals.

## Safeguards (all automatic)
- **Fingerprint lock:** build.js builds security.html first from `src/security/` only; if its sha256 differs from
  `APPROVED.sha256` the build stops and writes nothing. Approved: 3c90f0eca9db3c91e46801ec2de7029bd7c46653b261b94914701c054fc29e68.
- **Golden output:** `test-data/golden/security_sep26.xlsx` is finance's approved Sep-26 download. `test/golden.js` compares
  every part of the xlsx (values, formulas, saved results, styles, widths, merges); only created/modified time is ignored.
  Every provision gets its own golden file once finance approves a month. Golden files hold finance data: git-ignored.
- **Release check** (`npm run test:release`): lock, Security self-contained, site root = SHA256SUMS exactly, page names fit
  the Nginx rule, both Nginx configs unchanged in their rules, start-page links, no network code, CSP present.
- Versions pinned exactly: exceljs 4.4.0, esbuild-wasm 0.24.2, playwright 1.64.1. A different ExcelJS changes security.html.
- `.gitattributes` keeps LF line endings everywhere, so the fingerprint is the same on Windows and on the servers.

## Security rules implemented (SOP step -> code, src/security/engine.js)
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

## Finance rules added Oct 2026 (NOT in Security SOP v1.0; ask finance to add them as v1.1)
1. **Year not stated in the narration -> current FY** (FY of the posting month; FY28 from Apr-2027). No user step.
   Includes "month but no year" lines (finance confirmed FY27 for all). Lines are marked in Working "Remarks".
2. **Negative carry-forward:** in each of FY{cur} Pro CF and FY{prev} Pro CF separately, negatives are set to 0 and their
   total is taken from the 10 centres with the highest CF in that column, in proportion. Column total unchanged.
   Summary shows the final figure (plain number for changed centres, formula `=D-F` for others).
   Before/adjustment/after table is on the "Points to check" sheet. Constant: `SOP.negCfTopN = 10`.
   If negatives exceed what the top 10 can absorb -> NOT READY (check 16 fails).

## Security output workbook
Sheets: Summary, Pivot, Provision_Rajeev, Working, Raw, Points to check (amber tab when non-empty).
Summary: title A2 "Security Provision for the month of {Sept-26}", L2 =(I3+K3)/10^6, SUBTOTAL totals row 3, header row 4.
16 checks (SOP Section 04 checks 1-5, completeness, negative-CF rule). Status READY / NOT READY.

## Gotchas
- ExcelJS `cell.value` drops a formula result of 0: read `cell.model.result`. A formula with no saved value is an error, never 0.
- Amounts in rupees are floats from the opening file; compare with EPS = 0.005. Bookings are summed in paise.
- LibreOffice reuses cached values unless forced: use `test/security/lo_recalc_profile.xcu` (OOXMLRecalcMode=0). Prove it with a planted wrong value.
- The next month's opening reader matches headers exactly ("YTD FY27", "FY26 Pro CF"). Never add another column whose
  normalised header equals those.
- `test/security/run.js` "accepts all suggestions"; the page does not (the finance rule decides). For the golden test use
  `test/security/golden_run.js`, which does exactly what the page does.
- This PC: antivirus removes the native esbuild.exe (hence esbuild-wasm); Playwright cannot download its browser
  (company TLS inspection), so browser tests run with `PW_CHANNEL=msedge` (installed Edge).

## Testing (run all before any commit or deploy)
`npm run build`, then `npm test` (release check; Security: tagger 56, scenarios 20, Sep-26 compare, golden; core: tagger
parity + core-vs-approved-Security parity (1921 Summary figures); Housekeeping Expenses: 24 scenarios + SOP figures; Housekeeping Material: 20 scenarios + SOP figures), then the
browser tests `npm run test:e2e` (all four browser tests; prefix `PW_CHANNEL=msedge` on this PC).
Any change under `src/core/` must keep `test:core` passing: it proves the core still reproduces finance's approved Security statement.
Provision tests use made-up files built by `test/fixtures.js`; SOP figures (finance data) are in test-data/hk-expected.json.
Expected Sep-26 compare: 190 centres, 11 differ, all from one Payables bill finance excluded (details offline).
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
Workflow: change -> `npm run build` -> all tests -> commit + push -> on UAT `git pull` + `sha256sum -c SHA256SUMS`
-> finance checks -> same on production. Nginx changes go in BEFORE the `git pull` that needs them.
