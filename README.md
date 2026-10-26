# Security Provision Tool

Automates Indira IVF finance's monthly **Security Provision Working** (SOP v1.0, Oct 2026) plus two finance rules added after testing (Oct 2026).
Input: 3 Excel files. Output: the statement workbook. Runs entirely in the browser; no backend, no data leaves the computer.

- **Live (UAT):** https://uatsecprov.indiraivf.in
- **Version:** 1.1.0

## How it works
`index.html` (top level, the file the servers serve) is a single self-contained file built from:

| File | Purpose |
|---|---|
| `src/engine.js` | All rules (SOP steps 1-16, checks, finance rules, Excel output). Works in Node and the browser. |
| `src/ui.html` | Screens (Upload -> Download), styles and screen logic. |
| `node_modules/exceljs/dist/exceljs.min.js` | Excel read/write library (ExcelJS 4.4.0), inlined at build time. |
| `build.js` | Joins the three into `dist/Security_Provision_Tool.html`. |

## Setup
```
npm install
npx playwright install chromium     # only for the browser test
```

## Build
```
npm run build        # writes dist/Security_Provision_Tool.html and copies it to index.html
```

## Test (copy sample files into test-data/ first, see test-data/README.md)
```
npm run test:tagger      # 56 year-reading cases (no data needed)
npm run test:scenarios   # 20 problem cases: wrong month, bad files, rules, edge cases
npm run test:sep         # Sep-26 run compared centre by centre with finance's statement
npm run test:e2e         # drives the real page in a browser, downloads the Excel
```
Formula check (optional, needs LibreOffice): convert an output with a profile using `test/lo_recalc_profile.xcu`
(forces recalculation), then `python test/recalc_compare.py <output.xlsx> <recalculated.xlsx>`.

## Deploy
- **EC2 / Nginx:** see `deploy-ec2/DEPLOY_EC2.md` (Nginx serves only index.html; source files and .git are not reachable). Servers pull this repo; update with `git pull`.
- After any change: `npm run build`, run the tests, commit `index.html`, then `git pull` on UAT, test, then production.
- Check the deployed file: `sha256sum index.html` must equal the hash printed by `npm run build`.

## User guide
Kept offline (it shows real figures); ask finance or the owner for the PDF.
