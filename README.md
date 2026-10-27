# Finance Provision Tools

Automates Indira IVF finance's monthly provision workings. One page per provision, opened from a start page:

| Page | Provision | Status |
|---|---|---|
| `index.html` | Start page | |
| `security.html` | Security (SOP v1.0, Oct 2026, plus two finance rules) | Live, approved by finance. Frozen. |
| `hk-expenses.html` | Housekeeping Expenses (SOP v1.1, Oct 2026) | Built; awaiting finance's sample HK extract to verify the totals |
| `hk-material.html` | Housekeeping Material (SOP v1.0, Oct 2026) | Built; negative carry-forward setting awaiting owner's decision; awaiting finance's sample extract |

Each page reads 3 Excel files and produces the statement workbook. Everything runs in the browser: no backend,
no network access, no data leaves the computer. Each page is one self-contained HTML file (ExcelJS inlined).

- **Live:** https://uatsecprov.indiraivf.in (UAT), secprov.indiraivf.in (production)

## How it is built
| Source | Purpose |
|---|---|
| `src/start/index.html` | Start page template (one card per provision). |
| `src/security/` | The approved Security tool, frozen: `engine.js` (rules), `ui.html` (screens), `APPROVED.sha256` (its fingerprint). |
| `src/core/`, `src/provisions/<id>/` | Shared code and one module per provision (Housekeeping onward). |
| `build.js` | Builds every page into the site root and `dist/`, and writes `SHA256SUMS`. |

`npm run build` builds Security first and **stops without writing anything** if `security.html` would differ from
the approved fingerprint. A new Security version needs the owner's explicit approval (`npm run approve:security`).

## Setup
```
npm ci
npx playwright install chromium     # only for the browser tests; or use the installed Edge (PW_CHANNEL=msedge)
```

## Build
```
npm run build
```

## Test (copy sample files into test-data/ first, see test-data/README.md)
```
npm test                       # release check + Security, core and both Housekeeping tests (not the browser ones)
npm run test:release           # fingerprint lock, page list, Nginx rules, no network code (no data needed)
npm run test:security:tagger    # 56 year-reading cases (no data needed)
npm run test:security:scenarios # 20 problem cases: wrong month, bad files, rules, edge cases
npm run test:security:sep      # Sep-26 run compared centre by centre with finance's statement
npm run test:security:golden   # Sep-26 output identical to finance's approved file, cell for cell
npm run test:security:e2e      # the real page in a browser; the download must equal the approved file
npm run test:core              # shared core: same year tagging as Security; reproduces the approved Security statement
npm run test:hk-expenses       # Housekeeping Expenses: 24 rules/problem cases (made-up data) + the SOP's own figures
npm run test:hk-expenses:e2e   # hk-expenses.html in a browser; the download must equal the Node result
npm run test:hk-material       # Housekeeping Material: 20 rules/problem cases (both negative-CF settings) + the SOP's figures
npm run test:hk-material:e2e   # hk-material.html in a browser; the download must equal the Node result
npm run test:e2e               # all browser tests
npm run test:start:e2e         # start page: every card opens its page, no network requests
```
Browser tests: set `PW_CHANNEL=msedge` to use the installed Edge if Playwright cannot download its browser.

Formula check (optional, needs LibreOffice): convert an output with a profile using `test/security/lo_recalc_profile.xcu`
(forces recalculation), then `node test/security/recalc_compare.js <output.xlsx> <recalculated.xlsx>`.

## Deploy
- **EC2 / Nginx:** see `deploy-ec2/DEPLOY_EC2.md`. Nginx serves only `/` and top-level `.html` pages; source files,
  `.git` and everything else return 404. Servers update with `git pull`.
- After any change: `npm run build`, run all tests, commit the built pages and `SHA256SUMS`, then on UAT
  `git pull` and `sha256sum -c SHA256SUMS`, test, then production.

## User guide
Kept offline (it shows real figures); ask finance or the owner for the PDF.
