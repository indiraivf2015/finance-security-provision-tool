// Browser test of hk-expenses.html with made-up files (see test/provision_e2e.js). Usage: node test/hk-expenses/e2e.js
'use strict';
const fs = require('fs'); const path = require('path'); const { execFileSync } = require('child_process');
const F = require('../fixtures');
(async () => {
  const out = path.join(__dirname, '..', '..', 'out'); fs.mkdirSync(out, { recursive: true });
  const f = (n) => path.join(out, n);
  const L = (code, bal, desc, x) => ({ code, name: 'Centre ' + code, account: 51610, bal, desc, ...(x || {}) });
  const codes = Array.from({ length: 12 }, (_, i) => 9001 + i);
  fs.writeFileSync(f('hk_in_extract.xlsx'), await F.extract({ period: 'Sep-26', from: '2026-09-01', to: '2026-09-30', lines: [
    L(9001, 50000, 'Housekeeping for the month of August 2026', { attr3: 'Alpha Services' }), L(9002, 20000, 'Housekeeping for the month of Jul-26', { account: 51613, attr3: 'Beta Facility', source: 'Payables' }),
    L(9003, 4000, 'Housekeeping for the month of March 2026', { account: 51614 }), L(9004, 1000, 'Housekeeping invoice'), L(9006, 700000, 'Housekeeping for the month of August 2026')] }));
  fs.writeFileSync(f('hk_in_opening.xlsx'), await F.statement({ title: 'Housekeeping Exp – Provision for the month of Aug-26',
    headers: ['Cost Center', 'Cost Center Name', 'FY27 Pro', 'FY26 Pro', 'FY27 Booking', 'FY26 Booking', 'FY27 Pro CF', 'FY26 Pro CF', 'Aug-26 Pro as per Rajeev', 'YTD FY27'],
    rows: codes.map((c) => [c, 'Centre ' + c, 0, 0, 0, 0, 0, 5000, 0, c === 9006 ? 600000 : 100000 * (c - 9000)]) }));
  fs.writeFileSync(f('hk_in_rajeev.xlsx'), await F.hkFile({ sheet: 'HK Sep-26', rows: codes.map((c) => ({ code: c, name: 'Centre ' + c, manpower: 118000, materials: 5000, rent: 5900 })), extraSheets: ['Security Sep-26'] }));
  execFileSync(process.execPath, [path.join(__dirname, '..', 'provision_e2e.js'), 'hk-expenses', f('hk_in_extract.xlsx'), f('hk_in_opening.xlsx'), f('hk_in_rajeev.xlsx')], { stdio: 'inherit' });
})().catch((e) => { process.exit(e.status || 1); });
