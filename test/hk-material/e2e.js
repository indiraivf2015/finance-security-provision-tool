// Browser test of hk-material.html with made-up files (see test/provision_e2e.js). Usage: node test/hk-material/e2e.js
'use strict';
const fs = require('fs'); const path = require('path'); const { execFileSync } = require('child_process');
const F = require('../fixtures');
(async () => {
  const out = path.join(__dirname, '..', '..', 'out'); fs.mkdirSync(out, { recursive: true });
  const f = (n) => path.join(out, n);
  const L = (code, bal, desc, x) => ({ code, name: 'Centre ' + code, account: 51611, bal, desc, ...(x || {}) });
  const codes = Array.from({ length: 12 }, (_, i) => 8001 + i);
  fs.writeFileSync(f('hkm_in_extract.xlsx'), await F.extract({ period: 'Sep-26', from: '2026-09-01', to: '2026-09-30', glFrom: '51611', glTo: '51611', lines: [
    L(8001, 2500, 'Housekeeping material for the month of August 2026'), L(8002, 3000, 'Housekeeping material for the month of August 2026', { source: 'Payables' }),
    L(8003, 999, 'Material for the month of March 2026'), L(8004, 700, 'Housekeeping material invoice'), L(8005, 80000, 'Housekeeping material for the month of August 2026')] }));
  fs.writeFileSync(f('hkm_in_opening.xlsx'), await F.statement({ title: 'Material Provision for the month of Aug-26',
    headers: ['Cost Center', 'Cost Center Name', 'Jul-Provision', 'Aug-Booking', 'Prov CF', 'Aug-Provision', 'YTD'], rows: codes.map((c) => [c, 'Centre ' + c, 0, 0, 0, 0, 10000 * (c - 8000)]) }));
  fs.writeFileSync(f('hkm_in_rajeev.xlsx'), await F.hkFile({ sheet: 'HK Sep-26', rows: codes.map((c) => ({ code: c, name: 'Centre ' + c, manpower: 118000, materials: 1000.25 * (c - 8000), rent: 5900 })), extraSheets: ['Security Sep-26'] }));
  execFileSync(process.execPath, [path.join(__dirname, '..', 'provision_e2e.js'), 'hk-material', f('hkm_in_extract.xlsx'), f('hkm_in_opening.xlsx'), f('hkm_in_rajeev.xlsx')], { stdio: 'inherit' });
})().catch((e) => { process.exit(e.status || 1); });
