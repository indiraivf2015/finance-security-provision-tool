// Release checks (no finance data needed). Run after `npm run build`, before every commit and deploy.
//  1. security.html (the committed file) has the approved fingerprint, and src/security/ still builds to it
//  2. Security is self-contained: src/security/ holds only its own files and loads nothing but ExcelJS
//  3. The site root holds only the pages listed in SHA256SUMS, and every listed hash is right
//  4. Every start-page link points to a built page
//  5. No network code in our own pages (ExcelJS excluded; the CSP also blocks network access)
'use strict';
const fs = require('fs'); const path = require('path');
const { buildSecurity, readLock, sha256 } = require('../build.js');

const ROOT = path.join(__dirname, '..');
const results = [];
const check = (name, ok, detail) => results.push({ name, ok: !!ok, detail });
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

// 1
const lock = readLock();
check('security.html has the approved fingerprint', sha256(read('security.html')) === lock, `approved ${lock}`);
check('src/security/ builds to the approved fingerprint', sha256(buildSecurity()) === lock);

// 2
const secFiles = fs.readdirSync(path.join(ROOT, 'src/security')).sort();
check('src/security/ contains only engine.js, ui.html, APPROVED.sha256', JSON.stringify(secFiles) === JSON.stringify(['APPROVED.sha256', 'engine.js', 'ui.html']), secFiles.join(', '));
const engine = read('src/security/engine.js');
const reqs = [...engine.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)].map((m) => m[1]);
check('Security engine loads nothing but exceljs', reqs.every((r) => r === 'exceljs') && !/\bimport\s*[({'"]|\bfrom\s+['"]/.test(engine), reqs.join(', '));

// 3
const sums = read('SHA256SUMS').trim().split('\n').map((l) => l.match(/^([0-9a-f]{64}) {2}(.+)$/));
check('SHA256SUMS is well formed', sums.every(Boolean));
const listed = sums.filter(Boolean).map((m) => m[2]).sort();
const atRoot = fs.readdirSync(ROOT).filter((f) => /\.html?$/i.test(f)).sort();
check('site root holds exactly the built pages', JSON.stringify(listed) === JSON.stringify(atRoot), `listed [${listed}] at root [${atRoot}]`);
for (const m of sums.filter(Boolean)) check(`${m[2]} matches SHA256SUMS`, fs.existsSync(path.join(ROOT, m[2])) && sha256(read(m[2])) === m[1]);
for (const f of listed) check(`${f} name is allowed by the Nginx rule ^/[a-z0-9-]+\\.html$`, /^[a-z0-9-]+\.html$/.test(f));
for (const conf of ['deploy-ec2/uat.nginx.conf', 'deploy-ec2/prod.nginx.conf']) {
  const locs = read(conf).split('\n').map((l) => l.trim()).filter((l) => l.startsWith('location'));
  check(`${conf} serves only "/" and top-level .html pages`, JSON.stringify(locs) === JSON.stringify(['location = / { try_files /index.html =404; }', 'location ~ ^/[a-z0-9-]+\\.html$ { }', 'location / { return 404; }']), locs.join(' | '));
}

// 4
const start = read('index.html');
const links = [...start.matchAll(/href="([^"#][^"]*)"/g)].map((m) => m[1]);
check('start page links to every other page', listed.filter((f) => f !== 'index.html').every((f) => links.includes(f)), links.join(', '));
check('every start-page link exists', links.length > 0 && links.every((l) => listed.includes(l)), links.join(', '));

// 5
const NET = /\bfetch\s*\(|XMLHttpRequest|WebSocket|EventSource|sendBeacon|importScripts|\bimport\s*\(/;
const exceljs = fs.readFileSync(path.join(ROOT, 'node_modules/exceljs/dist/exceljs.min.js'), 'utf8');
for (const f of listed) {
  if (f === 'security.html') continue; // frozen and fingerprint-locked; checked as approved
  const own = read(f).split(exceljs).join('');
  check(`${f} has no network code`, !NET.test(own), (own.match(NET) || [''])[0]);
  check(`${f} keeps the CSP with connect-src 'none'`, /<meta http-equiv="Content-Security-Policy" content="[^"]*connect-src 'none'/.test(own));
}

let bad = 0;
for (const r of results) { if (!r.ok) bad++; console.log(`${r.ok ? 'ok  ' : 'FAIL'} ${r.name}${!r.ok && r.detail ? ' — ' + r.detail : ''}`); }
console.log(`\n${results.length - bad} passed, ${bad} failed`);
process.exit(bad ? 1 : 0);
