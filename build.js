// Builds every page of the site. Each page is ONE self-contained HTML file (ExcelJS and all code inlined).
//
//   index.html     start page with one card per provision (src/start/)
//   security.html  the Security tool, built from src/security/ ONLY, exactly as it was approved.
//                  Its sha256 must equal src/security/APPROVED.sha256 or the build stops and writes nothing.
//   <id>.html      one page per provision in src/provisions/<id>/ (added from Step 2)
//
// Usage: node build.js                      build everything
//        node build.js --approve-security   accept a NEW Security fingerprint (only with the owner's explicit approval)
'use strict';
const fs = require('fs'); const path = require('path'); const crypto = require('crypto');

const ROOT = __dirname;
const p = (...a) => path.join(ROOT, ...a);
const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex');
const LOCK = p('src/security/APPROVED.sha256');

function noScriptClose(name, s) { if (/<\/script/i.test(s)) throw new Error(name + ' contains </script'); }

/* ---------- Security: frozen, assembled exactly as before the restructure ---------- */
function buildSecurity() {
  const tpl = fs.readFileSync(p('src/security/ui.html'), 'utf8');
  const exceljs = fs.readFileSync(p('node_modules/exceljs/dist/exceljs.min.js'), 'utf8');
  const engine = fs.readFileSync(p('src/security/engine.js'), 'utf8');
  for (const [n, s] of [['exceljs', exceljs], ['engine', engine]]) noScriptClose(n, s);
  return tpl.replace('/*__EXCELJS__*/', () => exceljs).replace('/*__ENGINE__*/', () => engine);
}

function readLock() {
  const line = fs.readFileSync(LOCK, 'utf8').split(/\r?\n/).map((l) => l.trim()).find((l) => l && !l.startsWith('#'));
  const m = line && line.match(/^([0-9a-f]{64})\b/);
  if (!m) throw new Error('src/security/APPROVED.sha256 has no fingerprint line');
  return m[1];
}

function writeLock(hash) {
  const text = fs.readFileSync(LOCK, 'utf8').split(/\r?\n/).map((l) => (/^[0-9a-f]{64}\b/.test(l.trim()) ? `${hash}  security.html` : l)).join('\n');
  fs.writeFileSync(LOCK, text);
}

/* ---------- Start page ---------- */
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// Security cannot describe itself (its source is frozen), so its card is described here.
const SECURITY_CARD = { file: 'security.html', title: 'Security Provision', text: 'Security services provision. ERP accounts 51701–51704.', sop: 'SOP v1.0 · Oct 2026' };

function provisionCards() {
  // From Step 2: one card per src/provisions/<id>/meta.json, in "order"
  const dir = p('src/provisions');
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((id) => fs.existsSync(path.join(dir, id, 'meta.json')))
    .map((id) => ({ id, ...JSON.parse(fs.readFileSync(path.join(dir, id, 'meta.json'), 'utf8')) }))
    .sort((a, b) => a.order - b.order).map((m) => ({ file: m.id + '.html', title: m.title, text: m.text, sop: m.sop }));
}

function buildStart(cards) {
  const tpl = fs.readFileSync(p('src/start/index.html'), 'utf8');
  const html = cards.map((c) => `      <a class="prov" href="${esc(c.file)}">
        <span class="ic"><svg class="i"><use href="#i-sheet"/></svg></span>
        <span class="tx"><b>${esc(c.title)}</b><small>${esc(c.text)}</small><small class="sop">${esc(c.sop)}</small></span>
        <svg class="i go"><use href="#i-arrow-right"/></svg>
      </a>`).join('\n');
  if (!tpl.includes('<!--__PROVISIONS__-->')) throw new Error('src/start/index.html has no <!--__PROVISIONS__--> placeholder');
  return tpl.replace('<!--__PROVISIONS__-->', () => html);
}

/* ---------- Main ---------- */
function main() {
  const approve = process.argv.includes('--approve-security');
  const pages = new Map(); // file name at the site root -> html

  // 1. Security first. A wrong fingerprint stops the whole build before anything is written.
  const sec = buildSecurity(); const secHash = sha256(sec); const locked = readLock();
  if (secHash !== locked) {
    if (!approve) {
      console.error('BUILD STOPPED: security.html would change.');
      console.error('  approved fingerprint : ' + locked);
      console.error('  built fingerprint    : ' + secHash);
      console.error('Security is frozen. Undo the change under src/security/ (or to ExcelJS), or, only with the owner\'s');
      console.error('explicit approval of a new Security version, run: npm run approve:security');
      process.exit(1);
    }
    writeLock(secHash);
    console.log('NEW SECURITY FINGERPRINT APPROVED: ' + locked + ' -> ' + secHash);
  }
  pages.set('security.html', sec);

  // 2. Provision pages (Step 2 onward)
  const cards = [SECURITY_CARD, ...provisionCards()];

  // 3. Start page
  pages.set('index.html', buildStart(cards));
  for (const c of cards) if (!pages.has(c.file)) throw new Error(`start page links to ${c.file}, which was not built`);

  // 4. Write: dist/ first, then the site root; SHA256SUMS lets the servers check every file after git pull
  fs.mkdirSync(p('dist'), { recursive: true });
  const sums = [];
  for (const [name, html] of pages) {
    fs.writeFileSync(p('dist', name), html); fs.writeFileSync(p(name), html);
    sums.push(`${sha256(html)}  ${name}`);
    console.log(`${name.padEnd(18)} ${(Buffer.byteLength(html) / 1024).toFixed(0).padStart(5)} KB  sha256 ${sha256(html)}`);
  }
  fs.writeFileSync(p('SHA256SUMS'), sums.sort((a, b) => a.slice(66).localeCompare(b.slice(66))).join('\n') + '\n');
  console.log('security.html matches the approved fingerprint.');
}

if (require.main === module) main();
module.exports = { buildSecurity, readLock, buildStart, provisionCards, SECURITY_CARD, sha256 };
