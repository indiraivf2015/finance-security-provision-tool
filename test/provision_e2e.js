// Browser test of a provision page: upload the three files, get the result, download the Excel.
// The download must be identical (every part of the .xlsx except its timestamp) to what the Node engine produces
// from the same files, which proves the bundled page computes exactly what the tests check.
// Usage: node test/provision_e2e.js <id> <extract.xlsx> <opening.xlsx> <rajeev.xlsx>   (PW_CHANNEL=msedge for Edge)
'use strict';
const { chromium } = require('playwright');
const fs = require('fs'); const path = require('path'); const { execFileSync } = require('child_process');

(async () => {
  const [id, ex, op, rj] = process.argv.slice(2);
  const root = path.join(__dirname, '..'); const out = path.join(root, 'out'); fs.mkdirSync(out, { recursive: true });
  const E = require(path.join(root, 'src/provisions', id));
  const st = E.analyse(await E.load({ extract: fs.readFileSync(ex), opening: fs.readFileSync(op), rajeev: fs.readFileSync(rj) }), {});
  if (st.needChoice) throw new Error('this test expects files that need no question: ' + JSON.stringify(st.needChoice));
  const nodeOut = path.join(out, `e2e_${id}_node.xlsx`);
  fs.writeFileSync(nodeOut, Buffer.from((await E.build(st, new Map(), { now: new Date(0) })).buffer));

  const browser = await chromium.launch(process.env.PW_CHANNEL ? { channel: process.env.PW_CHANNEL } : {});
  const page = await (await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } })).newPage();
  const errors = []; const requests = [];
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', (e) => errors.push('PAGEERROR ' + e.message));
  page.on('request', (r) => { if (!/^(file|blob|data):/.test(r.url())) requests.push(r.url()); });
  await page.goto('file://' + path.join(root, id + '.html').split(path.sep).join('/'));
  console.log('page:', await page.title(), '|', await page.textContent('h1'), '|', await page.textContent('#foot'));
  await page.setInputFiles('#f-extract', ex); await page.setInputFiles('#f-opening', op); await page.setInputFiles('#f-rajeev', rj);
  await page.click('#b-read');
  await page.waitForFunction(() => !document.getElementById('result').classList.contains('hidden') || !document.getElementById('errors').classList.contains('hidden') || !document.getElementById('choice').classList.contains('hidden'), null, { timeout: 60000 });
  if (await page.isVisible('#errors')) throw new Error('page shows errors: ' + await page.textContent('#errors'));
  console.log('status:', await page.textContent('#r-status'));
  console.log('key figures:', (await page.$$eval('#r-keys .key', (k) => k.map((x) => x.textContent))).join(' | '));
  console.log('points:', await page.textContent('#r-points-sum'));
  await page.screenshot({ path: path.join(out, `e2e_${id}_result.png`), fullPage: true });
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#b-dl')]);
  const pageOut = path.join(out, `e2e_${id}_page.xlsx`); await dl.saveAs(pageOut);
  console.log('downloaded', dl.suggestedFilename());
  await page.click('a.chip[href="index.html"]'); await page.waitForLoadState('load');
  console.log('"All provisions" link opens:', path.basename(new URL(page.url()).pathname));
  await browser.close();
  console.log('console errors:', errors.length ? errors : 'none');
  console.log('network requests (should be none):', requests.length ? requests : 'none');
  let same = true;
  try { execFileSync(process.execPath, [path.join(__dirname, 'golden.js'), nodeOut, pageOut], { stdio: 'inherit' }); } catch (e) { same = false; }
  process.exit(same && !errors.length && !requests.length ? 0 : 1);
})().catch((e) => { console.error(e.stack || e); process.exit(1); });
