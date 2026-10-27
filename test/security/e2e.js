// End-to-end browser test (2-step flow): upload -> (YTD column choice if asked) -> statement -> download.
// Usage: node test/security/e2e.js <html> <extract> <opening> <rajeev> <outXlsx> <screenshotPrefix>
const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const [html, ex, op, rj, outX, shot] = process.argv.slice(2);
  // PW_CHANNEL=msedge uses the installed Edge where Playwright's own browser cannot be downloaded
  const browser = await chromium.launch(process.env.PW_CHANNEL ? { channel: process.env.PW_CHANNEL } : {});
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
  const page = await ctx.newPage();
  const consoleErrors = []; const requests = [];
  page.on('console', (m) => { if (m.type() === 'error') consoleErrors.push(m.text()); });
  page.on('pageerror', (e) => consoleErrors.push('PAGEERROR ' + e.message));
  page.on('request', (r) => { if (!/^(file|blob|data):/.test(r.url())) requests.push(r.url()); });
  await page.goto('file://' + require('path').resolve(html));
  await page.setInputFiles('#f-extract', ex); await page.setInputFiles('#f-opening', op); await page.setInputFiles('#f-rajeev', rj);
  await page.click('#b-read');
  await page.waitForFunction(() => !document.getElementById('choice').classList.contains('hidden') || !document.getElementById('result').classList.contains('hidden') || !document.getElementById('errors').classList.contains('hidden'), null, { timeout: 60000 });
  if (await page.isVisible('#errors')) console.log('ERRORS', await page.textContent('#errors'));
  if (await page.isVisible('#choice')) {
    console.log('CHOICE:', await page.textContent('#choice-title'));
    await page.selectOption('#choice-body select', 'M'); await page.click('#b-choice');
  }
  await page.waitForSelector('#result:not(.hidden)', { timeout: 60000 });
  console.log('steps visible:', await page.$$eval('.stepper .st', (s) => s.map((x) => x.textContent.trim())).then((a) => a.join(' | ')));
  console.log('status:', await page.textContent('#r-status'));
  console.log('totals:', (await page.$$eval('#r-totals .kv', (k) => k.map((x) => x.textContent))).join(' | '));
  console.log('points:', await page.textContent('#r-points-sum'));
  await page.screenshot({ path: shot + '_result.png', fullPage: true });
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#b-dl')]);
  await dl.saveAs(outX);
  console.log('downloaded', dl.suggestedFilename(), fs.statSync(outX).size, 'bytes');
  // Back -> Continue must show the same result without re-reading
  await page.click('#b-back2'); await page.click('#b-read'); await page.waitForSelector('#sec-gen:not(.hidden)');
  console.log('back+continue ok:', await page.textContent('#r-status'));
  console.log('console errors:', consoleErrors.length ? consoleErrors : 'none');
  console.log('network requests (should be none):', requests.length ? requests : 'none');
  await browser.close();
})();
