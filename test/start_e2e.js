// Browser test of the start page: every card opens its page, and nothing touches the network.
// Usage: node test/start_e2e.js   (PW_CHANNEL=msedge to use the installed Edge)
'use strict';
const { chromium } = require('playwright');
const path = require('path'); const fs = require('fs');
(async () => {
  const root = path.join(__dirname, '..');
  const browser = await chromium.launch(process.env.PW_CHANNEL ? { channel: process.env.PW_CHANNEL } : {});
  const page = await (await browser.newContext({ viewport: { width: 1280, height: 900 } })).newPage();
  const errors = []; const requests = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('request', (r) => { if (!/^(file|blob|data):/.test(r.url())) requests.push(r.url()); });
  const home = 'file://' + path.join(root, 'index.html').replace(/\\/g, '/');
  await page.goto(home);
  const cards = await page.$$eval('a.prov', (as) => as.map((a) => ({ href: a.getAttribute('href'), title: a.querySelector('b').textContent })));
  console.log('start page title:', await page.title(), '| cards:', cards.map((c) => c.title).join(', '));
  let bad = 0;
  for (const c of cards) {
    await page.goto(home);
    await Promise.all([page.waitForLoadState('load'), page.click(`a.prov[href="${c.href}"]`)]);
    const landed = path.basename(decodeURIComponent(new URL(page.url()).pathname));
    const ok = landed === c.href && fs.existsSync(path.join(root, c.href));
    if (!ok) bad++;
    console.log(`${ok ? 'ok  ' : 'FAIL'} "${c.title}" opens ${landed} — page title "${await page.title()}"`);
  }
  // Narrow screen: no horizontal scrolling
  await page.setViewportSize({ width: 375, height: 800 }); await page.goto(home);
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth);
  if (overflow) bad++;
  console.log(`${overflow ? 'FAIL' : 'ok  '} fits a phone-width screen`);
  console.log('console errors:', errors.length ? errors : 'none');
  console.log('network requests (should be none):', requests.length ? requests : 'none');
  if (errors.length || requests.length || !cards.length) bad++;
  await browser.close();
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
