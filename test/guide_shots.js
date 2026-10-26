// Screenshots of each page (2-step flow) for the user guide.
const { chromium } = require('playwright');
(async () => {
  const [html, ex, op, rj, dir] = process.argv.slice(2);
  const b = await chromium.launch();
  const p = await (await b.newContext({ viewport: { width: 1200, height: 800 }, colorScheme: 'light', deviceScaleFactor: 2 })).newPage();
  await p.goto('file://' + html);
  await p.setInputFiles('#f-extract', ex); await p.setInputFiles('#f-opening', op); await p.setInputFiles('#f-rajeev', rj);
  await p.waitForTimeout(400); await p.screenshot({ path: dir + '/1_upload.png' });
  await p.click('#b-read'); await p.waitForSelector('#choice:not(.hidden)'); await p.waitForTimeout(300);
  await p.locator('#choice').screenshot({ path: dir + '/2_choice.png' });
  await p.selectOption('#choice-body select', 'M'); await p.click('#b-choice');
  await p.waitForSelector('#result:not(.hidden)'); await p.click('#r-points-sum'); await p.waitForTimeout(800);
  await p.screenshot({ path: dir + '/3_result.png', fullPage: true });
  await b.close();
})();
