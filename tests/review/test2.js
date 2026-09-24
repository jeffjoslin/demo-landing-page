const { chromium } = require('playwright');
const path = require('path');
const S = __dirname;
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  await ctx.route(/cdnjs\.cloudflare\.com/, r => r.fulfill({ path: path.join(S, 'cdn', r.request().url().split('/').pop()), contentType: 'application/javascript' }));
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('dialog', d => d.accept());
  await page.goto('file://' + path.resolve(__dirname, '../../review.html'));
  await page.fill('#myName', 'Jane Smith'); await page.press('#myName', 'Tab');
  await page.setInputFiles('#sourceFile', path.join(S, 'source.pdf'));
  await page.setInputFiles('#docFile', path.join(S, 'threaded.docx'));
  await page.waitForSelector('#docView');
  const dump = async () => page.$$eval('.card', cs => cs.map(c => ({
    author: c.querySelector('.name').textContent, badge: c.querySelector('.badge').textContent,
    replies: [...c.querySelectorAll('.reply')].map(r => r.querySelector('.name').textContent + ': ' + r.querySelector('.text').textContent) })));
  console.log('summary:', await page.textContent('#commentSummary'));
  console.log('cards:', JSON.stringify(await dump()));
  console.log('resolved marks:', await page.locator('mark.hl.resolved').count(), 'toggle visible:', await page.isVisible('#resolvedToggle'));
  await page.screenshot({ path: path.join(S, 'threads.png') });
  await page.uncheck('#showResolved');
  console.log('hidden resolved -> cards:', JSON.stringify(await dump()), 'resolved marks:', await page.locator('mark.hl.resolved').count());
  await page.check('#showResolved');
  // add a new comment and save
  await page.evaluate(() => {
    const view = document.getElementById('docView');
    const w = document.createTreeWalker(view, NodeFilter.SHOW_TEXT); let n;
    while ((n = w.nextNode())) { const i = n.data.indexOf('consequential'); if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 13); getSelection().removeAllRanges(); getSelection().addRange(r); break; } }
    document.dispatchEvent(new MouseEvent('mouseup'));
  });
  await page.click('#floatBtn');
  await page.fill('#composerText', 'Define consequential damages.');
  await page.click('#composerAdd');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#exportBtn')]);
  await dl.saveAs(path.join(S, 'threaded-reviewed.docx'));
  await page.setInputFiles('#docFile', path.join(S, 'threaded-reviewed.docx'));
  await page.waitForTimeout(800);
  console.log('round trip:', JSON.stringify(await dump()));
  console.log('errors:', errors);
  await browser.close();
})().catch(e => { console.error('FAIL', e); process.exit(1); });
