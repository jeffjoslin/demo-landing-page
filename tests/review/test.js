const { chromium } = require('playwright');
const path = require('path'); const fs = require('fs');
const S = __dirname;
const APP = 'file://' + path.resolve(__dirname, '../../review.html');
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  await ctx.route(/cdnjs\.cloudflare\.com/, route => {
    const f = route.request().url().split('/').pop();
    fs.existsSync(path.join(S, 'cdn', f)) ? route.fulfill({ path: path.join(S, 'cdn', f), contentType: 'application/javascript' }) : route.continue();
  });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('dialog', d => d.accept());
  await page.goto(APP);
  await page.fill('#myName', 'Jane Smith'); await page.press('#myName', 'Tab');
  await page.setInputFiles('#sourceFile', path.join(S, 'source.pdf'));
  await page.setInputFiles('#docFile', path.join(S, 'customer.docx'));
  await page.waitForSelector('#docView');
  await page.waitForSelector('.pdf-pages canvas');
  console.log('existing cards:', await page.locator('.card.ext').count(), await page.locator('.card.ext .name').allTextContents());

  async function comment(phrase, author, text) {
    await page.evaluate(phrase => {
      const view = document.getElementById('docView');
      const w = document.createTreeWalker(view, NodeFilter.SHOW_TEXT);
      const nodes = []; let all = '';
      while (w.nextNode()) { nodes.push([w.currentNode, all.length]); all += w.currentNode.data; }
      const i = all.indexOf(phrase); if (i < 0) throw new Error('phrase not found ' + phrase);
      const at = pos => { for (let k = nodes.length - 1; k >= 0; k--) if (nodes[k][1] <= pos) return [nodes[k][0], pos - nodes[k][1]]; };
      const [sn, so] = at(i); const [en, eo] = at(i + phrase.length);
      const r = document.createRange(); r.setStart(sn, so); r.setEnd(en, eo);
      const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
      document.dispatchEvent(new MouseEvent('mouseup'));
    }, phrase);
    await page.waitForSelector('#floatBtn', { state: 'visible' });
    await page.click('#floatBtn');
    const quote = await page.textContent('#composerQuote');
    await page.fill('#composerAuthor', author);
    await page.fill('#composerText', text);
    await page.click('#composerAdd');
    return quote;
  }
  console.log('Q1', JSON.stringify(await comment('deliver the Services', 'Jane Smith', 'Please define Services.')));
  console.log('Q2', JSON.stringify(await comment('sixty (60)', 'Review Team', 'We cannot accept 60 days.\nPlease revert to 30.')));
  console.log('Q3', JSON.stringify(await comment('$10,000', 'Jane Smith', 'Fee should be $8,000.')));
  console.log('Q4', JSON.stringify(await comment('Effective Date.Payment terms', 'Jane Smith', 'Cross-paragraph comment.')));
  console.log('Q5', JSON.stringify(await comment('indirect', 'Jane Smith', 'temp')));
  // edit Q1, delete Q5
  const cardFor = t => page.locator('.card', { hasText: t });
  await cardFor('Please define Services.').locator('button', { hasText: 'Edit' }).click();
  await cardFor('Jane Smith').locator('textarea').first().fill('Please define "Services" precisely.');
  await page.locator('.card button', { hasText: 'Save' }).click();
  await cardFor('temp').locator('button', { hasText: 'Delete' }).click();
  console.log('after edit/delete new cards:', await page.locator('.card:not(.ext)').count());
  await page.screenshot({ path: path.join(S, 'desktop.png') });

  await page.reload();
  await page.waitForSelector('#docView');
  await page.waitForTimeout(500);
  console.log('after reload new:', await page.locator('.card:not(.ext)').count(), 'ext:', await page.locator('.card.ext').count(),
    'source restored:', await page.locator('.pdf-pages canvas').count() > 0, 'name:', await page.inputValue('#myName'));
  console.log('texts:', await page.locator('.card .text').allTextContents());

  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#exportBtn')]);
  await dl.saveAs(path.join(S, 'reviewed.docx'));
  console.log('download name:', dl.suggestedFilename());

  // round trip: load the saved file back in; all comments should show as existing
  await page.setInputFiles('#docFile', path.join(S, 'reviewed.docx'));
  await page.waitForTimeout(800);
  console.log('round trip ext cards:', await page.locator('.card.ext').count(), await page.locator('.card.ext .quote').allTextContents());

  const m = await browser.newContext({ viewport: { width: 390, height: 844 } });
  await m.route(/cdnjs\.cloudflare\.com/, route => route.fulfill({ path: path.join(S, 'cdn', route.request().url().split('/').pop()), contentType: 'application/javascript' }));
  const mp = await m.newPage();
  await mp.goto(APP);
  await mp.setInputFiles('#docFile', path.join(S, 'customer.docx'));
  await mp.waitForSelector('#docView');
  const sw = await mp.evaluate(() => document.documentElement.scrollWidth);
  console.log('mobile scrollWidth', sw);
  await mp.screenshot({ path: path.join(S, 'mobile.png') });
  console.log('errors:', errors);
  await browser.close();
})().catch(e => { console.error('FAIL', e); process.exit(1); });
