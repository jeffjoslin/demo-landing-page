const { chromium } = require('playwright');
const path = require('path');
const S = __dirname;
const [,, input, output, shot] = process.argv;
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 1000 }, acceptDownloads: true });
  await ctx.route(/cdnjs\.cloudflare\.com/, r => r.fulfill({ path: path.join(S, 'cdn', r.request().url().split('/').pop()), contentType: 'application/javascript' }));
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('dialog', d => d.accept());
  await page.goto('file://' + path.resolve(__dirname, '../../review.html'));
  await page.fill('#myName', 'Jane Smith'); await page.press('#myName', 'Tab');
  await page.setInputFiles('#docFile', path.join(S, input));
  await page.waitForSelector('#docView');
  const dump = () => page.$$eval('#commentList .card', cs => cs.map(c => `${c.querySelector('.name').textContent} [${c.querySelector('.badge').textContent}] "${c.querySelector(':scope > .text')?.textContent}"` +
    [...c.querySelectorAll('.reply')].map(r => `\n      ↳ ${r.querySelector('.name').textContent}: ${r.querySelector('.text').textContent}`).join('')).join('\n  '));
  const card = t => page.locator('#commentList .card', { hasText: t });
  const reply = async (t, author, text) => {
    await card(t).getByRole('button', { name: 'Reply' }).click();
    const f = card(t).locator('.inline-form');
    await f.locator('input').fill(author);
    await f.locator('textarea').fill(text);
    await f.getByRole('button', { name: 'Reply' }).click();
  };
  console.log('loaded:\n  ' + await dump());

  const first = await page.locator('#commentList .card .text').first().textContent();
  if (input === 'threaded.docx') {
    await reply('Should liability be capped?', 'Jane Smith', 'Draft cap language added in section 4.');
    await card('Should liability be capped?').getByRole('button', { name: 'Resolve' }).click();
    await card('Customer proposes net 60').getByRole('button', { name: 'Reopen' }).click();
  } else {
    await reply('Customer proposes net 60', 'Review Team', 'We accept net 30.');
    await card('Customer proposes net 60').getByRole('button', { name: 'Resolve' }).click();
  }
  // new comment + reply + edit reply + resolve
  await page.evaluate(() => {
    const w = document.createTreeWalker(document.getElementById('docView'), NodeFilter.SHOW_TEXT); let n;
    while ((n = w.nextNode())) { const i = n.data.indexOf('consequential'); if (i >= 0) { const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 13); getSelection().removeAllRanges(); getSelection().addRange(r); break; } }
    document.dispatchEvent(new MouseEvent('mouseup'));
  });
  await page.click('#floatBtn');
  await page.fill('#composerText', 'Define consequential damages.');
  await page.click('#composerAdd');
  await reply('Define consequential damages.', 'Review Team', 'Typo reply');
  await card('Define consequential damages.').locator('.reply.mine', { hasText: 'Typo reply' }).getByRole('button', { name: 'Edit' }).click();
  await card('Define consequential damages.').locator('.inline-form textarea').fill('Use the definition from the MSA.');
  await card('Define consequential damages.').locator('.inline-form').getByRole('button', { name: 'Save' }).click();
  await card('Define consequential damages.').getByRole('button', { name: 'Resolve' }).click();
  console.log('summary:', await page.textContent('#commentSummary'));
  await page.reload();
  await page.waitForSelector('#docView'); await page.waitForTimeout(400);
  console.log('after reload:\n  ' + await dump());
  if (shot) await page.screenshot({ path: path.join(S, shot) });
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#exportBtn')]);
  await dl.saveAs(path.join(S, output));
  await page.setInputFiles('#docFile', path.join(S, output));
  await page.waitForTimeout(800);
  console.log('saved file reopened:\n  ' + await dump());
  console.log('errors:', errors);
  await browser.close();
})().catch(e => { console.error('FAIL', e); process.exit(1); });
