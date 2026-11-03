// Query 수출입 실적 for an HS code (with sub-codes) over a year range; save the JSON the page receives.
const { chromium } = require(process.env.PW_CORE);
const fs = require('fs');
const [dir, hs, fromY, toY] = process.argv.slice(2);

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1400 } });
  const captured = [];
  page.on('dialog', async d => { console.log('DIALOG:', d.message()); await d.accept(); });
  page.on('request', r => { if (r.method()==='POST' && !/SubTitle|log/i.test(r.url())) console.log('POST', r.url()); });
  page.on('response', async r => {
    const ct = r.headers()['content-type'] || '';
    if (r.request().method() === 'POST' && /json|javascript|text\/plain/.test(ct)) {
      try { captured.push({ url: r.url(), post: r.request().postData(), body: await r.text() }); } catch (e) {}
    }
  });
  await page.goto('https://tradedata.go.kr/cts/index.do', { waitUntil: 'networkidle', timeout: 60000 });
  await page.evaluate(() => ets_f_prccMenuLoad('/cts/hmpg/openETS0100019Q.do', { menuId: 'ETS_MNK_10200000' }));
  await page.waitForSelector('#ETS0100019Q_hsSgn', { state: 'visible', timeout: 30000 });
  await page.fill('#ETS0100019Q_hsSgn', hs);
  await page.locator('#ETS0100019Q_hsSgn').locator('xpath=following::button[1]').click();
  await page.waitForTimeout(1500);
  await page.screenshot({ path: `${dir}/after_plus.png` });
  console.log('items label:', await page.locator('text=/\d+개 품목/').first().innerText().catch(()=>'?'));
  const sub = page.locator('#ETS0100019Q_subHsSgn');
  if (!(await sub.isChecked())) await sub.check({ force: true });
  await page.check('#radioYearPc', { force: true });
  await page.selectOption('#ETS0100019Q_formYearPc', fromY);
  await page.selectOption('#ETS0100019Q_toYearPc', toY);
  await page.selectOption('#ETS0100019Q_showPagingLine', '100').catch(()=>{});
  captured.length = 0;
  await page.locator('button:visible', { hasText: /^\s*조회\s*$/ }).first().click();
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(4000);
  fs.writeFileSync(`${dir}/result_${hs}.json`, JSON.stringify(captured, null, 1), 'utf8');
  const tableText = await page.$$eval('table', ts => ts.map(t => t.innerText).filter(x => /\d{4}/.test(x)).join('\n=====\n'));
  fs.writeFileSync(`${dir}/table_${hs}.txt`, tableText, 'utf8');
  await page.screenshot({ path: `${dir}/result_${hs}.png`, fullPage: true });
  console.log('captured', captured.length, 'tableChars', tableText.length);
  await browser.close();
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
