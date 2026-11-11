const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const [out, ...urls] = process.argv.slice(2);
  const b = await chromium.launch();
  const ctx = await b.newContext({ userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36', viewport: {width: 1400, height: 1000} });
  for (const [i, u] of urls.entries()) {
    const p = await ctx.newPage();
    try {
      await p.goto(u, { waitUntil: 'domcontentloaded', timeout: 45000 });
      await p.waitForTimeout(6000);
      for (let k = 0; k < 8; k++) { await p.mouse.wheel(0, 2500); await p.waitForTimeout(400); }
      const t = await p.evaluate(() => document.body.innerText);
      const imgs = await p.evaluate(() => [...document.querySelectorAll('img')].map(i => i.currentSrc || i.src).filter(s => s && s.startsWith('http')));
      fs.writeFileSync(`${out}/${i}.txt`, `URL: ${u}\nFINAL: ${p.url()}\n\n${t}\n\nIMAGES:\n${[...new Set(imgs)].join('\n')}`);
      console.log(i, u, t.length);
    } catch (e) { console.log(i, u, 'ERR', e.message.split('\n')[0]); }
    await p.close();
  }
  await b.close();
})();
