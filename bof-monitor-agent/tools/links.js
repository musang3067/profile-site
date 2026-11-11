const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({userAgent:'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36'});
  for (const u of process.argv.slice(3)) {
    try { await p.goto(u, {waitUntil:'domcontentloaded', timeout:45000}); await p.waitForTimeout(5000);
      const l = await p.evaluate(() => [...document.querySelectorAll('a')].map(a => a.innerText.trim().slice(0,40)+' | '+a.href));
      console.log('== '+u); console.log(l.filter(x => new RegExp(process.argv[2],'i').test(x)).slice(0,25).join('\n'));
    } catch(e) { console.log(u, 'ERR', e.message.split('\n')[0]); }
  }
  await b.close();
})();
