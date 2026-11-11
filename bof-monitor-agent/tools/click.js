const { chromium } = require('playwright'); const fs=require('fs');
(async () => {
  const [u, label, out] = process.argv.slice(2);
  const b = await chromium.launch(); const p = await b.newPage({userAgent:'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36', viewport:{width:1400,height:1000}});
  await p.goto(u, {waitUntil:'domcontentloaded', timeout:60000}); await p.waitForTimeout(8000);
  if (label) { try { await p.getByText(label, {exact:true}).first().click({timeout:8000}); await p.waitForTimeout(4000);} catch(e){console.log('noclick', e.message.split('\n')[0]);} }
  for (let k=0;k<10;k++){await p.mouse.wheel(0,2500); await p.waitForTimeout(400);}
  fs.writeFileSync(out, await p.evaluate(() => document.body.innerText)); console.log('ok');
  await b.close();
})();
