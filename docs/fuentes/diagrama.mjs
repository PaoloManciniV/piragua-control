import { chromium } from 'playwright';
const b = await chromium.launch();
const pg = await b.newPage({deviceScaleFactor:2, viewport:{width:1100,height:900}});
await pg.goto(new URL('img/diag.html', import.meta.url).href);
await (await pg.$('#logo')).screenshot({path:new URL('img/logo.png', import.meta.url).pathname});
await (await pg.$('#arq')).screenshot({path:new URL('img/arquitectura.png', import.meta.url).pathname});
await b.close();
