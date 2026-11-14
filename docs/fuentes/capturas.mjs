import { chromium } from 'playwright';
const b = await chromium.launch();
const pg = await b.newPage({locale:'es-CO',viewport:{width:390,height:780}, deviceScaleFactor:2});
const errs=[]; pg.on('pageerror',e=>errs.push(e.message)); pg.on('console',m=>{if(m.type()==='error')errs.push(m.text())});
const U = new URL('../../index.html', import.meta.url).href;
const sh = async (n,keep) => { await pg.waitForTimeout(250); if(!keep) await pg.evaluate(()=>document.querySelectorAll('.toast').forEach(t=>t.remove())); await pg.screenshot({path:new URL(`shots/${n}.png`, import.meta.url).pathname}); };
const go = async h => { await pg.evaluate(h=>location.hash=h,h); await pg.waitForTimeout(200); };
await pg.goto(U); await pg.evaluate(()=>localStorage.clear()); await pg.reload();
await sh('01_inicio');
// HU-01
await go('entrada');
await pg.fill('#f-prod','Leche entera 1 L (bolsa)'); await pg.dispatchEvent('#f-prod','input');
await pg.fill('#f-qty','24');
const v = await pg.evaluate(()=>dia(8)); await pg.fill('#f-venc', v);
await pg.click('text=Usar esa ubicación'); await pg.waitForTimeout(200);
await sh('02_entrada_form');
await pg.click('#b-guardar'); await sh('03_entrada_exito');
await go('inicio'); await go('entrada');
await pg.fill('#f-prod','Arroz 500 g'); await pg.dispatchEvent('#f-prod','input');
await pg.fill('#f-qty','0'); await pg.dispatchEvent('#f-qty','input');
await pg.fill('#f-venc', await pg.evaluate(()=>dia(-3))); await pg.dispatchEvent('#f-venc','input');
await pg.click('#b-guardar'); await pg.evaluate(()=>window.scrollTo(0,0)); await sh('04_entrada_errores');
await pg.evaluate(()=>{EN={};});
// HU-02
await go('buscar'); await pg.fill('#f-buscar','leche'); await pg.dispatchEvent('#f-buscar','input'); await sh('05_buscar');
await go('producto/p1'); await sh('06_producto');
await go('buscar'); await pg.fill('#f-buscar','mortadela'); await pg.dispatchEvent('#f-buscar','input'); await sh('07_buscar_noexiste');
await pg.fill('#f-buscar','harina'); await pg.dispatchEvent('#f-buscar','input'); await sh('08_buscar_agotado');
// HU-03
await go('vencen'); await sh('09_vencen');
await pg.click('text=Dar de baja >> nth=0'); await sh('10_baja_modal');
await pg.click('text=Sí, dar de baja'); await pg.waitForTimeout(400);
// HU-04
await go('vender'); await pg.fill('#f-vbuscar','kumis'); await pg.dispatchEvent('#f-vbuscar','input');
await pg.click('.add'); await sh('11_vender_vencido');
await pg.click('text=Entendido');
await pg.fill('#f-vbuscar','queso'); await pg.dispatchEvent('#f-vbuscar','input');
for(let i=0;i<5;i++) await pg.click('.add');
await sh('12_vender_sinstock',1);
await pg.waitForTimeout(3300);
await pg.fill('#f-vbuscar','arroz'); await pg.dispatchEvent('#f-vbuscar','input'); await pg.click('.add'); await pg.click('.add');
await pg.fill('#f-vbuscar',''); await pg.dispatchEvent('#f-vbuscar','input');
await sh('13_vender_carro');
await pg.click('.total-bar button'); await pg.fill('#f-recibe','40000'); await pg.dispatchEvent('#f-recibe','input'); await sh('14_cobrar_corto');
await pg.fill('#f-recibe','60000'); await pg.dispatchEvent('#f-recibe','input'); await sh('15_cobrar_ok');
await pg.click('#b-confirmar'); await sh('16_venta_ok');
// HU-05
await go('caja'); await sh('17_caja');
await pg.click('#b-cerrar'); await pg.evaluate(()=>window.scrollTo(0,document.body.scrollHeight)); await sh('18_caja_vacio');
const esp = await pg.evaluate(()=>D.base+D.ventas.reduce((s,v)=>s+v.total,0)-D.gastos.reduce((s,g)=>s+g.v,0));
await pg.fill('#f-contado', String(esp-15000)); await pg.dispatchEvent('#f-contado','input');
await pg.click('#b-cerrar'); await sh('19_caja_diferencia');
await pg.fill('#f-nota','Le fié a doña Carmen, paga el lunes'); await pg.click('text=Cerrar con nota'); await pg.waitForTimeout(300); await sh('20_caja_cerrada');
await go('vender'); await pg.click('.add'); await pg.click('.total-bar button'); await sh('21_venta_caja_cerrada');
await go('mapa'); await sh('22_mapa');
await go('faltantes'); await sh('23_faltantes');
console.log('ERRORES:', errs);
await b.close();
