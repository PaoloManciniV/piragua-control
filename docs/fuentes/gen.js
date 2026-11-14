const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, WidthType, BorderStyle, ShadingType, PageBreak, Footer, PageNumber,
  HeadingLevel, TabStopType, TabStopPosition, LeaderType, VerticalAlign,
  Bookmark, InternalHyperlink, ExternalHyperlink
} = require('docx');
const bid = t => { const m = t.match(/^(HU-\d+)/); if (m) return m[1].replace('-', '').toLowerCase(); return 's' + t.split(' ')[0].replace(/\.$/, '').replace(/\./g, '_'); };
const LINK = { color: '0563C1', underline: {} };

const PAGES = fs.existsSync('pages.json') ? JSON.parse(fs.readFileSync('pages.json')) : {};
const F = 'Arial';
const W = 9360;

// ---------- helpers ----------
function runs(text, base = {}) {
  // **negrita**, __cursiva__, [[texto|ancla]], referencias a secciones y URL
  const out = [];
  const parts = String(text).split(/(\*\*[^*]+\*\*|__[^_]+__|\[\[[^\]]+\]\]|secci(?:ón|ones) \d+(?:\.\d+)?(?: y \d+(?:\.\d+)?)?|https:\/\/(?:piragua-control\.netlify\.app|github\.com\/PaoloManciniV\/piragua-control|trello\.com\/b\/z9pio1m6))/);
  for (const p of parts) {
    if (!p) continue;
    if (p.startsWith('[[')) { const [t, a] = p.slice(2, -2).split('|'); out.push(new InternalHyperlink({ anchor: a, children: [new TextRun({ text: t, font: F, ...base, ...LINK })] })); continue; }
    if (p.startsWith('https://')) { out.push(new ExternalHyperlink({ link: p, children: [new TextRun({ text: p, font: F, ...base, ...LINK })] })); continue; }
    if (/^secci/.test(p)) {
      const [pal, ...resto] = p.split(' ');
      out.push(new TextRun({ text: pal + ' ', font: F, ...base }));
      resto.forEach(x => {
        if (x === 'y') out.push(new TextRun({ text: ' y ', font: F, ...base }));
        else out.push(new InternalHyperlink({ anchor: 's' + x.replace(/\./g, '_'), children: [new TextRun({ text: x, font: F, ...base, ...LINK })] }));
      });
      continue;
    }
    if (p.startsWith('**')) out.push(new TextRun({ text: p.slice(2, -2), bold: true, font: F, ...base }));
    else if (p.startsWith('__')) out.push(new TextRun({ text: p.slice(2, -2), italics: true, font: F, ...base }));
    else out.push(new TextRun({ text: p, font: F, ...base }));
  }
  return out;
}
const P = (t, o = {}) => new Paragraph({
  children: runs(t, { size: o.size || 22, color: o.color }),
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 160, before: o.before ?? 0, line: 276 },
  indent: o.indent, keepNext: o.keepNext,
});
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new Bookmark({ id: bid(t), children: [new TextRun({ text: t, bold: true, font: F, size: 28, color: '000000' })] })], spacing: { before: 0, after: 200 }, keepNext: true });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new Bookmark({ id: bid(t), children: [new TextRun({ text: t, bold: true, font: F, size: 24, color: '000000' })] })], spacing: { before: 240, after: 120 }, keepNext: true });
const H3 = t => new Paragraph({ children: [new TextRun({ text: t, bold: true, font: F, size: 22, color: '0F5C6E' })], spacing: { before: 180, after: 100 }, keepNext: true });
const BR = () => new Paragraph({ children: [new PageBreak()] });
const bullet = (t, lvl = 0) => new Paragraph({ children: runs(t, { size: 22 }), numbering: { reference: 'vin', level: lvl }, spacing: { after: 80, line: 276 }, alignment: AlignmentType.JUSTIFIED });
const num = (t, ref) => new Paragraph({ children: runs(t, { size: 22 }), numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 276 }, alignment: AlignmentType.JUSTIFIED });
let numRefs = [];
function lista(items) { const ref = 'n' + numRefs.length; numRefs.push(ref); return items.map(t => num(t, ref)); }
let nTabla = 0, nFig = 0;
const cap = (t, kind) => new Paragraph({ children: [new TextRun({ text: `${kind === 'f' ? 'Figura ' + (++nFig) : 'Tabla ' + (++nTabla)}. ${t}`, font: F, size: 18, italics: false })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 220 } });

const bd = { style: BorderStyle.SINGLE, size: 4, color: '9AA5AA' };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
function cell(t, w, o = {}) {
  const paras = Array.isArray(t) ? t : String(t).split('\n');
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: o.head ? { fill: 'D9E4E8', type: ShadingType.CLEAR, color: 'auto' } : (o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined),
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: paras.map(x => x instanceof Paragraph ? x : new Paragraph({
      children: runs(x, { size: o.size || 19, bold: o.head ? true : undefined }),
      alignment: o.align || (o.head ? AlignmentType.CENTER : AlignmentType.LEFT), spacing: { after: 20 }
    }))
  });
}
function tabla(headers, rows, widths, o = {}) {
  const tot = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (headers) trs.push(new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => cell(h, widths[i], { head: true, size: o.size })) }));
  rows.forEach(r => trs.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { size: o.size, align: (o.center || []).includes(i) ? AlignmentType.CENTER : undefined, fill: o.fills ? o.fills(r, i) : undefined })) })));
  return new Table({ width: { size: tot, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
function img(path, wpx, hpx) {
  return new ImageRun({ type: 'png', data: fs.readFileSync(path), transformation: { width: wpx, height: hpx } });
}
function figura(path, wpx, hpx, t) {
  return [new Paragraph({ children: [img(path, wpx, hpx)], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true }), cap(t, 'f')];
}
// cuadricula de capturas: 3 por fila
function capturas(list) {
  const out = [];
  const nb = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const nbs = { top: nb, bottom: nb, left: nb, right: nb };
  for (let i = 0; i < list.length; i += 3) {
    const fila = list.slice(i, i + 3);
    while (fila.length < 3) fila.push(null);
    out.push(new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [3120, 3120, 3120],
      rows: [new TableRow({ cantSplit: true, children: fila.map(f => new TableCell({
        borders: nbs, width: { size: 3120, type: WidthType.DXA }, margins: { top: 40, bottom: 40, left: 60, right: 60 },
        children: f ? [
          new Paragraph({ children: [img('shots/' + f[0] + '.png', 116, 232)], alignment: AlignmentType.CENTER }),
          new Paragraph({ children: [new TextRun({ text: `Figura ${++nFig}. ${f[1]}`, font: F, size: 16 })], alignment: AlignmentType.CENTER, spacing: { before: 30, after: 60 } })
        ] : [new Paragraph('')]
      })) })]
    }));
  }
  return out;
}

// ---------- datos ----------
const EQUIPO = ['Andrés Parra', 'Dylan Mariño', 'Paolo Mancini Villalobos', 'Juan Pablo Mora Gómez', 'Milan Libardo Mejía Reales', 'Juan Esteban Quintero Gelvez'];

const BACKLOG = [
  ['HU-01', 'Como **dueño de la tienda**, quiero registrar la mercancía que llega con su cantidad, costo, fecha de vencimiento y el lugar donde la guardo, para saber qué tengo y dónde quedó.', 'Alta', 5],
  ['HU-02', 'Como **ayudante de mostrador**, quiero buscar un producto y ver en qué zona, estante y nivel está guardado y cuántas unidades hay, para encontrarlo sin revolver la bodega.', 'Alta', 3],
  ['HU-03', 'Como **dueño de la tienda**, quiero ver qué productos están vencidos o vencen en los próximos días, con su ubicación, para retirarlos o ponerlos adelante antes de perder la plata.', 'Alta', 3],
  ['HU-04', 'Como **ayudante de mostrador**, quiero registrar cada venta y que se descuenten las unidades del inventario, para que las existencias estén al día sin contar a mano.', 'Alta', 5],
  ['HU-05', 'Como **dueño de la tienda**, quiero hacer el cierre de caja del día y ver cuánto vendí, cuánto me costó, cuánto gasté y cuánto gané, para saber si el negocio está dejando plata.', 'Alta', 5],
  ['HU-06', 'Como **dueño de la tienda**, quiero una lista de los productos que están por debajo del mínimo con la cantidad sugerida para pedir, para no quedarme sin lo que más se vende.', 'Media', 3],
  ['HU-07', 'Como **dueño de la tienda**, quiero dar de baja los productos vencidos o dañados, para que la pérdida quede registrada y se descuente del inventario.', 'Media', 2],
  ['HU-08', 'Como **dueño de la tienda**, quiero definir las zonas, estantes y niveles de la tienda y la bodega, para que cada producto tenga un lugar fijo donde guardarse.', 'Media', 3],
  ['HU-09', 'Como **dueño de la tienda**, quiero registrar los gastos del día (servicios, bolsas, transporte), para que el cierre muestre la ganancia real.', 'Media', 2],
  ['HU-10', 'Como **dueño de la tienda**, quiero ver un resumen del mes con ventas, gastos, pérdidas y ganancia, para decidir qué comprar y qué dejar de traer.', 'Media', 5],
  ['HU-11', 'Como **dueño de la tienda**, quiero hacer un conteo físico por estante y ajustar las existencias, para corregir las diferencias entre lo que dice la app y lo que hay.', 'Baja', 5],
  ['HU-12', 'Como **dueño de la tienda**, quiero guardar los datos de mis proveedores y armar el pedido desde la lista de faltantes, para no olvidar pedir nada.', 'Baja', 5],
  ['HU-13', 'Como **dueño de la tienda**, quiero que el ayudante entre con un PIN que solo le permita vender y buscar, para proteger la información de costos y ganancias.', 'Baja', 3],
];
const totalSP = BACKLOG.reduce((s, r) => s + r[3], 0);

// estimaciones IA vs equipo
const EST = [
  ['HU-01', 3, 'Formulario de registro con pocos campos.', 5, 'Además del formulario hay lotes por fecha, selección de zona, estante y nivel, sugerencia de la ubicación donde ya está el producto y validación de fechas vencidas.'],
  ['HU-02', 5, 'Búsqueda más un mapa visual de la tienda.', 3, 'El mapa es una cuadrícula sencilla de zonas y estantes, no un plano dibujado. La búsqueda es local, sin servidor.'],
  ['HU-03', 3, 'Consulta por fecha y lista con colores.', 3, 'Coinciden.'],
  ['HU-04', 3, 'Carrito de compras básico.', 5, 'Hay que descontar primero el lote que vence antes, bloquear la venta de vencidos, calcular el cambio y avisar cuando el producto queda bajo.'],
  ['HU-05', 8, 'Reporte contable con varias fuentes.', 5, 'No es contabilidad formal: suma datos que ya están guardados en el día. Se bajó a 5 tras la segunda ronda.'],
  ['HU-06', 2, 'Filtro de productos bajo el mínimo.', 3, 'Incluye la cantidad sugerida para pedir y debe actualizarse con cada venta.'],
  ['HU-07', 2, 'Botón para descontar unidades.', 2, 'Coinciden.'],
  ['HU-08', 5, 'Editor de planos de la tienda.', 3, 'El dueño no necesita dibujar la tienda: basta con una lista de zonas, estantes y niveles.'],
  ['HU-09', 1, 'Formulario de un solo registro.', 2, 'El gasto tiene que reflejarse en el cierre del día y en el resumen mensual.'],
  ['HU-10', 5, 'Agrupación mensual y gráficas.', 5, 'Coinciden.'],
  ['HU-11', 3, 'Edición de cantidades.', 5, 'El conteo es por estante, guarda el motivo de cada ajuste y debe funcionar sin internet.'],
  ['HU-12', 8, 'Integración con WhatsApp para enviar el pedido.', 5, 'Se quitó la integración: el pedido se arma y se comparte con el botón de compartir del celular.'],
  ['HU-13', 3, 'Inicio de sesión con roles.', 3, 'Coinciden.'],
];
const totalIA = EST.reduce((s, r) => s + r[1], 0);

// costo
const HORAS_SP = 4;
const horasTot = totalSP * HORAS_SP;
const ROLES = [
  ['Desarrollador frontend (pantallas, PWA, reglas en el celular)', 88, 30000],
  ['Desarrollador backend (Supabase, base de datos, sincronización)', 52, 35000],
  ['Diseñador UX/UI (prototipo, pruebas con el dueño)', 24, 30000],
  ['Pruebas (QA) de flujos felices y alternos', 20, 25000],
  ['Scrum Master y Product Owner (ceremonias, backlog)', 12, 35000],
];
const sumH = ROLES.reduce((s, r) => s + r[1], 0);
if (sumH !== horasTot) throw new Error('horas no cuadran ' + sumH + ' ' + horasTot);
const costoRoles = ROLES.reduce((s, r) => s + r[1] * r[2], 0);
const HERR = [
  ['Netlify (hosting de la PWA, plan gratis)', 0],
  ['Supabase (API y PostgreSQL, plan gratis hasta 500 MB)', 0],
  ['GitHub y Trello (plan gratis)', 0],
  ['Dominio piraguacontrol.com.co por un año', 70000],
  ['Suscripción de IA (Claude Pro) durante el desarrollo, 2 meses aprox.', 170000],
];
const costoHerr = HERR.reduce((s, r) => s + r[1], 0);
const subtotal = costoRoles + costoHerr;
const conting = Math.round(subtotal * 0.10);
const total = subtotal + conting;
const $ = n => '$ ' + n.toLocaleString('es-CO');

// ---------- contenido ----------
const ch = [];

// PORTADA
const centro = (t, s, b, after = 120, color) => new Paragraph({ children: [new TextRun({ text: t, font: F, size: s, bold: b, color })], alignment: AlignmentType.CENTER, spacing: { after } });
ch.push(new Paragraph({ spacing: { before: 600 } }));
ch.push(centro('UNIVERSIDAD SIMÓN BOLÍVAR', 26, true, 60));
ch.push(centro('Facultad de Ingenierías', 22, false, 60));
ch.push(centro('Programa de Ingeniería de Sistemas', 22, false, 700));
ch.push(new Paragraph({ children: [img('img/logo.png', 390, 105)], alignment: AlignmentType.CENTER, spacing: { after: 200 } }));
ch.push(centro('Aplicación de inventario, ubicación de mercancía y caja', 22, false, 40));
ch.push(centro('para la tienda La Piragua (Barranquilla)', 22, false, 600));
ch.push(centro('Proyecto final: IA en el ciclo de proyectos de software', 24, true, 60));
ch.push(centro('Documento de entrega con prototipo navegable', 22, false, 600));
ch.push(centro('Presentado por:', 22, true, 60));
EQUIPO.forEach(n => ch.push(centro(n, 22, false, 40)));
ch.push(new Paragraph({ spacing: { after: 200 } }));
ch.push(new Paragraph({ children: runs('**Presentado a:** Rafael Pérez Meriño', { size: 22 }), alignment: AlignmentType.CENTER, spacing: { after: 40 } }));
ch.push(new Paragraph({ children: runs('**Asignatura:** Gestión de Proyectos (5925)', { size: 22 }), alignment: AlignmentType.CENTER, spacing: { after: 40 } }));
ch.push(new Paragraph({ children: runs('**Unidades:** 3, IA en el ciclo de proyectos de software, y 4, Metodologías ágiles con IA', { size: 22 }), alignment: AlignmentType.CENTER, spacing: { after: 40 } }));
ch.push(new Paragraph({ children: runs('**Fecha de entrega:** 16 de noviembre de 2026', { size: 22 }), alignment: AlignmentType.CENTER, spacing: { after: 500 } }));
ch.push(centro('Barranquilla, noviembre de 2026', 22, false, 0));
ch.push(BR());

// CONTENIDO
const TOC = [
  [0, '2. Planteamiento del problema'], [1, '2.1 Situación actual'], [1, '2.2 Problema específico'], [1, '2.3 Por qué es importante resolverlo'], [1, '2.4 Quién tiene el problema'], [1, '2.5 Solución propuesta'],
  [0, '3. Objetivos'], [1, '3.1 Objetivo general'], [1, '3.2 Objetivos específicos'],
  [0, '4. Información del proyecto e investigación'], [1, '4.1 Perfil de la persona entrevistada'], [1, '4.2 Resumen de entrevistas'], [1, '4.3 Hallazgos clave y citas'], [1, '4.4 Resultados del cuestionario'], [1, '4.5 Problem statement'],
  [0, '5. Backlog del producto'], [1, '5.1 Backlog priorizado'], [1, '5.2 Proceso con IA: qué propuso la IA y qué ajustó el equipo'], [1, '5.3 Fuera del alcance'],
  [0, '6. Historias de usuario desarrolladas: top 5'],
  [0, '7. Estimaciones'], [1, '7.1 Estimación en story points'], [1, '7.2 Conversión a tiempo'], [1, '7.3 Estimación de costo'], [1, '7.4 Cronograma de sprints'],
  [0, '8. Stack tecnológico'], [1, '8.1 Tecnología por capa'], [1, '8.2 Diagrama de arquitectura'],
  [0, '9. Link del prototipo'], [1, '9.1 Enlace y cómo recorrerlo'], [1, '9.2 Flujos alternos visibles'], [1, '9.3 Revisión con el dueño de la tienda'], [1, '9.4 Capturas de los flujos principales'], [1, '9.5 Enlaces de la entrega'],
];
ch.push(new Paragraph({ children: [new TextRun({ text: 'Contenido', bold: true, font: F, size: 28 })], spacing: { after: 240 } }));
TOC.forEach(([lvl, t]) => ch.push(new Paragraph({
  children: [new InternalHyperlink({ anchor: bid(t), children: [new TextRun({ text: t, font: F, size: 22, bold: lvl === 0 }), new TextRun({ text: '\t' + (PAGES[t] || ''), font: F, size: 22 })] })],
  tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX, leader: LeaderType.DOT }],
  indent: { left: lvl ? 440 : 0 }, spacing: { after: 70, before: lvl ? 0 : 90 }
})));
ch.push(BR());

// 2. PLANTEAMIENTO
ch.push(H1('2. Planteamiento del problema'));
ch.push(H2('2.1 Situación actual'));
ch.push(P('La Piragua es una tienda de barrio ubicada en la Calle 11 #14-06 de Barranquilla. Su dueño, Alfredo Ortega, vende productos de la canasta familiar: lácteos, granos, abarrotes, bebidas, panadería y aseo. Maneja alrededor de 300 productos distintos repartidos entre el mostrador, un pasillo de estantes, la nevera y una bodega pequeña en la parte de atrás.'));
ch.push(P('Su día empieza a las siete de la mañana. Revisa qué hizo falta, recibe los pedidos de los proveedores, atiende todo el día y en la noche cuenta la plata de la caja. Todo el control se lleva a mano: lo que entra lo anota en un cuaderno "cuando me acuerdo", lo que sale no lo registra y las facturas de compra quedan guardadas en una bolsa. No usa ningún sistema ni calculadora para el cierre.'));
ch.push(P('El punto más crítico es la forma de guardar la mercancía. En la visita a la tienda vimos los productos repartidos sin un lugar fijo y sin orden por fecha de vencimiento: un mismo producto puede estar en el estante y en la bodega a la vez, y lo nuevo queda adelante de lo viejo. Uno de los integrantes del equipo, que conoce la tienda desde antes, lo resumió así: la tienda es un desastre en la logística de guardar. Como nadie sabe con certeza qué hay ni dónde está, el dueño depende de su memoria y de ir a mirar.'));
ch.push(H2('2.2 Problema específico'));
ch.push(P('**Alfredo Ortega, dueño de la tienda La Piragua en Barranquilla, pierde entre 150.000 y 300.000 pesos al mes por productos que se vencen o se dañan y no sabe cuánto gana realmente, porque guarda la mercancía sin una ubicación definida ni orden por fecha de vencimiento y lleva el inventario y las cuentas en cuadernos que no le muestran qué tiene, dónde está ni cuándo vence.**'));
ch.push(P('El enunciado se puede verificar con el dueño: el rango de pérdida lo marcó él mismo en el cuestionario (ítem 6), la falta de orden se observó en la visita y la dificultad para conocer la ganancia la confirmó en la entrevista (pregunta 8) y en la escala del ítem 8, donde marcó 4 de 5.'));
ch.push(H2('2.3 Por qué es importante resolverlo'));
[
  '**Pérdida de dinero directa.** Con pérdidas de 150.000 a 300.000 pesos al mes, la tienda bota entre 1,8 y 3,6 millones de pesos al año en mercancía vencida o dañada. Para un negocio de barrio eso puede ser la ganancia de varias semanas.',
  '**Pérdida de clientes.** El dueño contó que a veces se entera de que un producto está vencido porque se lo dice el cliente, y lo describió como "lo peor". Un cliente que se lleva algo vencido puede no volver.',
  '**Compras a ciegas.** Como no sabe cuánto tiene, a veces pide de más lo que no rota y se queda sin lo que más se vende. Lo primero inmoviliza plata y lo segundo hace perder ventas.',
  '**Tiempo perdido.** Para saber si hay un producto tiene que ir a mirar al estante o a la bodega. Ese tiempo sale de la atención al público.',
  '**Decisiones sin datos.** Se guía por lo que hay en la caja al final del día, pero esa plata no descuenta lo que pagó a proveedores ni los gastos. No sabe si el negocio está ganando o solo moviendo plata.',
].forEach(t => ch.push(bullet(t)));
ch.push(H2('2.4 Quién tiene el problema'));
ch.push(P('Alfredo Ortega, propietario y administrador de La Piragua. Es quien compra, recibe la mercancía, atiende la caja, paga a los proveedores y lleva las cuentas. Usa el celular para tener los números de los proveedores, pero no maneja herramientas digitales para el negocio y él mismo dice que no es "muy de tecnología". El perfil completo está en la sección 4.1.'));
ch.push(H2('2.5 Solución propuesta'));
ch.push(P('**Piragua Control**, una aplicación web para el celular que registra la mercancía con su ubicación y fecha de vencimiento, avisa qué se va a vencer o se está acabando, descuenta las ventas del inventario y muestra en el cierre del día cuánto se ganó, incluso sin conexión a internet.'));
ch.push(BR());

// 3. OBJETIVOS
ch.push(H1('3. Objetivos'));
ch.push(H2('3.1 Objetivo general'));
ch.push(P('Desarrollar una aplicación web progresiva para Alfredo Ortega, dueño de la tienda La Piragua, que permita registrar la mercancía con su ubicación y fecha de vencimiento, controlar las ventas y conocer la ganancia de cada día, mediante tecnologías web que funcionen en el celular aun sin conexión a internet.'));
ch.push(H2('3.2 Objetivos específicos'));
lista([
  '**Identificar**, mediante entrevista semiestructurada, cuestionario y observación directa con el dueño, cómo se guarda, se registra y se vende hoy la mercancía en La Piragua y dónde se originan las mayores pérdidas.',
  `**Diseñar** el backlog del producto con ${BACKLOG.length} historias de usuario priorizadas por valor de negocio y desarrollar las 5 más prioritarias con criterios DADO, CUANDO, ENTONCES, flujo feliz y flujos alternos.`,
  '**Construir** un prototipo navegable que cubra los flujos felices de las 5 historias prioritarias y muestre por lo menos 2 flujos alternos, y revisarlo con el dueño de la tienda.',
  '**Estimar** el esfuerzo, el tiempo y el costo del desarrollo usando story points, planning poker y la capacidad real del equipo, organizados en sprints de dos semanas.',
  '**Definir** el stack tecnológico por capa, justificando cada decisión con las condiciones de la tienda: uso en el celular, conexión a internet no garantizada y presupuesto bajo.',
]).forEach(p => ch.push(p));
ch.push(P('Cada objetivo tiene un entregable que se puede revisar en este documento: el primero en la sección 4, el segundo en las secciones 5 y 6, el tercero en la sección 9, el cuarto en la sección 7 y el quinto en la sección 8.', { before: 120 }));
ch.push(BR());

// 4. INVESTIGACIÓN
ch.push(H1('4. Información del proyecto e investigación'));
ch.push(H2('4.1 Perfil de la persona entrevistada'));
ch.push(tabla(null, [
  ['**Nombre**', 'Alfredo Ortega'],
  ['**Rol**', 'Propietario y administrador de la tienda La Piragua'],
  ['**Negocio**', 'Tienda de barrio de productos de la canasta familiar. Unos 300 productos distintos en mostrador, pasillo de estantes, nevera y una bodega pequeña atrás.'],
  ['**Ubicación**', 'Calle 11 #14-06, Barranquilla, Atlántico'],
  ['**Rutina relacionada con el problema**', 'Abre a las 7:00 a. m., revisa qué falta, recibe pedidos en la mañana, atiende todo el día y cuenta el efectivo de la caja en la noche. Revisa el inventario completo "casi nunca", solo cuando ve que algo se está acabando.'],
  ['**Cómo lleva el control hoy**', 'Dos cuadernos (entradas y gastos), una bolsa con facturas de compra y los números de los proveedores en el celular y en una libreta.'],
  ['**Relación con la tecnología**', 'Usa el celular a diario, pero no tiene herramientas digitales para el negocio. Pide algo "fácil de usar" y "que no sea complicado".'],
  ['**Cómo llegamos a él**', 'Un integrante del equipo conoce la tienda y nos puso en contacto con el dueño.'],
], [2600, 6760]));
ch.push(cap('Perfil del entrevistado.'));
ch.push(H2('4.2 Resumen de entrevistas'));
ch.push(P('Aplicamos el kit de elicitación diseñado en el Taller 1, ya ajustado después de la prueba piloto cruzada. La entrevista se hizo en la tienda para poder combinarla con la observación directa de los estantes y la bodega.'));
ch.push(tabla(['#', 'Fecha', 'Medio', 'Duración', 'Técnica y alcance', 'Participantes'], [
  ['1', '7 nov 2026', 'Presencial, en la tienda', '45 min', 'Preguntas filtro, entrevista semiestructurada con el guion en embudo (11 preguntas en 5 etapas) y cuestionario de 10 ítems.', 'Alfredo Ortega. Entrevistador: Juan Esteban Quintero. Notas: Paolo Mancini.'],
  ['2', '7 nov 2026', 'Presencial, en la tienda', '20 min', 'Observación directa del mostrador, la nevera y la bodega, y de cómo se registra una entrada de mercancía.', 'Dylan Mariño y Milan Mejía.'],
  ['3', '15 nov 2026', 'Presencial, en la tienda, con el celular del dueño', '30 min', 'Revisión del prototipo navegable: recorrido de los 5 flujos y de los mensajes de error (detalle en la sección 9.3).', 'Alfredo Ortega. Andrés Parra (PO) y Juan Pablo Mora.'],
], [400, 1050, 1350, 1000, 3160, 2400], { center: [0, 3] }));
ch.push(cap('Registro de las sesiones con el dueño de La Piragua.'));
ch.push(H2('4.3 Hallazgos clave y citas'));
ch.push(P('Estas son las respuestas del dueño que más pesaron en las decisiones del producto. Las citas son textuales, tomadas de las notas de la entrevista del 7 de noviembre.'));
ch.push(tabla(['Hallazgo', 'Cita textual del dueño', 'Decisión de producto'], [
  ['Las salidas de mercancía no se registran.', '"Lo que sale no lo registro, me doy cuenta cuando veo el estante vacío."', 'Cada venta descuenta del inventario (HU-04) y alimenta la lista por pedir (HU-06).'],
  ['No sabe cuánto tiene ni dónde está.', '"No sé cuánto tengo de cada cosa. Toca ir a mirar al estante o a la bodega."', 'Pantalla "¿Dónde está?" con zona, estante y nivel (HU-02) y mapa de la tienda (HU-08).'],
  ['La mercancía se guarda sin un lugar fijo.', '"Lo que entra lo anoto en un cuaderno cuando me acuerdo."', 'La entrada no se guarda sin ubicación y sugiere el sitio donde ya está el producto (HU-01).'],
  ['Los vencidos los descubre el cliente.', '"A veces me entero cuando el cliente me dice que un producto está vencido, y eso es lo peor."', 'Alertas de vencimiento por colores y bloqueo de la venta de vencidos (HU-03 y HU-04).'],
  ['Las pérdidas se concentran en perecederos.', '"Lo que se vence, sobre todo lácteos y granos, y lo que se daña."', 'La nevera es una zona propia en el mapa y la alerta empieza 15 días antes.'],
  ['No sabe cuánto gana.', '"La verdad no sé exacto. [...] Me guío por lo que hay en la caja."', 'Cierre de caja con costo de lo vendido, gastos, pérdidas y ganancia del día (HU-05 y HU-09).'],
  ['Pide sin orden y se le olvidan cosas.', '"Pido cuando veo que falta, no tengo un orden fijo y a veces se me olvida pedir algo."', 'Lista por pedir con cantidad sugerida (HU-06) y pedido a proveedores (HU-12).'],
  ['Su prioridad es el inventario y los vencimientos.', '"Primero lo del inventario y los vencimientos, porque ahí es donde más plata pierdo."', 'Esas historias quedaron en los primeros lugares del backlog.'],
  ['La herramienta debe ser simple.', '"Que no sea complicado, porque no soy muy de tecnología."', 'Botones grandes, letra grande, pocas pantallas y mensajes en lenguaje de la tienda.'],
], [2300, 3560, 3500], { size: 18 }));
ch.push(cap('Hallazgos, evidencia y decisión que tomó el equipo.'));
ch.push(H3('Observación directa'));
ch.push(P('La mercancía está en estantes y en una bodega pequeña atrás, sin orden por fecha de vencimiento. Un mismo producto aparece en dos o tres lugares y lo que llega nuevo se pone adelante, así que lo viejo se queda atrás hasta vencerse. Las cuentas están en dos cuadernos y una bolsa con facturas. El cierre se hace contando solo el efectivo. Se nota que el negocio depende de la memoria del dueño.'));
ch.push(H2('4.4 Resultados del cuestionario'));
ch.push(tabla(['#', 'Ítem', 'Respuesta del dueño'], [
  ['1', 'Rol en La Piragua', 'Propietario'],
  ['2', 'Frecuencia con que revisa el inventario', 'Rara vez'],
  ['3', 'Llevar el inventario en papel me genera errores o pérdidas de tiempo (1 a 5)', '5, totalmente de acuerdo'],
  ['4', 'Productos distintos que maneja', 'Alrededor de 300'],
  ['5', 'Frecuencia con que encuentra productos vencidos', 'A veces'],
  ['6', 'Pérdida mensual por vencidos o dañados', 'Entre 150.000 y 300.000 pesos'],
  ['7', 'Controles que lleva en papel', 'Inventario, gastos y proveedores. No controla vencimientos ni ventas.'],
  ['8', 'Me cuesta saber cuánto gano al final del mes (1 a 5)', '4, de acuerdo'],
  ['9', '¿Usaría una herramienta digital sencilla?', 'Sí'],
  ['10', '¿Qué esperaría de esa herramienta?', '"Algo fácil de usar, que me avise qué se va a vencer y qué se está acabando, y que me diga cuánto gané."'],
], [500, 4600, 4260], { center: [0] }));
ch.push(cap('Respuestas del cuestionario estructurado (Google Forms).'));
ch.push(P('El ítem 7 confirma el vacío principal: hoy nadie controla los vencimientos ni las ventas, que son justo las dos cosas que más le cuestan plata.'));
ch.push(H2('4.5 Problem statement'));
ch.push(P('**Alfredo Ortega, dueño de La Piragua, pierde entre 150.000 y 300.000 pesos al mes en productos vencidos o dañados y no sabe cuánto gana, porque no tiene un registro que le diga qué mercancía tiene, dónde está guardada y cuándo se vence.**'));
ch.push(BR());

// 5. BACKLOG
ch.push(H1('5. Backlog del producto'));
ch.push(P(`El backlog tiene ${BACKLOG.length} historias de usuario ordenadas de mayor a menor valor de negocio. El valor lo definió el Product Owner (Andrés Parra) con base en lo que dijo el dueño: primero inventario, ubicación y vencimientos, después ventas y caja, y al final lo administrativo. Los story points son la estimación final del equipo en planning poker (sección 7.1). Ninguna historia usa el rol genérico "usuario": los roles son el dueño de la tienda y el ayudante de mostrador.`));
ch.push(H2('5.1 Backlog priorizado'));
ch.push(tabla(['ID', 'Historia de usuario', 'Prioridad', 'SP'],
  BACKLOG.map((r, i) => [i < 5 ? `[[${r[0]}|${r[0].replace('-', '').toLowerCase()}]]` : r[0], r[1], r[2], String(r[3])]).concat([['', '**Total**', '', `**${totalSP}**`]]),
  [900, 6560, 1100, 800], { center: [0, 2, 3], size: 18, fills: (r, i) => i === 2 ? (r[2] === 'Alta' ? 'FDE3D2' : r[2] === 'Media' ? 'FFF4D1' : r[2] === 'Baja' ? 'E9F1F3' : undefined) : undefined }));
ch.push(cap('Backlog del producto Piragua Control. Las cinco primeras se desarrollan en la sección 6.'));
ch.push(H2('5.2 Proceso con IA: qué propuso la IA y qué ajustó el equipo'));
ch.push(P('Usamos Claude (Anthropic) como asistente. Le dimos el problem statement, las citas de la entrevista y los resultados del cuestionario, y le pedimos un borrador de historias de usuario con el formato "Como, quiero, para". La IA propuso 12 historias. El equipo las revisó en una sesión de una hora contra las notas de la entrevista y decidió qué se quedaba, qué se cambiaba y qué se agregaba. La regla fue la del curso: la IA propone, el equipo revisa y decide.'));
ch.push(tabla(['Lo que propuso la IA', 'Lo que decidió el equipo', 'Por qué'], [
  ['Todas las historias con el rol "usuario" o "administrador".', 'Se reescribieron con roles reales: dueño de la tienda y ayudante de mostrador.', 'La guía no acepta el rol genérico y en la tienda hay dos personas con necesidades distintas.'],
  ['No incluyó nada sobre dónde se guarda la mercancía.', 'Se agregaron HU-02 (¿Dónde está?) y HU-08 (mapa de zonas, estantes y niveles), y la ubicación se volvió obligatoria en HU-01.', 'Es el problema que más vimos en la observación y el que más comentó el equipo que conoce la tienda. La IA no lo podía saber porque no estuvo allí.'],
  ['Registro de productos con lector de código de barras usando la cámara.', 'Se descartó para esta versión. La búsqueda es por nombre.', 'Muchos productos se venden sueltos o vienen sin código, y el dueño pidió algo simple.'],
  ['Unir ventas y cierre de caja en una sola historia.', 'Se separaron en HU-04 y HU-05.', 'Las hace en momentos distintos del día y la historia unida quedaba muy grande para un sprint.'],
  ['Facturación electrónica ante la DIAN.', 'Fuera del alcance (sección 5.3).', 'El dueño no la mencionó y no resuelve las pérdidas por vencimiento.'],
  ['Inicio de sesión con correo y contraseña.', 'Se cambió por un PIN para el ayudante (HU-13) y quedó con prioridad baja.', 'En la tienda se comparte un solo celular. Un correo y una contraseña estorban al vender.'],
  ['Alertas de vencimiento 30 días antes.', 'Se ajustó a 15 días, en tres grupos de colores.', 'Con 30 días casi todo aparecía en la lista y las alertas dejaban de servir. Los lácteos duran menos de 15 días.'],
  ['No mencionó qué hacer con lo vencido.', 'Se agregó HU-07 (dar de baja con registro de la pérdida).', 'Sin esa historia la pérdida mensual seguiría siendo invisible, que es parte del problema.'],
], [2900, 3300, 3160], { size: 18 }));
ch.push(cap('Bitácora del Taller 3: propuesta de la IA frente a la decisión del equipo.'));
ch.push(H2('5.3 Fuera del alcance'));
ch.push(P('Quedan por fuera de esta versión la facturación electrónica ante la DIAN, el pago con Nequi o Daviplata integrado, el lector de código de barras y el manejo de varias tiendas. Se podrán revisar con el dueño cuando el producto esté en uso.'));
ch.push(BR());

// 6. TOP 5
ch.push(H1('6. Historias de usuario desarrolladas: top 5'));
ch.push(P('Estas son las cinco historias de prioridad alta. Cada una tiene su narrativa, sus criterios de aceptación con el formato DADO, CUANDO, ENTONCES, el flujo feliz tal como se recorre en el prototipo y sus flujos alternos con el mensaje exacto que ve la persona. Los flujos alternos marcados con (P) están construidos en el prototipo y se pueden probar.'));

const HUS = [
  {
    id: 'HU-01', t: 'Registrar la entrada de mercancía con su ubicación', sp: 5,
    narr: BACKLOG[0][1],
    crit: [
      ['el dueño está en "Entrada de mercancía"', 'escribe un producto que ya está guardado en la tienda', 'la app le muestra en qué zona, estante y nivel está y le ofrece el botón "Usar esa ubicación".'],
      ['llenó producto, cantidad, costo, fecha de vencimiento y ubicación', 'toca "Guardar entrada"', 've la pantalla "Entrada guardada" con la cantidad, la ubicación, la fecha de vencimiento y el total de unidades que hay ahora en la tienda.'],
      ['el producto es de aseo o no se vence', 'marca "Este producto no se vence"', 'el campo de fecha se desactiva y la entrada se guarda sin fecha de vencimiento.'],
      ['no eligió zona, estante o nivel', 'toca "Guardar entrada"', 'la entrada no se guarda y aparece en rojo el mensaje de ubicación obligatoria.'],
    ],
    feliz: [
      'Desde Inicio toca "Entrada de mercancía".',
      'Escribe "leche" y elige "Leche entera 1 L (bolsa)". La app llena el costo y muestra que ya está guardada en Nevera, estante N1, nivel 2.',
      'Escribe la cantidad que llegó (24) y la fecha de vencimiento que trae el producto.',
      'Toca "Usar esa ubicación" para guardarla en el mismo sitio.',
      'Toca "Guardar entrada".',
      'Ve "Entrada guardada" con el resumen y elige "Registrar otra entrada" o "Ir al inicio".',
    ],
    sad: [
      ['La cantidad está vacía o en cero. (P)', '"Escriba una cantidad mayor que cero." La entrada no se guarda.'],
      ['La fecha de vencimiento ya pasó. (P)', '"La fecha de vencimiento (11 nov) ya pasó. Ese producto no se debe recibir: devuélvaselo al proveedor."'],
      ['No eligió dónde guardar el producto. (P)', '"Elija dónde va a guardar el producto. Sin ubicación no se puede guardar la entrada."'],
      ['El producto no existe en la lista. (P)', '"\'Mortadela\' no está en la lista de productos. Revise cómo lo escribió o créelo primero como producto nuevo."'],
      ['El costo es igual o mayor que el precio de venta. (P)', 'Ventana "El costo es igual o mayor que el precio de venta" que explica que pierde plata en cada venta, con los botones "Revisar costo" y "Guardar igual".'],
    ],
  },
  {
    id: 'HU-02', t: 'Ubicar un producto en la tienda', sp: 3,
    narr: BACKLOG[1][1],
    crit: [
      ['el ayudante está en "¿Dónde está?"', 'escribe al menos dos letras del producto', 've los productos que coinciden con sus unidades y una etiqueta con estante, nivel y zona.'],
      ['aparece el producto en los resultados', 'lo toca', 've cada lote con su ubicación completa y su fecha de vencimiento, y el estante resaltado en naranja en el mapa de la tienda.'],
      ['el producto tiene lotes en dos lugares', 'abre el detalle', 'los lotes aparecen ordenados por fecha de vencimiento, primero el que vence antes.'],
      ['el producto no tiene unidades', 'aparece en la búsqueda', 'se muestra la etiqueta "Agotado" y el botón "Ver lista por pedir".'],
    ],
    feliz: [
      'Toca "¿Dónde está?" en la barra inferior.',
      'Escribe "leche".',
      'Ve "Leche entera 1 L (bolsa)", 48 unidades, con la etiqueta "N1 nivel 2 · Nevera".',
      'Toca el resultado.',
      'Ve los lotes ordenados por vencimiento y el estante N1 resaltado en el mapa.',
      'Va a la nevera y saca primero el lote que vence antes.',
    ],
    sad: [
      ['Escribe menos de dos letras. (P)', '"Escriba al menos 2 letras para buscar."'],
      ['El producto no existe. (P)', '"No encontramos \'mortadela\'. Revise cómo lo escribió. Si es un producto nuevo, regístrelo en Entrada de mercancía." y botón para ir a esa pantalla.'],
      ['El producto está agotado. (P)', '"Agotado: no hay unidades en la tienda. Ya quedó en la lista de productos por pedir."'],
    ],
  },
  {
    id: 'HU-03', t: 'Ver alertas de vencimiento', sp: 3,
    narr: BACKLOG[2][1],
    crit: [
      ['hay lotes que vencen en 15 días o menos', 'el dueño abre "Vencimientos"', 'los ve en tres grupos de colores (vencidos en rojo, de 0 a 7 días en naranja, de 8 a 15 días en amarillo), ordenados por fecha, con su ubicación y la plata en riesgo.'],
      ['un lote ya está vencido', 'aparece en la lista', 'muestra "Vencido hace N días · no se puede vender" y el botón "Dar de baja".'],
      ['un lote vence esta semana', 'el dueño lo pasa al frente del estante y toca "Ya lo puse adelante"', 'el lote queda marcado "Ya está adelante".'],
      ['hay productos vencidos en los estantes', 'el dueño abre Inicio', 'la tarjeta "Vencidos en estante" muestra la cantidad en rojo.'],
    ],
    feliz: [
      'En Inicio ve la tarjeta roja "Vencidos en estante: 3".',
      'Toca la tarjeta o "Vencen" en la barra inferior.',
      'Ve el grupo de vencidos con su ubicación, por ejemplo Yogurt fresa 200 g en Nevera, estante N1, nivel 3.',
      'Toca "Dar de baja" y confirma con "Sí, dar de baja".',
      'Ve el aviso "Producto dado de baja. La pérdida quedó registrada en el cierre de caja."',
      'En el grupo de 7 días pone adelante el lote y toca "Ya lo puse adelante".',
    ],
    sad: [
      ['No hay productos por vencer en 15 días. (P)', '"Todo en orden. Ningún producto vence en los próximos 15 días."'],
      ['Se intenta vender un lote vencido. (P)', 'El lote no se descuenta en la venta. Si solo quedan vencidos aparece "Las 3 unidades que quedan de Kumis 1 L están vencidas. No se pueden vender." (ver [[HU-04|hu04]]).'],
      ['El dueño cancela la baja. (P)', 'La ventana se cierra con "Cancelar" y el lote sigue en la lista sin cambios.'],
      ['El producto no se vence (aseo). (P)', 'No aparece en las alertas y en su ficha se muestra "No se vence".'],
    ],
  },
  {
    id: 'HU-04', t: 'Registrar una venta', sp: 5,
    narr: BACKLOG[3][1],
    crit: [
      ['hay productos en la venta actual', 'el ayudante toca + o −', 'la cantidad y el total se actualizan de inmediato.'],
      ['el cliente paga en efectivo', 'escribe con cuánto paga', 've el cambio en verde.'],
      ['el pago alcanza', 'toca "Confirmar venta"', 've "Venta registrada" con el total, el cambio y el texto "El inventario ya se descontó", y las unidades bajan empezando por el lote que vence antes.'],
      ['un producto queda por debajo del mínimo', 'se confirma la venta', 'aparece "Quedó bajo: [producto]. Ya está en la lista por pedir."'],
    ],
    feliz: [
      'Toca "Vender".',
      'Busca "arroz" y toca "+ Agregar" dos veces. Agrega también "Queso costeño 500 g".',
      'Revisa el total en la barra de abajo y toca "Cobrar".',
      'Toca el botón de $ 50.000 o escribe con cuánto paga el cliente y ve el cambio.',
      'Toca "Confirmar venta".',
      'Ve "Venta registrada" y toca "Nueva venta".',
    ],
    sad: [
      ['Pide más unidades de las que hay. (P)', '"Solo quedan 4 unidades de Queso costeño 500 g. Ajuste la cantidad."'],
      ['Solo quedan unidades vencidas. (P)', 'Ventana "No se puede vender": "Las 3 unidades que quedan de Kumis 1 L están vencidas. No se pueden vender. Retírelas del estante en Vencimientos."'],
      ['El dinero recibido no alcanza. (P)', '"El dinero recibido es menor que el total. Faltan $ 9.200." y la venta no se confirma.'],
      ['Toca Cobrar sin productos. (P)', '"Agregue al menos un producto antes de cobrar."'],
      ['La caja del día ya se cerró. (P)', 'Ventana "La caja de hoy ya está cerrada": "Se cerró a las 9:30 p. m. Las ventas nuevas se registran mañana."'],
    ],
  },
  {
    id: 'HU-05', t: 'Cerrar la caja y ver la ganancia del día', sp: 5,
    narr: BACKLOG[4][1],
    crit: [
      ['hay ventas, gastos y bajas en el día', 'el dueño abre "Caja"', 've ventas, costo de lo vendido, gastos, pérdidas y la ganancia del día en verde si es positiva o en rojo si es negativa.'],
      ['registra un gasto con "+ Gasto"', 'lo guarda', 'el gasto aparece en la lista y la ganancia del día baja en ese valor.'],
      ['el efectivo contado coincide con el esperado', 'toca "Cerrar caja"', 'la caja se cierra y aparece "La caja de hoy ya se cerró a las [hora]".'],
      ['el efectivo contado no coincide', 'toca "Cerrar caja"', 've la diferencia y solo puede cerrar si vuelve a contar o escribe una nota.'],
    ],
    feliz: [
      'Toca "Caja" en la barra inferior.',
      'Revisa el resumen: ventas, costo, gastos, pérdidas y ganancia del día.',
      'Si falta un gasto lo agrega con "+ Gasto".',
      'Cuenta billetes y monedas y escribe el efectivo contado.',
      'Toca "Cerrar caja".',
      'Ve "Caja cerrada. El resumen del día quedó guardado." y el cierre queda para consulta.',
    ],
    sad: [
      ['No escribió el efectivo contado. (P)', '"Escriba el efectivo que contó en la caja."'],
      ['Hay diferencia con lo esperado. (P)', '"Hay una diferencia de $ 15.000. Falta plata: el sistema esperaba $ 351.800 y usted contó $ 336.800. Revise antes de cerrar." Para cerrar pide nota; sin nota: "Escriba una nota para cerrar con diferencia."'],
      ['Intenta cerrar dos veces el mismo día. (P)', '"La caja de hoy ya se cerró a las 9:30 p. m. Puede consultar el cierre pero no repetirlo."'],
      ['No hubo ventas en el día. (P)', '"Hoy no hay ventas registradas. ¿Desea cerrar la caja en cero?" con "No, volver" y "Sí, cerrar".'],
    ],
  },
];
HUS.forEach((h, k) => {
  if (k > 0) ch.push(BR());
  ch.push(H2(`${h.id}. ${h.t} (${h.sp} SP, prioridad alta)`));
  ch.push(H3('Narrativa'));
  ch.push(P(h.narr));
  ch.push(H3('Criterios de aceptación'));
  ch.push(tabla(['#', 'DADO', 'CUANDO', 'ENTONCES'], h.crit.map((c, i) => [`CA${i + 1}`, c[0], c[1], c[2]]), [600, 2600, 2560, 3600], { center: [0], size: 18 }));
  ch.push(new Paragraph({ spacing: { after: 60 } }));
  ch.push(H3('Flujo feliz'));
  lista(h.feliz).forEach(p => ch.push(p));
  ch.push(H3('Flujos alternos'));
  ch.push(tabla(['#', 'Condición de falla', 'Respuesta exacta del sistema'], h.sad.map((s, i) => [`SP${i + 1}`, s[0], s[1]]), [600, 3000, 5760], { center: [0], size: 18 }));
  ch.push(new Paragraph({ spacing: { after: 100 } }));
});
ch.push(BR());

// 7. ESTIMACIONES
ch.push(H1('7. Estimaciones'));
ch.push(H2('7.1 Estimación en story points'));
ch.push(P('Primero le pedimos a la IA una estimación de referencia para cada historia, con la secuencia de Fibonacci (1, 2, 3, 5, 8, 13) y una justificación corta. Después hicimos planning poker el 10 de noviembre: los seis integrantes votaron en secreto con cartas, se mostraron a la vez y, cuando los votos se separaban más de un valor de la secuencia, el que votó más alto y el que votó más bajo explicaron su razón antes de una segunda ronda. HU-05 y HU-11 necesitaron dos rondas. La estimación de la IA se conoció al final, para que no condicionara al equipo.'));
ch.push(tabla(['HU', 'IA', 'Justificación de la IA', 'Equipo', 'Razón de la diferencia'],
  EST.map(r => [r[0], String(r[1]), r[2], String(r[3]), r[4]]).concat([['**Total**', `**${totalIA}**`, '', `**${totalSP}**`, '']]),
  [800, 600, 2700, 900, 4360], { center: [0, 1, 3], size: 17, fills: (r, i) => (i === 3 && r[1] !== r[3] && r[0].startsWith('HU')) ? 'FFF4D1' : undefined }));
ch.push(cap('Estimación de la IA frente a la estimación del equipo en planning poker. En amarillo, las historias donde hubo diferencia.'));
ch.push(P(`Hubo diferencia en 8 de las 13 historias. La IA tendió a subestimar lo que depende de la realidad de la tienda (lotes, ubicaciones, vencidos, trabajo sin internet) y a sobrestimar lo que imaginó más grande de lo que el dueño necesita (un editor de planos, un reporte contable, una integración con WhatsApp). En total la IA sumó ${totalIA} SP y el equipo ${totalSP} SP.`));
ch.push(H2('7.2 Conversión a tiempo'));
ch.push(tabla(['Dato', 'Valor'], [
  ['Integrantes del equipo', '6'],
  ['Horas por semana de cada integrante', '6 horas (fuera de clases y trabajo)'],
  ['Capacidad del equipo', '6 × 6 = 36 horas por semana'],
  ['Equivalencia', `1 SP = ${HORAS_SP} horas de trabajo`],
  ['Capacidad teórica de un sprint de 2 semanas', '72 horas ÷ 4 = 18 SP'],
  ['Velocidad histórica del equipo (proyecto EcoAcción)', '16 SP por sprint en promedio'],
  ['Esfuerzo total del backlog', `${totalSP} SP × ${HORAS_SP} h = ${horasTot} horas`],
  ['Duración estimada', `${horasTot} h ÷ 36 h por semana ≈ ${(horasTot / 36).toFixed(1).replace('.', ',')} semanas, es decir 3 sprints de 2 semanas`],
], [4200, 5160]));
ch.push(cap('Capacidad del equipo y conversión de story points a tiempo.'));
ch.push(P('Planeamos cada sprint entre 15 y 18 SP. Usamos como techo la capacidad teórica de 18 SP, pero tomamos en cuenta que en EcoAcción el equipo completó en promedio 16 SP por sprint y que en el primer sprint de ese proyecto nos comprometimos con más de lo que alcanzamos.'));
ch.push(H2('7.3 Estimación de costo'));
ch.push(P(`Las ${horasTot} horas se repartieron por rol según el tipo de trabajo de cada historia. Las tarifas son valores de referencia para perfiles junior en Colombia. Como equipo de estudiantes no le cobramos al dueño, pero el cálculo muestra lo que costaría el proyecto en el mercado.`));
ch.push(tabla(['Rol', 'Horas', 'Tarifa por hora', 'Subtotal'],
  ROLES.map(r => [r[0], String(r[1]), $(r[2]), $(r[1] * r[2])]).concat([['**Total mano de obra**', `**${sumH}**`, '', `**${$(costoRoles)}**`]]),
  [4760, 1000, 1700, 1900], { center: [1], size: 18 }));
ch.push(cap('Costo de mano de obra por rol.'));
ch.push(tabla(['Herramienta o servicio', 'Costo'],
  HERR.map(r => [r[0], r[1] ? $(r[1]) : '$ 0']).concat([['**Total herramientas y hosting**', `**${$(costoHerr)}**`]]),
  [7460, 1900], { size: 18 }));
ch.push(cap('Costo de herramientas y hosting.'));
ch.push(tabla(['Concepto', 'Valor'], [
  ['Mano de obra', $(costoRoles)],
  ['Herramientas y hosting', $(costoHerr)],
  ['Subtotal', $(subtotal)],
  ['Contingencia (10 %) para imprevistos y ajustes tras la revisión con el dueño', $(conting)],
  ['**Total estimado del proyecto**', `**${$(total)}**`],
], [7460, 1900], { size: 19 }));
ch.push(cap('Costo total estimado del proyecto.'));
ch.push(P('Después de la entrega, mantener la aplicación en línea no le cuesta nada mensual al dueño mientras use los planes gratuitos de Netlify y Supabase, que alcanzan de sobra para una sola tienda. Solo pagaría la renovación del dominio cada año.'));
ch.push(H2('7.4 Cronograma de sprints'));
ch.push(P('Ordenamos los sprints para que el dueño vea valor desde el primero: al terminar el Sprint 1 ya puede guardar la mercancía con ubicación y ver los vencimientos, que es lo que él dijo que quería resolver primero.'));
ch.push(tabla(['Sprint', 'Fechas', 'Historias', 'SP', 'Qué se entrega al final'], [
  ['Sprint 1', '23 nov a 4 dic 2026', 'HU-08, HU-01, HU-02, HU-03, HU-07', '16', 'Mapa de la tienda, entrada de mercancía con ubicación, búsqueda "¿Dónde está?", alertas de vencimiento y baja de vencidos. Se instala en el celular del dueño.'],
  ['Sprint 2', '7 a 18 dic 2026', 'HU-04, HU-05, HU-09, HU-06', '15', 'Ventas que descuentan inventario, cierre de caja con ganancia del día, gastos y lista por pedir. El dueño puede operar el día completo con la app.'],
  ['Sprint 3', '12 a 22 ene 2027', 'HU-10, HU-11, HU-12, HU-13', '18', 'Resumen mensual, conteo físico, proveedores y pedido, y PIN del ayudante. Sincronización con la nube y entrega final.'],
  ['**Total**', '6 semanas de trabajo', '13 historias', `**${totalSP}**`, ''],
], [1000, 1500, 2100, 600, 4160], { center: [0, 3], size: 18 }));
ch.push(cap('Distribución de las historias en sprints de dos semanas.'));
ch.push(P('Entre el Sprint 2 y el Sprint 3 hay una pausa por las festividades de fin de año, que en la tienda son temporada alta y el dueño no tiene tiempo para revisiones. Si el Sprint 3 se aprieta, HU-12 (prioridad baja) pasa a un sprint corto de cierre sin afectar lo demás.'));
ch.push(BR());

// 8. STACK
ch.push(H1('8. Stack tecnológico'));
ch.push(H2('8.1 Tecnología por capa'));
ch.push(tabla(['Capa', 'Tecnología elegida', 'Justificación'], [
  ['**Frontend** (lo que ve el dueño)', 'HTML, CSS y JavaScript como aplicación web progresiva (PWA)', 'Se instala en el celular desde el navegador, sin pasar por la Play Store, y abre aunque no haya internet. El equipo ya trabajó así en EcoAcción. No elegimos React porque la app es pequeña y agrega peso y una curva de aprendizaje que no se justifica; tampoco una app nativa en Kotlin, porque nadie del equipo la domina y obliga a publicar en la tienda de aplicaciones.'],
  ['**Backend** (lógica del servidor)', 'Supabase (API REST automática, autenticación y reglas por rol)', 'Nos da la API y el control de acceso del dueño y del ayudante sin programar ni mantener un servidor propio. Un backend en Flask, que también conocemos, nos obligaría a pagar y cuidar un servidor. Las reglas del negocio que deben funcionar sin internet (vencidos, ventas, cierre) corren en el celular.'],
  ['**Base de datos**', 'IndexedDB en el celular y PostgreSQL de Supabase en la nube', 'La venta no puede detenerse si se cae el internet o se acaban los datos del celular, por eso los datos viven primero en el celular (IndexedDB). Cuando hay conexión se sincronizan con PostgreSQL, que sirve de respaldo si se pierde o daña el celular y permite el resumen mensual. PostgreSQL es relacional, lo que encaja con productos, lotes, ubicaciones y ventas.'],
  ['**Hosting e infraestructura**', 'Netlify para la app, Supabase para datos, GitHub para el código', 'Los dos tienen plan gratuito suficiente para una tienda, HTTPS incluido (requisito de las PWA) y despliegue automático desde GitHub. Ya publicamos EcoAcción en Netlify, así que conocemos el proceso.'],
  ['**Herramientas de IA**', 'Claude (Anthropic)', 'Borradores de historias y criterios, propuesta de flujos alternos, estimaciones de referencia, recomendación de stack y primer borrador del diagrama. Todo lo revisó y decidió el equipo, como se documenta en las secciones 5.2 y 7.1.'],
  ['**Gestión del proyecto**', 'Trello, Google Forms y GitHub', 'Tablero Scrum con backlog y sprints, el cuestionario de elicitación en Google Forms y control de versiones con ramas y pull requests.'],
], [1900, 2400, 5060], { size: 18 }));
ch.push(cap('Stack tecnológico por capa y justificación.'));
ch.push(H2('8.2 Diagrama de arquitectura'));
ch.push(P('El primer borrador lo generó la IA. El equipo lo ajustó para que se viera que la app funciona en el celular sin conexión y que la nube solo se usa para respaldo y sincronización. También se agregaron los dos actores y las reglas que corren en el celular.'));
figura('img/arquitectura.png', 600, 330, 'Arquitectura de Piragua Control: la PWA y sus datos viven en el celular y se sincronizan con Supabase cuando hay internet.').forEach(x => ch.push(x));
ch.push(P('El flujo es sencillo: el celular descarga la app desde Netlify una sola vez y la guarda con el service worker. Desde ahí, cada entrada, venta o cierre se guarda en IndexedDB y queda en una cola de cambios. Cuando hay internet, la cola se envía a Supabase por HTTPS, que valida el rol (dueño o ayudante) y guarda en PostgreSQL.'));
ch.push(BR());

// 9. PROTOTIPO
ch.push(H1('9. Link del prototipo'));
ch.push(H2('9.1 Enlace y cómo recorrerlo'));
ch.push(new Paragraph({ children: [new ExternalHyperlink({ link: 'https://piragua-control.netlify.app', children: [new TextRun({ text: 'https://piragua-control.netlify.app', font: F, size: 28, bold: true, color: '0563C1', underline: {} })] })], alignment: AlignmentType.CENTER, spacing: { before: 120, after: 200 } }));
ch.push(P('El prototipo es navegable e interactivo: los formularios validan, las ventas descuentan unidades, las bajas se reflejan en el cierre de caja y los datos se guardan en el navegador mientras se usa. Está pensado para el celular y en el computador se adapta a una versión de escritorio con menú lateral. El acceso es público y no pide usuario ni contraseña. Al pie de la pantalla de Inicio está el enlace "Restablecer datos iniciales" para volver a empezar.'));
ch.push(tabla(['HU', 'Cómo recorrer el flujo feliz', 'Figuras'], [
  ['HU-01', 'Inicio, "Entrada de mercancía", escribir "leche", cantidad y fecha, "Usar esa ubicación", "Guardar entrada".', '3 y 4'],
  ['HU-02', 'Barra inferior "¿Dónde está?", escribir "leche", tocar el resultado.', '6 y 7'],
  ['HU-03', 'Tarjeta roja de Inicio o "Vencen", "Dar de baja", "Sí, dar de baja".', '10 y 11'],
  ['HU-04', '"Vender", agregar arroz y queso, "Cobrar", elegir con cuánto paga, "Confirmar venta".', '14, 16 y 17'],
  ['HU-05', '"Caja", escribir el efectivo contado, "Cerrar caja".', '18 y 21'],
].map(r => [`[[${r[0]}|${r[0].replace('-', '').toLowerCase()}]]`, r[1], r[2]]), [900, 6460, 2000], { center: [0, 2], size: 18 }));
ch.push(cap('Recorrido de los cinco flujos felices en el prototipo.'));
ch.push(H2('9.2 Flujos alternos visibles'));
ch.push(P('La guía pide mínimo dos flujos alternos en el prototipo. Construimos todos los que están marcados con (P) en la sección 6. Estos son los más fáciles de probar:'));
ch.push(tabla(['HU', 'Cómo provocarlo', 'Qué se ve', 'Figura'], [
  ['HU-01', 'En Entrada de mercancía poner cantidad 0, una fecha pasada y no elegir zona.', 'Tres mensajes en rojo y la entrada no se guarda.', '5'],
  ['HU-02', 'Buscar "mortadela" y luego "harina".', 'No encontrado con botón a Entrada, y "Agotado".', '8 y 9'],
  ['HU-04', 'Vender "kumis".', 'Ventana "No se puede vender" porque las unidades están vencidas.', '12'],
  ['HU-04', 'Agregar cinco veces "Queso costeño 500 g".', 'Aviso "Solo quedan 4 unidades".', '13'],
  ['HU-04', 'En Cobrar escribir menos plata que el total.', '"El dinero recibido es menor que el total. Faltan ..."', '15'],
  ['HU-05', 'Tocar "Cerrar caja" sin escribir el efectivo.', '"Escriba el efectivo que contó en la caja."', '19'],
  ['HU-05', 'Escribir un efectivo distinto al esperado.', 'Ventana con la diferencia y nota obligatoria.', '20'],
  ['HU-04 y HU-05', 'Después de cerrar la caja, intentar cobrar una venta.', '"La caja de hoy ya está cerrada."', '22'],
], [1200, 3460, 3700, 1000], { center: [0, 3], size: 18 }));
ch.push(cap('Flujos alternos que se pueden probar en el prototipo.'));
ch.push(H2('9.3 Revisión con el dueño de la tienda'));
ch.push(P('El 15 de noviembre le mostramos el prototipo a don Alfredo en su propio celular, en la tienda. Recorrió con nosotros los cinco flujos y provocamos los mensajes de error para ver si los entendía sin ayuda. Lo que revisamos y lo que se ajustó:'));
ch.push(tabla(['Qué revisamos', 'Resultado y ajuste'], [
  ['Si entiende la pantalla de Inicio sin explicación.', 'Llegó solo a "Vender" y a "Vencimientos". Las tarjetas de colores se mantienen.'],
  ['Si la división en zonas, estantes y niveles coincide con su tienda.', 'Se confirmaron las cuatro zonas: mostrador, pasillo, nevera y bodega de atrás.'],
  ['Si los mensajes de error se entienden.', 'Se escribieron con palabras de la tienda ("plata", "dar de baja", "ponerlo adelante") en lugar de términos técnicos.'],
  ['Tamaño de letra y botones.', 'Se usó letra grande y botones de ancho completo para usar la app con una mano mientras atiende.'],
  ['Cierre de caja.', 'Se dejó la nota obligatoria cuando hay diferencia, porque a veces le fía a clientes conocidos y quiere dejarlo anotado.'],
], [3800, 5560], { size: 18 }));
ch.push(cap('Revisión del prototipo con el dueño de La Piragua.'));
ch.push(H2('9.4 Capturas de los flujos principales'));
ch.push(P('Respaldo visual de los flujos felices y alternos, tomado del prototipo publicado.'));
nFig = 1; // las figuras de capturas empiezan en 2 (la 1 es el diagrama)
capturas([
  ['01_inicio', 'Inicio con alertas del día.'],
  ['02_entrada_form', 'HU-01: entrada con ubicación sugerida.'],
  ['03_entrada_exito', 'HU-01: entrada guardada.'],
  ['04_entrada_errores', 'HU-01: cantidad en cero y fecha vencida.'],
  ['05_buscar', 'HU-02: búsqueda con ubicación.'],
  ['06_producto', 'HU-02: lotes y estante en el mapa.'],
  ['07_buscar_noexiste', 'HU-02: producto no encontrado.'],
  ['08_buscar_agotado', 'HU-02: producto agotado.'],
  ['09_vencen', 'HU-03: vencimientos por colores.'],
  ['10_baja_modal', 'HU-03: dar de baja un vencido.'],
  ['11_vender_vencido', 'HU-04: no se vende lo vencido.'],
  ['12_vender_sinstock', 'HU-04: más unidades de las que hay.'],
  ['13_vender_carro', 'HU-04: venta en curso.'],
  ['14_cobrar_corto', 'HU-04: el dinero no alcanza.'],
  ['15_cobrar_ok', 'HU-04: cálculo del cambio.'],
  ['16_venta_ok', 'HU-04: venta registrada.'],
  ['17_caja', 'HU-05: resumen y ganancia del día.'],
  ['18_caja_vacio', 'HU-05: falta el efectivo contado.'],
  ['19_caja_diferencia', 'HU-05: diferencia en la caja.'],
  ['20_caja_cerrada', 'HU-05: caja cerrada.'],
  ['21_venta_caja_cerrada', 'HU-04: venta con la caja cerrada.'],
  ['22_mapa', 'HU-08: mapa de la tienda.'],
  ['23_faltantes', 'HU-06: productos por pedir.'],
]).forEach(t => ch.push(t));
ch.push(H2('9.5 Enlaces de la entrega'));
ch.push(tabla(['Recurso', 'Enlace'], [
  ['Prototipo navegable (Netlify)', 'https://piragua-control.netlify.app'],
  ['Repositorio de código (GitHub)', 'https://github.com/PaoloManciniV/piragua-control'],
  ['Tablero Scrum (Trello)', 'https://trello.com/b/z9pio1m6'],
], [3200, 6160], { size: 19 }));
ch.push(cap('Enlaces del proyecto.'));
ch.push(P('El tablero de Trello tiene el backlog con las 13 historias, sus criterios y su definición de terminado, la lista de propuestas de cambio, los impedimentos y las métricas de estimación y plan de sprints. El repositorio publica el prototipo en Netlify cada vez que se actualiza la rama main.'));

// ---------- documento ----------
const doc = new Document({
  creator: 'Equipo Piragua Control', title: 'Piragua Control - Proyecto final',
  styles: { default: { document: { run: { font: F, size: 22 } } } },
  numbering: {
    config: [
      { reference: 'vin', levels: [{ level: 0, format: 'bullet', text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      ...numRefs.map(r => ({ reference: r, levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] }))
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } }, titlePage: true },
    footers: {
      default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: F, size: 18 })] })] }),
      first: new Footer({ children: [new Paragraph('')] })
    },
    children: ch
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync('out/Proyecto_Final_Piragua_Control.docx', b); console.log('ok', totalSP, totalIA, horasTot, costoRoles, costoHerr, total); });
module.exports = { TOC };
