# Piragua Control

Aplicación web para el celular que ayuda a la tienda **La Piragua** (Calle 11 #14-06, Barranquilla) a saber qué mercancía tiene, dónde está guardada, qué se va a vencer y cuánto gana cada día.

Proyecto final de **Gestión de Proyectos (5925)**, Universidad Simón Bolívar, Ingeniería de Sistemas.
Profesor: Rafael Pérez Meriño.

**Prototipo publicado:** https://piragua-control.netlify.app

## Problema

Alfredo Ortega, dueño de La Piragua, pierde entre 150.000 y 300.000 pesos al mes en productos vencidos o dañados y no sabe cuánto gana, porque no tiene un registro que le diga qué mercancía tiene, dónde está guardada y cuándo se vence.

## Qué hace el prototipo

| HU | Pantalla | Qué permite |
|----|----------|-------------|
| HU-01 | Entrada de mercancía | Registrar lo que llega con cantidad, costo, vencimiento y ubicación (zona, estante y nivel). Sugiere el sitio donde ya está guardado el producto. |
| HU-02 | ¿Dónde está? | Buscar un producto y ver sus lotes, ubicación y el estante resaltado en el mapa. |
| HU-03 | Vencimientos | Lotes vencidos, por vencer en 7 días y en 15 días, con ubicación y plata en riesgo. Dar de baja y marcar "ya lo puse adelante". |
| HU-04 | Vender | Venta rápida que descuenta el inventario empezando por el lote que vence antes y bloquea lo vencido. |
| HU-05 | Cierre de caja | Ventas, costo, gastos, pérdidas y ganancia del día. Cierre con efectivo contado y nota si hay diferencia. |
| HU-06 | Por pedir | Productos bajo el mínimo con cantidad sugerida. |
| HU-08 | Mapa de la tienda | Zonas, estantes y lo que hay en cada nivel. |

Flujos alternos que se pueden probar: cantidad en cero, fecha vencida, entrada sin ubicación, producto no encontrado, producto agotado, venta de un producto vencido (kumis), más unidades de las que hay (queso), dinero insuficiente, cierre sin efectivo, diferencia en caja y venta con la caja cerrada.

Los datos se guardan en el navegador (localStorage). Al pie de Inicio está **Restablecer datos iniciales** para volver a empezar.

## Cómo correrlo

No necesita servidor ni instalación. Basta con abrir `index.html` en el navegador.
Para verlo como en el celular: F12 en Chrome y activar la vista de dispositivo móvil.

## Publicar en Netlify

Opción 1, arrastrando: entrar a https://app.netlify.com/drop y soltar la carpeta del proyecto. Luego en *Site configuration > Change site name* poner `piragua-control`.

Opción 2, desde GitHub: *Add new site > Import an existing project*, elegir este repositorio y dejar el directorio de publicación en `.` (ya está en `netlify.toml`). Cada merge a `main` se publica solo.

## Estructura

```
index.html        Prototipo completo (HTML, CSS y JavaScript)
manifest.json     Manifiesto de la PWA
sw.js             Service worker para abrir sin internet
netlify.toml      Configuración de publicación
docs/
  Proyecto_Final_Piragua_Control.pdf    Documento de entrega
  Proyecto_Final_Piragua_Control.docx   Versión editable
  fuentes/        Scripts que generan el documento, las capturas y el diagrama
```

Para regenerar el documento (requiere Node y LibreOffice):

```
cd docs/fuentes
npm install
npx playwright install chromium
npm run capturas
npm run diagrama
npm run documento
```

## Stack

- Frontend: PWA en HTML, CSS y JavaScript.
- Backend (versión final): Supabase con API REST, autenticación y reglas por rol.
- Base de datos: IndexedDB en el celular y PostgreSQL en Supabase.
- Hosting: Netlify.
- IA en el proceso: Claude (Anthropic), con revisión del equipo.

## Equipo

- Andrés Parra (Product Owner)
- Dylan Mariño
- Paolo Mancini Villalobos
- Juan Pablo Mora Gómez
- Milan Libardo Mejía Reales
- Juan Esteban Quintero Gelvez
