const C='piragua-v7';
const ilustraciones = [
  'velero','nevera','caja','pedido','mapa-maqueta',
  'leche','arroz','huevos','aceite','cafe','azucar','pasta','atun','cola','pan','queso',
  'yogurt','aseo','papel','salchichas','despensa','panela','harina','kumis'
];
const recursos = ['./','./index.html','./manifest.json','./theme.css',...ilustraciones.map(nombre=>'./assets/'+nombre+'.webp')];
self.addEventListener('install',e=>e.waitUntil(caches.open(C).then(c=>c.addAll(recursos)).then(()=>self.skipWaiting())));
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('piragua-')&&k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request))));
