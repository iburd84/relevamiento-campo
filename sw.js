// Guarda la app en el celular para que abra sin señal (igual que Remitos de Campo).
// Solo archivos propios: el login de Microsoft y SharePoint siempre van por la red.
const VERSION = 'relevamiento-campo-1.1.0';
const ARCHIVOS = [
  './', 'index.html', 'styles.css', 'manifest.webmanifest',
  'img/logo_alonger.jpg', 'img/icono-192.png', 'img/icono-512.png', 'img/icono-512-maskable.png', 'img/icono-apple-180.png',
  'vendor/msal-browser-4.30.0.min.js', 'vendor/jszip-3.10.1.min.js',
  'js/app.js', 'js/config.js', 'js/data.js', 'js/auth.js', 'js/almacen.js', 'js/nube.js', 'js/bases.js',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION)
    .then(c => c.addAll(ARCHIVOS.map(u => new Request(u, { cache: 'reload' }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Primero la red (para tomar actualizaciones); sin señal, lo guardado.
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request, { cache: 'no-cache' })
      .then(r => {
        if (r.ok) { const copia = r.clone(); caches.open(VERSION).then(c => c.put(e.request, copia)); }
        return r;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('index.html'))),
  );
});
