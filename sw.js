// Service worker: permite abrir el formulario sin conexión.
// Al publicar una versión nueva del formulario, cambiar VERSION para renovar la caché.
const VERSION = 'mgpp-v1.0';
const ARCHIVOS = ['./', './index.html', './logo-egpp-blanco.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;                          // los envíos al Sheets nunca pasan por la caché
  if (url.hostname.endsWith('google.com') || url.hostname.endsWith('googleusercontent.com')) return;
  // Primero la red (para recibir actualizaciones); si no hay conexión, la copia guardada.
  e.respondWith(
    fetch(e.request).then(res => {
      const copia = res.clone();
      caches.open(VERSION).then(c => c.put(e.request, copia));
      return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
