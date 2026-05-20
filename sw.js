const CACHE = 'hso-portal-v2';
const ASSETS = ['/', '/index.html', '/intern.html', '/manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('fetch', e => {
  e.respondWith(
    caches.match(e.request).then(cached => {
      return fetch(e.request).then(fresh => {
        caches.open(CACHE).then(c => c.put(e.request, fresh.clone()));
        return fresh;
      }).catch(() => cached || caches.match('/'));
    })
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});
