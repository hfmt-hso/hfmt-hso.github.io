const CACHE = 'hso-portal-v3';
const ASSETS = ['/', '/index.html', '/unterricht.html', '/manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('fetch', e => {
  // Web App-Requests nie cachen
  if (e.request.url.includes('script.google.com') || e.request.url.includes('googleusercontent.com')) {
    e.respondWith(fetch(e.request));
    return;
  }

  // Nur GET-Requests cachen (Cache API unterstützt keine anderen Methoden)
  if (e.request.method !== 'GET') {
    e.respondWith(fetch(e.request));
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      return fetch(e.request).then(fresh => {
        // WICHTIG: sofort klonen, bevor der Body vom Browser konsumiert wird —
        // sonst schlägt clone() asynchron mit "Response body is already used" fehl.
        const copy = fresh.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
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