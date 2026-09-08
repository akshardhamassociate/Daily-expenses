// Akshardham Associates expenses — service worker (installable + offline).
// Network-first: always try fresh, fall back to cache when offline.
const CACHE = 'akshardham-expenses-v1';
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;              // never cache the POST to Google Sheet
  e.respondWith(
    fetch(req)
      .then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)).catch(()=>{}); return res; })
      .catch(() => caches.match(req))
  );
});
