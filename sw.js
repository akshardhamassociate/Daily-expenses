// Akshardham Associates expenses — service worker (installable + offline).
// Caches only the app shell (html/css/icon). NEVER caches the Google Sheet API, so data is always fresh.
const CACHE = 'akshardham-expenses-v2';

self.addEventListener('install', () => self.skipWaiting());

// take control immediately and delete old caches (so a stale app/data copy can't linger)
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
));

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                      // POSTs to the Sheet: pass through
  if (req.url.indexOf('script.google') > -1) return;     // live API: always network, never cache
  e.respondWith(
    fetch(req)
      .then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)).catch(()=>{}); return res; })
      .catch(() => caches.match(req))
  );
});
