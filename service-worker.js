// service-worker.js
// Minimal app-shell cache for offline-friendly behavior.
//
// Notes:
// - This caches only local assets. Pyodide is loaded from a CDN and is not cached here.
// - Extend later: add a user prompt, versioning strategy, and optional CDN caching.

const CACHE_NAME = 'pkl-pivot-pwa-v60';

const APP_SHELL = [
  './',
  './index.html',
  './styles.css?v=33',
  './app.js?v=70',
  './pyodide-loader.js?v=33',
  './pivot.js?v=31',
  './manifest.json?v=1',
  './icons/icon.svg',
  './icons/maskable.svg',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      await cache.addAll(APP_SHELL);
      // Activate the new service worker ASAP (dev-friendly).
      await self.skipWaiting();
    })()
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => (k === CACHE_NAME ? null : caches.delete(k))));
      await self.clients.claim();
    })()
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Only handle same-origin requests for the app shell.
  if (url.origin !== self.location.origin) return;

  // Dev-friendly: for page navigations, prefer network so updates show up.
  if (req.mode === 'navigate') {
    event.respondWith(
      (async () => {
        try {
          const fresh = await fetch(req);
          const cache = await caches.open(CACHE_NAME);
          cache.put('./index.html', fresh.clone());
          return fresh;
        } catch {
          const cached = await caches.match('./index.html');
          return cached || Response.error();
        }
      })()
    );
    return;
  }

  event.respondWith(
    (async () => {
      const cached = await caches.match(req);
      if (cached) return cached;

      const res = await fetch(req);
      // Cache GET requests (best-effort).
      if (req.method === 'GET' && res.ok) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(req, res.clone());
      }
      return res;
    })()
  );
});
