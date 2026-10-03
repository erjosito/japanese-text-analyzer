// Minimal service worker: cache the app shell for installability & fast repeat loads.
// Network calls to Azure OpenAI / Entra ID are never cached (always network).
const CACHE = 'jta-shell-v11';
const SHELL = [
  './',
  './index.html',
  './auth.html',
  './css/style.css?v=11',
  './js/vendor/msal-browser.min.js',
  './js/vendor/msal-popup-relay.min.js?v=5',
  './js/app.js?v=11',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // Keep same-origin app files current while retaining cached offline fallback.
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
