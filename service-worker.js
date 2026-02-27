const CACHE_NAME = 'liftlog-v2';
const ASSETS = [
  '/PowerliftingLog/',
  '/PowerliftingLog/index.html',
  '/PowerliftingLog/parsers.js',
  '/PowerliftingLog/manifest.json',
  '/PowerliftingLog/icons/icon-192.png',
  '/PowerliftingLog/icons/icon-512.png',
];
// Install — pre-cache all core assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});
// Activate — clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});
// Fetch — network-first, fall back to cache
self.addEventListener('fetch', event => {
  // Only handle GET requests for our origin
  if (event.request.method !== 'GET') return;
  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Cache a fresh copy if it's one of our assets
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
