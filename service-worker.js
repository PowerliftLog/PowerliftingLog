const CACHE_NAME = 'liftlog-v6';
const ASSETS = [
  '/PowerliftingLog/',
  '/PowerliftingLog/index.html',
  '/PowerliftingLog/parsers.js',
  '/PowerliftingLog/manifest.json',
  '/PowerliftingLog/icons/icon-192.png',
  '/PowerliftingLog/icons/icon-512.png',
  '/PowerliftingLog/icons/icon-180.png',
  // CDN dependencies — required for offline use
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];

// Google Fonts are fetched dynamically (CSS then WOFF2 files).
// We cache them on first use via the fetch handler below.

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

// Fetch — cache-first for assets, with background refresh
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // Google API scripts (GIS, GAPI) — network only, not needed offline
  if (url.hostname === 'accounts.google.com' || url.hostname === 'apis.google.com') return;

  // Cache-first strategy: serve from cache instantly, update in background
  event.respondWith(
    caches.match(event.request).then(cached => {
      // Background refresh — fetch new version and update cache
      const fetchPromise = fetch(event.request).then(response => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => null);

      // Return cached immediately if available, otherwise wait for network
      if (cached) return cached;
      return fetchPromise.then(resp => resp || new Response('Offline', { status: 503 }));
    })
  );
});
