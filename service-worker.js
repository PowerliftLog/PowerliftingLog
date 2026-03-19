const CACHE_NAME = 'liftlog-v131';
const ASSETS = [
  '/PowerliftingLog/',
  '/PowerliftingLog/index.html',
  '/PowerliftingLog/parsers.js',
  '/PowerliftingLog/reportWorker.js',
  '/PowerliftingLog/manifest.json',
  '/PowerliftingLog/icons/icon-192.png',
  '/PowerliftingLog/icons/icon-512.png',
  '/PowerliftingLog/icons/icon-180.png',
  // SheetJS bundled locally — no CDN dependency for offline import
  '/PowerliftingLog/libs/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
];

// Google Fonts are fetched dynamically (CSS then WOFF2 files).
// We cache them on first use via the fetch handler below.

// Listen for SKIP_WAITING message from the page to activate a waiting SW
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Install — pre-cache all core assets
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
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

// Fetch — network-first for index.html, cache-first for everything else
self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);

  // Google API scripts (GIS, GAPI) — network only, not needed offline
  if (url.hostname === 'accounts.google.com' || url.hostname === 'apis.google.com') return;

  // Network-first for index.html and root — always get fresh version, fall back to cache offline
  const isNavigation = url.pathname.endsWith('/index.html') || url.pathname.endsWith('/PowerliftingLog/') || url.pathname === '/PowerliftingLog';
  if (isNavigation) {
    event.respondWith(
      fetch(event.request).then(response => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        }
        // Non-OK response (500, 403, etc.) — fall back to cache
        return caches.match(event.request).then(cached => cached || response);
      }).catch(() => caches.match(event.request).then(cached => cached || new Response('Offline', { status: 503 })))
    );
    return;
  }

  // Cache-first strategy for all other assets: serve from cache instantly, update in background
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
