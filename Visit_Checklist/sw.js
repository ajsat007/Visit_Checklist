/**
 * ============================================================================
 * SERVICE WORKER — Smart Services Tour Visit Checklist (PWA offline support)
 * ============================================================================
 */

const CACHE_VERSION = 'v2.0.0';
const CACHE_NAME = `smart-services-tour-visit-${CACHE_VERSION}`;

// Only list files that actually exist in this project. cache.addAll() is
// all-or-nothing — a single 404 (e.g. a page that was planned but never
// shipped) silently fails the *entire* precache batch. This build uses
// Promise.allSettled per-resource instead, so one bad/slow URL never takes
// the rest of the offline shell down with it.
const PRECACHE_URLS = [
  './',
  'index.html',
  'visit.html',
  'dashboard.html',
  'manifest.json',
  'assets/js/config.js',
  'assets/js/auth.js',
  'assets/js/api.js',
  'assets/js/utils.js',
  'assets/images/logo-full.png',
  'assets/images/logo-icon.png',
  'assets/images/icon-192.png',
  'assets/images/icon-512.png',
  'https://cdn.tailwindcss.com',
  'https://cdn.jsdelivr.net/npm/sweetalert2@11',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js',
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;500;600;700;800&display=swap'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      const results = await Promise.allSettled(
        PRECACHE_URLS.map((url) => cache.add(url))
      );
      results.forEach((r, i) => {
        if (r.status === 'rejected') {
          console.warn('[SW] Precache failed for', PRECACHE_URLS[i], r.reason);
        }
      });
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return; // never cache POST (all our API calls)
  const url = new URL(request.url);
  if (url.protocol === 'chrome-extension:') return;

  // Backend API calls: always try the network first; fall back to a clear
  // offline JSON response rather than a broken cached page.
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(networkFirst(request));
    return;
  }

  if (isStaticAsset(url.pathname)) {
    event.respondWith(cacheFirst(request));
    return;
  }

  event.respondWith(networkFirstWithCacheFallback(request));
});

async function cacheFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) cache.put(request, response.clone());
    return response;
  } catch (error) {
    return new Response('Offline', { status: 503 });
  }
}

async function networkFirst(request) {
  try {
    const response = await Promise.race([
      fetch(request),
      new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout')), 8000))
    ]);
    return response;
  } catch (error) {
    return new Response(
      JSON.stringify({ success: false, message: 'ऑफलाइन — सर्व्हरशी संपर्क होऊ शकला नाही', offline: true }),
      { status: 503, headers: { 'Content-Type': 'application/json' } }
    );
  }
}

async function networkFirstWithCacheFallback(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch (error) {
    const cache = await caches.open(CACHE_NAME);
    const cached = await cache.match(request);
    return cached || new Response('Offline — page not cached', { status: 503 });
  }
}

function isStaticAsset(pathname) {
  return ['.js', '.css', '.png', '.jpg', '.jpeg', '.svg', '.ico', '.woff', '.woff2'].some((ext) => pathname.endsWith(ext));
}

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
  if (event.data && event.data.type === 'CLEAR_CACHE') caches.delete(CACHE_NAME);
});
