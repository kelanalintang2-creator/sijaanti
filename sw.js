/**
 * ============================================================================
 * SiJaAnti PRO - Service Worker (sw.js)
 * Standar: /gas-pwa-github
 * ============================================================================
 */

const CACHE_NAME = 'sijaanti-pwa-v2.0.0';
const STATIC_ASSETS = [
  './',
  './index.html',
  './css/style.css',
  './js/config.js',
  './js/app.js',
  './manifest.json',
  './icon.svg'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[ServiceWorker] Caching App Shell static assets...');
      return cache.addAll(STATIC_ASSETS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keyList) => {
      return Promise.all(
        keyList.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[ServiceWorker] Menghapus cache usang:', key);
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const reqUrl = new URL(event.request.url);

  // 1. Permintaan ke Google Apps Script / Google Sheets API JANGAN DICACHE oleh Service Worker
  if (reqUrl.origin.includes('script.google.com') || reqUrl.origin.includes('script.googleusercontent.com')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // 2. Resource CDN (Tailwind, ChartJS, SheetJS, Google Fonts)
  if (reqUrl.origin.includes('cdn.jsdelivr.net') || reqUrl.origin.includes('fonts.googleapis.com') || reqUrl.origin.includes('fonts.gstatic.com')) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        if (cached) return cached;
        return fetch(event.request).then((response) => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          }
          return response;
        }).catch(() => cached);
      })
    );
    return;
  }

  // 3. File Statis Lokal Aplikasi (Cache First, Network Fallback)
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) return cachedResponse;
      return fetch(event.request).catch((err) => {
        console.warn('[ServiceWorker] Fetch offline error:', err);
      });
    })
  );
});
