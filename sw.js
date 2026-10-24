/* Rincovitch Site Inspections — offline service worker.
   Serves the app from cache (works with no signal) and serves stored photos at _blob/<id>. */
const VERSION = '9622bbc45dd3';
const CACHE = 'rinco-m365-' + VERSION;
const FILES = ["./", "index.html", "sync.js", "config.js", "manifest.webmanifest", "lib/jszip.min.js", "lib/pdf-lib.min.js", "fonts/Carlito-Regular.woff", "fonts/Carlito-Bold.woff", "fonts/Carlito-Italic.woff", "icons/icon-maskable-512.png", "icons/icon-192.png", "icons/apple-touch-icon.png", "icons/icon-512.png"];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k.startsWith('rinco-m365-') && k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function blobFromDb(id) {
  return new Promise((resolve) => {
    const r = indexedDB.open('rincovitch-inspections-m365', 1);
    r.onupgradeneeded = () => { r.transaction.abort(); resolve(null); };
    r.onerror = () => resolve(null);
    r.onsuccess = () => {
      const d = r.result;
      if (!d.objectStoreNames.contains('blobs')) { d.close(); resolve(null); return; }
      const g = d.transaction('blobs', 'readonly').objectStore('blobs').get(id);
      g.onsuccess = () => { d.close(); resolve(g.result || null); };
      g.onerror = () => { d.close(); resolve(null); };
    };
  });
}

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin || e.request.method !== 'GET') return;
  const m = /\/_blob\/([0-9a-z]+)$/i.exec(url.pathname);
  if (m) {
    e.respondWith(blobFromDb(m[1]).then((rec) => rec
      ? new Response(rec.blob, { headers: { 'Content-Type': rec.type || 'application/octet-stream', 'Cache-Control': 'no-store' } })
      : new Response('', { status: 404 })));
    return;
  }
  // settings file: always the latest when online (so editing config.js takes effect), cached copy when offline
  if (/\/config\.js$/.test(url.pathname)) {
    e.respondWith(fetch(e.request, { cache: 'no-store' }).then((res) => { if (res.ok) { const c = res.clone(); caches.open(CACHE).then((k) => k.put(e.request, c)); } return res; })
      .catch(() => caches.match(e.request, { ignoreSearch: true })));
    return;
  }
  // app files: from this version's cache (a new version arrives as a new service worker), else network
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => hit || fetch(e.request).catch(() =>
    e.request.mode === 'navigate' ? caches.match('./index.html') : Response.error())));
});
