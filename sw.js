const CACHE = 'mt-v7';
const CDN = 'mt-cdn-v1';   // on-device slip reader (OCR engine + language data) — kept across app updates
const SHELL = ['./', './index.html', './config.js', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE && k !== CDN).map((k) => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === 'https://cdn.jsdelivr.net') {
    // cache first: after the first scan the reader starts instantly and works offline
    e.respondWith(caches.open(CDN).then((c) => c.match(req.url).then((hit) => hit || fetch(req.url, { mode: 'cors', credentials: 'omit' }).then((res) => { if (res.ok) c.put(req.url, res.clone()); return res; }))));
    return;
  }
  if (url.origin !== location.origin) return;
  // network first so updates show up; fall back to cache when offline
  e.respondWith(fetch(req).then((res) => { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(req, copy)); return res; }).catch(() => caches.match(req).then((r) => r || caches.match('./index.html'))));
});
