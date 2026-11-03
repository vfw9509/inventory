// Keeps the app shell and scanner libraries on the phone so it opens with no signal.
// Data calls to Google (script.google.com) are never cached.
const CACHE = 'vfwinv-v1';
const SHELL = ['./', 'index.html', 'manifest.json', 'icon-192.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || /google|openfoodfacts|upcitemdb/.test(url.hostname)) return;
  const isLib = /unpkg\.com|cdnjs\.cloudflare\.com/.test(url.hostname);
  if (isLib) {
    // Versioned libraries: cache first.
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    })));
    return;
  }
  if (url.origin === location.origin) {
    // App files: network first so updates show up, cache when offline.
    e.respondWith(fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request).then(hit => hit || caches.match('index.html'))));
  }
});
