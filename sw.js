const CACHE = 'robs-class-v5';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'logo.png'];
const CDN = [
  'https://cdn.tailwindcss.com',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(async c => {
      await Promise.all(SHELL.map(u => c.add(u).catch(() => {})));
      await Promise.all(CDN.map(u => c.add(new Request(u, { mode: 'no-cors' })).catch(() => {})));
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const isCdn = CDN.includes(req.url);
  if (url.origin !== location.origin && !isCdn) return; // never touch the webhook

  const store = res => {
    if (res && (res.ok || res.type === 'opaque')) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return res;
  };

  if (isCdn) {
    // libraries: cache first
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(store)));
    return;
  }
  // your own files: ask the server first (ignoring the browser's saved copy) so a new upload shows up right away; the cache is only for when offline
  e.respondWith(
    fetch(req, { cache: 'reload' }).then(store).catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('index.html')))
  );
});
