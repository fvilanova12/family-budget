// Keeps the app shell available offline. Always asks the network first and
// revalidates (GitHub Pages lets browsers cache files for 10 minutes), so
// updates show up straight away; Supabase requests are never cached.
const CACHE = 'our-budget-v2';
const SHELL = ['./', 'index.html', 'styles.css', 'app.js', 'ocr.js', 'config.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.hostname.endsWith('supabase.co')) return;
  const sameOrigin = url.origin === self.location.origin;
  const isLib = url.hostname === 'cdn.jsdelivr.net' || url.hostname.endsWith('gstatic.com') || url.hostname === 'fonts.googleapis.com';
  if (!sameOrigin && !isLib) return;
  e.respondWith(
    fetch(e.request, sameOrigin ? { cache: 'no-cache' } : undefined).then((res) => {
      if (res.ok && (sameOrigin || isLib)) {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
      }
      return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: sameOrigin }))
  );
});
