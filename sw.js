// Network-first cache of the app shell so the home-screen app opens fast and updates on every deploy.
const CACHE = 'attendance-v19';
const SHELL = ['./', './index.html', './admin.html', './app.css', './common.js', './employee.js', './native.js', './admin.js', './config.js',
  './icon.svg', './icon-192.png', './tgra-mark.png', './manifest.json', './manifest-admin.json'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).catch(function () {}));
  self.skipWaiting();
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});
self.addEventListener('fetch', function (e) {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // never touch API calls
  // always ask the server (cheap 304 when unchanged) so a new version never mixes with old files
  e.respondWith(fetch(e.request.url, { cache: 'no-cache', credentials: 'same-origin' }).then(function (res) {
    const copy = res.clone();
    caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
    return res;
  }).catch(function () { return caches.match(e.request, { ignoreSearch: true }); }));
});
