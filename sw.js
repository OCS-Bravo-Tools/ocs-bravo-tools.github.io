/* Offline cache for the hosted app (https only). Cache first: this version's files
   open instantly with no signal. A new version comes whole with its new worker
   (install below), so old and new files never mix in one cache. Requests to other
   sites (the collector) are never touched. */
const CACHE = 'peer-eval-v2-2026-11-17n-04e4ec04';
const FILES = [
  './', './index.html', './cadre.html', './admin.html', './manifest.webmanifest',
  './css/oc.css', './css/cadre.css', './css/admin.css',
  './js/oc.js', './js/cadre.js', './js/admin.js',
  './js/lib/attributes.js', './js/lib/board.js', './js/lib/codec.js', './js/lib/collector2.js', './js/lib/config.js', './js/lib/crypto.js', './js/lib/deal.js',
  './js/lib/envelope.js', './js/lib/linkpdf.js', './js/lib/qrparts.js', './js/lib/report.js', './js/lib/roster.js', './js/lib/scoring.js', './js/lib/xlsx.js',
  './js/vendor/qrcode.js', './js/vendor/jsQR.js', './apps-script/Collector.gs', './apps-script/Collector2.gs', './apps-script/appsscript.json'
];

// 'reload' skips the browser's own cache (GitHub Pages lets it keep a file for 10 minutes),
// so a new version never stores the old files.
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;
  // nothing fetched here goes into the cache: a fresh copy of one file next to old copies of the others
  // could load a module the old ones don't fit
  e.respondWith(caches.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req)));
});
