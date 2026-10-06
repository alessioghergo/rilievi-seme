/* Rilievi Seme: tiene in memoria l'app e le mappe già viste, così funziona anche con poco segnale. */
var V = 'rilievi-v6';
var SHELL = [
  './', 'index.html', 'manifest.webmanifest', 'icon-180.png', 'icon-512.png',
  'lib/leaflet/leaflet.css', 'lib/leaflet/leaflet.js',
  'lib/leaflet/images/layers.png', 'lib/leaflet/images/layers-2x.png',
  'lib/leaflet/images/marker-icon.png', 'lib/leaflet/images/marker-icon-2x.png', 'lib/leaflet/images/marker-shadow.png',
  'lib/xlsx.full.min.js', 'lib/jszip.min.js', 'lib/jspdf.umd.min.js'
];
var TILE_HOSTS = ['tile.openstreetmap.org', 'server.arcgisonline.com'];
var FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];
var TILE_MAX = 900;

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(V).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== V && k !== 'tiles' && k !== 'fonts'; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

function trim(cacheName, max) {
  return caches.open(cacheName).then(function (c) {
    return c.keys().then(function (keys) {
      var extra = keys.length - max;
      if (extra <= 0) return;
      return Promise.all(keys.slice(0, extra).map(function (k) { return c.delete(k); }));
    });
  });
}

function staleWhileRevalidate(req, cacheName, max) {
  return caches.open(cacheName).then(function (c) {
    return c.match(req).then(function (hit) {
      var net = fetch(req).then(function (res) {
        if (res && (res.ok || res.type === 'opaque')) {
          c.put(req, res.clone()).then(function () { if (max) trim(cacheName, max); });
        }
        return res;
      }).catch(function () { return hit; });
      return hit || net;
    });
  });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);

  if (url.origin === self.location.origin) {
    if (req.mode === 'navigate') {
      e.respondWith(
        fetch(req).then(function (res) {
          var copy = res.clone();
          caches.open(V).then(function (c) { c.put('index.html', copy); });
          return res;
        }).catch(function () {
          return caches.match('index.html').then(function (r) { return r || caches.match('./'); });
        })
      );
      return;
    }
    e.respondWith(
      caches.match(req, { ignoreSearch: true }).then(function (hit) {
        return hit || fetch(req).then(function (res) {
          if (res && res.ok) { var copy = res.clone(); caches.open(V).then(function (c) { c.put(req, copy); }); }
          return res;
        });
      })
    );
    return;
  }

  if (TILE_HOSTS.indexOf(url.hostname) >= 0) { e.respondWith(staleWhileRevalidate(req, 'tiles', TILE_MAX)); return; }
  if (FONT_HOSTS.indexOf(url.hostname) >= 0) { e.respondWith(staleWhileRevalidate(req, 'fonts', 0)); return; }
});
