/* Werkzeug-Bestellliste – Offline-Cache.
   Betrifft NUR die Bestellliste-Dateien; alle anderen Seiten von syncoperator.github.io laufen unverändert durch. */
const CACHE = 'bestellliste-v1';
const OWN = ['bestellliste.html', 'bestellliste.webmanifest', 'bestellliste-icon-180.png', 'bestellliste-icon-192.png', 'bestellliste-icon-512.png', 'bestellliste-icon-maskable-512.png'];
const EXT = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdnjs.cloudflare.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(OWN.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('bestellliste-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const own = url.origin === location.origin && OWN.some(f => url.pathname.endsWith('/' + f));
  const ext = EXT.includes(url.hostname);
  if (!own && !ext) return;                       // alles andere: nicht anfassen
  if (own && url.pathname.endsWith('bestellliste.html')) {
    // Seite: zuerst Netz (immer neueste Version), offline aus dem Cache
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put('bestellliste.html', cp)); return r; })
      .catch(() => caches.match('bestellliste.html')));
    return;
  }
  // Icons, Schriften, Excel-Bibliothek: zuerst Cache
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => { if (r.ok || r.type === 'opaque') { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); } return r; })));
});
