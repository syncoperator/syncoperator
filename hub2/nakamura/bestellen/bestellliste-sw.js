/* Werkzeug-Bestellliste – Offline-Cache.
   Betrifft NUR die Bestellliste-Dateien; alle anderen Seiten von syncoperator.github.io laufen unverändert durch. */
const CACHE = 'bestellliste-v3';
const OWN = ['bestellliste.html', 'bestellliste.webmanifest', 'bestellliste-icon-180.png', 'bestellliste-icon-192.png', 'bestellliste-icon-512.png', 'bestellliste-icon-maskable-512.png'];
const EXT = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'unpkg.com'];
// Seite kann als bestellliste.html ODER als index.html / Ordner-URL ausgeliefert werden
const SCOPE = new URL(self.registration.scope).pathname;
const isPage = p => p.endsWith('/bestellliste.html') || (p.startsWith(SCOPE) && (p === SCOPE || p.endsWith('/index.html')));

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(OWN.map(f => c.add(new Request(f, { cache: 'reload' })).catch(() => null)))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('bestellliste-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const page = url.origin === location.origin && (isPage(url.pathname) || (req.mode === 'navigate' && url.pathname.startsWith(SCOPE)));
  const own = page || (url.origin === location.origin && OWN.some(f => url.pathname.endsWith('/' + f)));
  const ext = EXT.includes(url.hostname);
  if (!own && !ext) return;                       // alles andere: nicht anfassen
  if (page) {
    // Seite: zuerst Netz (immer neueste Version), offline aus dem Cache
    e.respondWith(fetch(req).then(r => { if (r.ok) { const cp = r.clone(); caches.open(CACHE).then(c => c.put('__page', cp)); } return r; })
      .catch(() => caches.match('__page').then(h => h || caches.match('bestellliste.html'))));
    return;
  }
  // Icons, Schriften, Excel-Bibliothek: zuerst Cache
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => { if (r.ok || r.type === 'opaque') { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); } return r; })));
});
