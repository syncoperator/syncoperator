// CitiTool Bestellliste – Service Worker
// Netz zuerst: jede neue Version kommt sofort an, offline gilt die letzte Kopie.
const CACHE = 'bestellliste-v15';

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './index.html']).catch(() => {})));
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith('bestellliste') && k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== location.origin) return;
  e.respondWith((async () => {
    try {
      const res = await fetch(r, { cache: 'no-store' });
      if (res && res.ok) { const c = await caches.open(CACHE); c.put(r, res.clone()); }
      return res;
    } catch {
      const hit = await caches.match(r, { ignoreSearch: r.mode === 'navigate' });
      return hit || (r.mode === 'navigate' ? caches.match('./index.html') : Response.error());
    }
  })());
});
