const VER = 'po-expired-v9';
const SHELL = ['./', 'index.html', 'style.css', 'core.js', 'app.js', 'manifest.json', 'assets/icon.png'];
const CDN = ['https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    const c = await caches.open(VER);
    await c.addAll(SHELL);
    for (const u of CDN) { try { await c.add(new Request(u, { mode: 'cors' })); } catch (_) {} }
    self.skipWaiting();
  })());
});
self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== VER) await caches.delete(k);
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    // cache dulu, perbarui di latar belakang
    e.respondWith((async () => {
      const c = await caches.open(VER);
      const hit = await c.match(req, { ignoreSearch: true });
      const net = fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }).catch(() => null);
      return hit || (await net) || (req.mode === 'navigate' ? c.match('index.html') : Response.error());
    })());
  } else if (/cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net/.test(url.host)) {
    e.respondWith((async () => {
      const c = await caches.open(VER);
      const hit = await c.match(req);
      if (hit) return hit;
      const r = await fetch(req);
      if (r.ok || r.type === 'opaque') c.put(req, r.clone());
      return r;
    })());
  }
});
