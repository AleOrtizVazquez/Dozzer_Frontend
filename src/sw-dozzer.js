// Service worker de DOZZER:
// - Guarda una copia de la app (página, código, íconos) para que abra aunque no haya señal.
//   Siempre intenta primero la versión de internet; la copia solo se usa sin conexión.
// - Recibe los avisos (Web Push) aunque la app esté cerrada.
// Los datos (API en Render) y las fotos (Cloudinary) no pasan por aquí: los maneja la app.

const CACHE = 'dozzer-app-v1';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
  await self.clients.claim();
})()));

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  const llave = req.mode === 'navigate' ? '/' : req;
  e.respondWith((async () => {
    try {
      const r = await fetch(req);
      if (r.ok) {
        const copia = r.clone();
        caches.open(CACHE).then(c => c.put(llave, copia)).catch(() => {});
      }
      return r;
    } catch (err) {
      const guardada = await caches.match(llave);
      if (guardada) return guardada;
      throw err;
    }
  })());
});

self.addEventListener('push', e => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) { d = { cuerpo: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.titulo || 'DOZZER', {
    body: d.cuerpo || '',
    icon: 'icono-192.png',
    badge: 'icono-badge-96.png',
    tag: d.etiqueta || 'dozzer',
    renotify: true,
    vibrate: [180, 80, 180],
    data: { url: d.url || '/' }
  }));
});

// Al tocar el aviso: si DOZZER ya está abierto se lleva ahí; si no, se abre
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || '/', self.location.origin).href;
  e.waitUntil((async () => {
    const abiertas = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of abiertas) {
      if (new URL(c.url).origin === self.location.origin) {
        await c.focus();
        c.postMessage({ tipo: 'dozzer-ir', url });
        return;
      }
    }
    await self.clients.openWindow(url);
  })());
});
