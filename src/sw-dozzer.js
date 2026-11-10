// Service worker de DOZZER: recibe los avisos (Web Push) aunque la app esté cerrada.
// No guarda nada en caché: la app siempre se carga del servidor.

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

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
