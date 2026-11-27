/* Necesario para instalar como app. No guarda copias: siempre usa la versión publicada y los datos en línea. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
