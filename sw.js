/* Service worker minimal: hanya agar aplikasi dapat dipasang ke Layar Utama. Tidak menyimpan cache. */
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () { /* biarkan jaringan menangani semua permintaan */ });
