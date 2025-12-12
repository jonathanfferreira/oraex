self.addEventListener('install', (e) => {
    e.waitUntil(
        caches.open('oraex-store').then((cache) => cache.addAll([
            './index.html',
            './oraex_logo.png'
        ]))
    );
});

self.addEventListener('fetch', (e) => {
    e.respondWith(
        caches.match(e.request).then((response) => response || fetch(e.request))
    );
});
