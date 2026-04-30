const CACHE_NAME = 'pos-cache-v10'; // Сейчас мы на 6-й версии!

const urlsToCache = [
    './pos2.html',
    './manifest.json',
    './config.js'
];

// Установка: скачиваем файлы и сразу активируем
self.addEventListener('install', event => {
    self.skipWaiting(); 
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
    );
});

// Активация: чистим старый кэш и МГНОВЕННО берем управление
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME) {
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(() => self.clients.claim()) // Это уберет надпись v: dev
    );
});

// Перехват запросов (работа в оффлайне)
self.addEventListener('fetch', event => {
    event.respondWith(
        caches.match(event.request).then(response => {
            return response || fetch(event.request);
        })
    );
});

// Ответ на запрос версии от pos2.html
self.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'GET_VERSION') {
        event.ports[0].postMessage({ version: CACHE_NAME });
    }
});
