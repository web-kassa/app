const CACHE_NAME = 'pos-v32'; // <--- ВОТ ТУТ ПИШЕТСЯ ВЕРСИЯ, КОТОРУЮ ПОКАЖЕТ КАССА

const ASSETS = [
  './',
  './index.html',
  './pos2.html',
  './config.js',
  './manifest.json',
  './icon.png' // <--- Вот она, разгадка! Мы указали правильное имя.
];

self.addEventListener('install', e => {
  self.skipWaiting(); // Заставляем обновиться без ожидания
  e.waitUntil(
    caches.open(CACHE_NAME).then(c => c.addAll(ASSETS))
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(self.clients.claim()); // Мгновенно берем кассу под контроль!
  
  e.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    ))
  );
});

self.addEventListener('fetch', e => {
  if (e.request.url.includes('script.google.com')) {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    return;
  }
  
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});

// ТОТ САМЫЙ БЛОК, КОТОРЫЙ ОТВЕЧАЕТ НА ВОПРОС КАССЫ "КАКАЯ ВЕРСИЯ?"
self.addEventListener('message', event => {
    if (event.data && event.data.type === 'GET_VERSION') {
        event.ports[0].postMessage({ version: CACHE_NAME });
    }
});
