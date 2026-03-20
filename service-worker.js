const CACHE_NAME = "brajwasi-v3";

const ASSETS = [
    "/",
    "/entry",
    "/static/style.css",
    "/static/icons/icon-192.png",
    "/manifest.json"
];

self.addEventListener("install", event => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
    );
});

self.addEventListener("activate", event => {
    event.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.map(key => {
                    if (key !== CACHE_NAME) {
                        return caches.delete(key);
                    }
                })
            );
        })
    );
    self.clients.claim();
});

self.addEventListener("fetch", event => {
    // Network first, fall back to cache
    event.respondWith(
        fetch(event.request)
            .then(response => {
                // Cache successful GET responses
                if (event.request.method === "GET" && response.status === 200) {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
                }
                return response;
            })
            .catch(() => caches.match(event.request))
    );
});