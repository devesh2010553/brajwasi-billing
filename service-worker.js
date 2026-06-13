const CACHE_NAME = "brajwasi-v4";

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
    caches.keys().then(keys =>
      Promise.all(keys.map(key => key !== CACHE_NAME && caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  event.respondWith(
    fetch(event.request)
      .then(response => {
        if (event.request.method === "GET" && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});

// ── Push Notification Handler ────────────────────────────────────────────────
self.addEventListener("push", event => {
  let data = { title: "Brajwasi Travels", body: "New message from admin" };
  try {
    data = event.data.json();
  } catch(e) {
    data.body = event.data ? event.data.text() : "New alert";
  }

  event.waitUntil(
    self.registration.showNotification(data.title, {
      body:    data.body,
      icon:    "/static/icons/icon-192.png",
      badge:   "/static/icons/icon-192.png",
      vibrate: [200, 100, 200],
      tag:     "brajwasi-alert",
      renotify: true,
      data:    { url: "/entry" }
    })
  );
});

self.addEventListener("notificationclick", event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({type: "window"}).then(clientList => {
      for (const client of clientList) {
        if (client.url.includes("/entry") && "focus" in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow("/entry");
    })
  );
});