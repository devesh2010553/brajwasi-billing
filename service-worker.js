const CACHE_NAME = "brajwasi-v5";

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
  const url = event.request.url;

  // ── Skip non-http requests (chrome-extension://, data:, etc.) ──
  if (!url.startsWith("http://") && !url.startsWith("https://")) return;

  // ── Skip POST/non-GET requests ──
  if (event.request.method !== "GET") return;

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Only cache valid same-origin responses
        if (
          response &&
          response.status === 200 &&
          response.type !== "opaque" &&
          url.startsWith(self.location.origin)
        ) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});

// ── Push Notification Handler ─────────────────────────────────────────────
self.addEventListener("push", event => {
  let data = { title: "Brajwasi Travels 🚗", body: "New message from admin" };
  try { data = event.data.json(); }
  catch(e) { data.body = event.data ? event.data.text() : "New alert"; }

  event.waitUntil(
    self.registration.showNotification(data.title, {
      body:     data.body,
      icon:     "/static/icons/icon-192.png",
      badge:    "/static/icons/icon-192.png",
      vibrate:  [200, 100, 200],
      tag:      "brajwasi-alert",
      renotify: true,
      data:     { url: "/entry" }
    })
  );
});

self.addEventListener("notificationclick", event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: "window" }).then(list => {
      for (const client of list) {
        if (client.url.includes("/entry") && "focus" in client) return client.focus();
      }
      if (clients.openWindow) return clients.openWindow("/entry");
    })
  );
});