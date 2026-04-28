const CACHE_NAME = "ai-workbench-v2";
const APP_SHELL = ["/", "/manifest.webmanifest", "/pwa-icon.svg"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)),
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))),
    ),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const url = new URL(event.request.url);

  if (url.origin !== self.location.origin || url.pathname.startsWith("/rest/")) {
    event.respondWith(fetch(event.request, { cache: "no-store" }));
    return;
  }

  event.respondWith(
    fetch(event.request, { cache: "no-store" })
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            if (response.ok) cache.put(event.request, copy);
          });
          return response;
        })
        .catch(() => caches.match(event.request).then((cached) => cached || caches.match("/"))),
  );
});
