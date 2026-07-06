const swPath = new URL(self.location.href).pathname;
const profileMatch = swPath.match(/^\/mobile\/(team|personal)\//);
const APP_BASE = profileMatch ? `/mobile/${profileMatch[1]}` : "/mobile";
const CACHE_NAME = `lucas-mobile-shell-v2-${profileMatch ? profileMatch[1] : "default"}`;
const APP_SHELL = [
  APP_BASE,
  `${APP_BASE}/`,
  `${APP_BASE}/index.html`,
  `${APP_BASE}/styles.css`,
  `${APP_BASE}/app.js`,
  `${APP_BASE}/manifest.webmanifest`,
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  const url = new URL(request.url);
  if (request.method !== "GET" || !url.pathname.startsWith(APP_BASE)) {
    return;
  }
  if (url.pathname.startsWith("/mobile/api/")) {
    return;
  }
  event.respondWith(
    caches.match(request).then((cached) => (
      cached || fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      })
    ))
  );
});
