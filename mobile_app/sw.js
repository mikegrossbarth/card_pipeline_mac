const swPath = new URL(self.location.href).pathname;
const profileMatch = swPath.match(/^\/mobile\/(team|personal)\//);
const APP_BASE = profileMatch ? `/mobile/${profileMatch[1]}` : "/mobile";
const CACHE_NAME = `lucas-mobile-shell-v6-${profileMatch ? profileMatch[1] : "default"}`;
const APP_SHELL = [
  APP_BASE,
  `${APP_BASE}/`,
  `${APP_BASE}/index.html`,
  `${APP_BASE}/styles.css`,
  `${APP_BASE}/app.js`,
  `${APP_BASE}/manifest.webmanifest`,
];
const OFFLINE_HTML = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <meta name="theme-color" content="#101820">
    <title>LUCAS Mobile Offline</title>
    <style>
      :root { color-scheme: dark; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #101820; color: #eef5f6; }
      body { margin: 0; min-height: 100vh; display: grid; place-items: center; padding: 24px; background: #101820; }
      main { max-width: 560px; padding: 18px; border: 1px solid #2d3d4a; border-radius: 8px; background: #17232d; }
      h1 { margin: 0 0 8px; font-size: 24px; }
      p { margin: 8px 0 0; color: #b8c7ce; line-height: 1.4; }
    </style>
  </head>
  <body>
    <main>
      <h1>LUCAS Mobile is offline</h1>
      <p>The app shell was not fully cached on this phone yet. Open LUCAS Mobile once while connected to desktop LUCAS, then it can reopen from your home screen without Wi-Fi.</p>
    </main>
  </body>
</html>`;

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => Promise.all(APP_SHELL.map((url) => fetch(url).then((response) => {
        if (response.ok) {
          return cache.put(url, response);
        }
        return Promise.resolve();
      }).catch(() => Promise.resolve()))))
      .then(() => self.skipWaiting())
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
  if (url.pathname.startsWith("/mobile/api/") || url.pathname.startsWith(`${APP_BASE}/api/`)) {
    return;
  }
  event.respondWith(
    fetch(request).then((response) => {
      const copy = response.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      return response;
    }).catch(() => {
      if (request.mode === "navigate") {
        return caches.match(`${APP_BASE}/index.html`)
          .then((cached) => cached || caches.match(`${APP_BASE}/`))
          .then((cached) => cached || new Response(OFFLINE_HTML, { headers: { "content-type": "text/html; charset=utf-8" } }));
      }
      return caches.match(request)
        .then((cached) => cached || caches.match(`${APP_BASE}/${url.pathname.split("/").pop()}`))
        .then((cached) => cached || caches.match(`/mobile/${url.pathname.split("/").pop()}`))
        .then((cached) => cached || new Response("", { status: 504, statusText: "Offline" }));
    })
  );
});
