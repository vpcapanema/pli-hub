const CACHE_VERSION = "pli-hub-v1";

const PRECACHE = [
  "/pli-hub/",
  "/pli-hub/index.html",
  "/pli-hub/manifest.webmanifest",
  "/pli-hub/icons/icon-192.png",
  "/pli-hub/icons/icon-512.png",
  "/pli-hub/relatorios_d11.html",
  "/pli-hub/aderencia_d11c_09_e_d11c_18_tdr_pli.html",
  "/pli-hub/sumario_resultados_temas_d11.html",
  "/pli-hub/carteira_projetos_d11c_09_18.html",
  "/pli-hub/sistematizacao_temas_d11_pli.html",
  "/pli-hub/resumo_metodologia_temas_d11.html",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(PRECACHE)),
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys.filter((key) => key !== CACHE_VERSION).map((key) => caches.delete(key)),
        ),
      )
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const networkFetch = fetch(event.request)
        .then((response) => {
          if (response && response.status === 200 && response.type === "basic") {
            const copy = response.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(event.request, copy));
          }
          return response;
        })
        .catch(() => cached);

      return cached || networkFetch;
    }),
  );
});
