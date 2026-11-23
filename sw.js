// Service worker: guarda o aplicativo no aparelho para abrir mesmo sem internet.
const VERSAO = 'sorteio-un-web-1.0.0';
const ARQUIVOS = [
  './', 'index.html', 'style.css', 'mobile.css', 'manifest.webmanifest',
  'api-web.js', 'renderer.js', 'historico.js', 'fundo.js', 'vertical.js', 'logo.js', 'video.js', 'tema.js',
  'vendor/xlsx.full.min.js', 'vendor/mp4-muxer.js',
  'assets/emblema.png', 'assets/abertura.webp',
  'fonts/outfit-latin-400-normal.woff2', 'fonts/outfit-latin-500-normal.woff2', 'fonts/outfit-latin-600-normal.woff2',
  'fonts/outfit-latin-700-normal.woff2', 'fonts/outfit-latin-800-normal.woff2',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png', 'icons/apple-touch-icon.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSAO).map((k) => caches.delete(k)))).then(() => self.clients.claim()),
  );
});

// Abre o que está guardado e, em segundo plano, atualiza para a próxima vez (só arquivos do próprio app).
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then((guardado) => {
      const rede = fetch(req).then((r) => {
        if (r && r.ok) { const copia = r.clone(); caches.open(VERSAO).then((c) => c.put(req, copia)); }
        return r;
      }).catch(() => guardado);
      return guardado || rede;
    }),
  );
});
