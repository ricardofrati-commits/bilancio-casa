/* Service worker: fa funzionare l'app anche senza internet. */
const V = 'bilancio-casa-v1';
const CORE = [
  './', 'index.html', 'style.css', 'app.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png',
  'vendor/fonts/bricolage.woff2', 'vendor/exceljs.min.js',
  'vendor/pdf.min.js', 'vendor/pdf.worker.min.js',
  'vendor/tesseract/tesseract.min.js', 'vendor/tesseract/worker.min.js',
  'vendor/tesseract/core/tesseract-core-lstm.wasm.js',
  'vendor/tesseract/core/tesseract-core-simd-lstm.wasm.js',
  'vendor/tesseract/lang/ita.traineddata.gz'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => Promise.allSettled(CORE.map(u => c.add(u)))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== location.origin) return;
  if (u.pathname.includes('/vendor/')) {            // librerie pesanti: prima la copia salvata
    e.respondWith(caches.match(r).then(hit => hit || fetch(r).then(res => {
      const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); return res; })));
    return;
  }
  e.respondWith(fetch(r).then(res => {              // pagina e codice: prima la rete, poi la copia
    const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); return res;
  }).catch(() => caches.match(r, { ignoreSearch: true }).then(h => h || caches.match('index.html'))));
});
