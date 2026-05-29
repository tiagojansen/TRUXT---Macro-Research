// TRUXT Macro Research — Service Worker (network-first)
// Sempre tenta buscar do servidor; cache e so fallback offline.
const CACHE_NAME = 'truxt-v1';

// Instala imediatamente, sem esperar a aba fechar
self.addEventListener('install', () => self.skipWaiting());

// Assume controle de todas as abas abertas imediatamente
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// Network-first para todos os GETs
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request)
      .then(res => {
        // Armazena resposta no cache para fallback offline
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
