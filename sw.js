'use strict';

/**
 * Service Worker do SalesLab · Ambev — deixa o app instalável e utilizável mesmo
 * sem rede nenhuma (depois da primeira visita), reforçando a proposta 100% client-side.
 * Sobe a versão do CACHE_NAME sempre que mudar algum arquivo do app shell.
 */
const CACHE_NAME = 'saleslab-ambev-v1';

const ARQUIVOS_APP = [
  'index.html',
  'style.css',
  'script.js',
  'etl.js',
  'worker.js',
  'manifest.json',
  'icon.svg',
];

const ARQUIVOS_EXTERNOS = [
  'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap',
  'https://cdn.jsdelivr.net/npm/remixicon@4.5.0/fonts/remixicon.css',
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js',
];

self.addEventListener('install', (evento) => {
  evento.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(ARQUIVOS_APP);
    // Bibliotecas de terceiros: cacheia best-effort (respostas opacas no-cors),
    // sem derrubar a instalação inteira se a rede bloquear alguma CDN agora.
    await Promise.all(ARQUIVOS_EXTERNOS.map(async (url) => {
      try {
        const resposta = await fetch(url, { mode: 'no-cors' });
        await cache.put(url, resposta);
      } catch {
        /* sem rede agora; será cacheada em runtime na próxima vez que carregar com sucesso */
      }
    }));
    self.skipWaiting();
  })());
});

self.addEventListener('activate', (evento) => {
  evento.waitUntil((async () => {
    const nomes = await caches.keys();
    await Promise.all(nomes.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)));
    self.clients.claim();
  })());
});

self.addEventListener('fetch', (evento) => {
  if (evento.request.method !== 'GET') return;
  evento.respondWith((async () => {
    const emCache = await caches.match(evento.request);
    if (emCache) return emCache;
    try {
      const resposta = await fetch(evento.request);
      const cache = await caches.open(CACHE_NAME);
      cache.put(evento.request, resposta.clone());
      return resposta;
    } catch {
      return emCache || Response.error();
    }
  })());
});
