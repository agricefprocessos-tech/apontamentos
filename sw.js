// ============================================================================
// sw.js — Service Worker cujo único trabalho é reenviar a fila offline de apontamentos em
// segundo plano via Background Sync, mesmo com o app fechado (funciona no Android/Chrome;
// o iOS/Safari não implementa a Background Sync API — nesse caso o registro do SW ainda
// funciona normalmente, só o `sync.register()` nunca chega a ser chamado, ver index.html).
//
// De propósito, NÃO tem handler de `fetch` — não cacheia nada. O único trabalho daqui é
// sincronizar a fila; adicionar cache de assets abriria uma segunda fonte de "versão
// desatualizada do app" em cima do problema que este arquivo existe pra resolver.
// ============================================================================

importScripts('agricef-fila-shared.js');

self.addEventListener('install', () => {
  self.skipWaiting(); // sem isso a primeira instalação só assume controle na próxima navegação
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('sync', (event) => {
  if (event.tag !== 'sync-fila') return;
  event.waitUntil(_executarSyncEmBackground());
});

async function _executarSyncEmBackground() {
  const hooks = {
    onSincronizado: async (n) => {
      const clientes = await self.clients.matchAll();
      clientes.forEach(c => c.postMessage({ tipo: 'agricef-fila-sincronizada', quantidade: n }));
    },
    onFalhaPermanente: async (item, motivo) => {
      const clientes = await self.clients.matchAll();
      clientes.forEach(c => c.postMessage({ tipo: 'agricef-fila-falha-permanente', item, motivo }));
    },
  };
  const resultado = await _sincronizarFilaImplCore(false, hooks);
  if (resultado.restantes > 0) {
    // Re-registra defensivamente — o comportamento de retry interno do navegador pra uma tag
    // de sync presa não tem garantia documentada de tentar de novo sozinho.
    try { await self.registration.sync.register('sync-fila'); } catch (e) {}
  }
}
