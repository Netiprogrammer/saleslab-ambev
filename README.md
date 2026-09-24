# SalesLab · Ambev — Command Center Operacional

Painel operacional 100% client-side (HTML + CSS + Vanilla JS) para consolidar a base
mestre de equipamentos/faturamento/status de SKUs por PDV. Não existe backend: o
Excel é lido, processado e descartado inteiramente na memória do navegador. Pode ser
hospedado direto no GitHub Pages.

## Rodando localmente

Basta abrir `index.html` num servidor estático (ou usar a extensão "Live Server").
**Não abra via `file://` direto** — o Web Worker de ingestão (`worker.js`) e o
Service Worker (`sw.js`) exigem `http://`/`https://`. Sirva a pasta com qualquer
servidor estático (`npx http-server`, `python -m http.server`, GitHub Pages etc.).

### Ao desenvolver: cuidado com o cache do Service Worker

Depois da primeira visita, o Service Worker passa a servir os arquivos do cache
antes da rede (é o que permite funcionar offline). Isso significa que, se você
editar `index.html`/`style.css`/`script.js`/`etl.js`/`worker.js` localmente, o
navegador pode continuar mostrando a versão antiga até você **subir o número em
`CACHE_NAME` no topo do `sw.js`** (ex.: `v1` → `v2`) — isso invalida o cache
antigo e força buscar tudo de novo. Em desenvolvimento, também dá pra abrir o
DevTools → Application → Service Workers → "Bypass for network"/"Unregister".

## Rede corporativa restrita (sem CDN)

O app tenta carregar as bibliotecas de `vendor/` primeiro e só usa a CDN se o
arquivo local não existir. Para deixar tudo 100% offline, baixe as duas libs uma
vez e commite a pasta `vendor/`:

```powershell
New-Item -ItemType Directory -Force vendor | Out-Null
Invoke-WebRequest -Uri "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js" -OutFile "vendor/xlsx.full.min.js"
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js" -OutFile "vendor/chart.umd.min.js"
```

## Como o ETL lê a planilha

1. **Seleção de aba**: procura por abas cujo nome contenha `BI DE EQUIPAMENTOS`,
   `VISIBILIDADE` ou `SKU-PDV`; usa `VISÃO GERENCIAL` como apoio; senão, a primeira aba.
2. **Cabeçalho inteligente**: varre as 15 primeiras linhas e usa a primeira que
   contiver pelo menos duas das palavras `PDV`, `STATUS`, `SETOR` — ignora
   linhas de título/data acima do cabeçalho real.
3. **Fuzzy matching sem depender de ordem**: cada campo (PDV, Nome, Setor,
   Responsável, Status Geral, Status SKU, Faturamento Real/Esperado) tem uma
   lista de palavras-chave por prioridade. Todas as combinações campo×coluna são
   pontuadas e a atribuição é feita globalmente (maior prioridade e match mais
   específico vencem primeiro), então uma coluna como "Status PDV" não rouba a
   coluna de código do PDV só por conter a substring "PDV".
4. **Linha de TOTAL** é descartada automaticamente.

Se a detecção automática errar, use o botão **Diagnóstico** no cabeçalho: ele
mostra qual coluna foi usada para cada campo (com uma amostra de valores reais
lidos, pra você confirmar sem precisar abrir a planilha) e permite trocar
manualmente — a escolha fica salva no navegador (`localStorage`) e é reaplicada
em cargas futuras. Se o cabeçalho não for identificado com confiança, um aviso
aparece tanto no Diagnóstico quanto em toast.

## Arquitetura: Web Worker + PWA offline

- **Ingestão fora da thread principal**: `worker.js` (com `etl.js` compartilhado
  via `importScripts`) faz todo o parse do Excel e o ETL num Web Worker, então a
  interface nunca trava durante a leitura de um arquivo grande — e o overlay de
  carregamento mostra o progresso real por etapa (lendo → detectando aba →
  mapeando colunas → processando linha X de Y → calculando indicadores).
- **100% offline depois da primeira visita**: `manifest.json` + `sw.js` cacheiam
  o app inteiro (HTML/CSS/JS + as libs de CDN) via Service Worker, então o app
  funciona e é instalável mesmo sem rede nenhuma — reforça a mesma proposta do
  `vendor/` para rede corporativa restrita. Veja a seção acima sobre versionar
  `CACHE_NAME` ao atualizar os arquivos.

## Atalhos e recursos de uso diário

- **`/`** foca a busca na Auditoria de PDVs; **Esc** fecha qualquer painel/modal
  aberto (Raio-X, Diagnóstico, modal de cópia); clicar fora de um modal também
  fecha.
- Cabeçalhos da tabela de Auditoria (PDV, Nome, Setor, Status Geral, Status SKU)
  são clicáveis para ordenar.
- **Exportar CSV** na Auditoria baixa exatamente a lista filtrada/ordenada na tela.
- **Limpar base** (aparece após carregar um arquivo) remove os dados da memória
  sem precisar dar F5 — não mexe no tema nem no mapeamento salvo.
- Os KPIs mostram a variação (▲/▼) em relação à última carga (comparação salva
  em `localStorage`, não é histórico de dias).

## Limitações conhecidas

- O gráfico do Raio-X é **ilustrativo/simulado** (determinístico por PDV), não
  reflete histórico real — não há coluna de série mensal mapeada ainda.
- A tabela de Auditoria renderiza no máximo 200 linhas por vez (botão
  "Mostrar mais" carrega mais) para não travar a DOM em bases grandes.
- Arquivo público: **não commite planilhas reais** (veja `.gitignore`).
