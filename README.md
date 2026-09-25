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

## Como o ETL lê a planilha (múltiplas abas, não só uma)

O app não escolhe "a melhor aba" e ignora o resto — ele lê **todas as abas
relevantes** e mescla os dados por PDV, porque numa base real (BI corporativo)
a informação vem espalhada: uma aba tem o status, outra tem o faturamento,
outra tem o responsável.

1. **Seleção de abas**: qualquer aba cujo nome contenha `BI DE EQUIPAMENTOS`,
   `VISIBILIDADE` ou `SKU-PDV` entra na leitura "geral" (Giro/SKU por PDV); uma
   aba com `CHOPEIRA` no nome entra como uma dimensão extra (equipamento de
   chope, com status/meta/faturamento próprios). Se nenhuma aba "geral" for
   encontrada, cai no fallback `VISÃO GERENCIAL` ou na primeira aba do arquivo.
2. **Cabeçalho inteligente** (por aba): varre as 15 primeiras linhas e usa a
   primeira que contiver pelo menos duas das palavras `PDV`, `STATUS`, `SETOR`
   — ignora linhas de título/data acima do cabeçalho real.
3. **Fuzzy matching sem depender de ordem** (por aba): cada campo tem uma lista
   de palavras-chave por prioridade; todas as combinações campo×coluna são
   pontuadas e a atribuição é feita globalmente, então uma coluna como "Status
   PDV" não rouba a coluna de código do PDV só por conter a substring "PDV".
4. **Mesclagem por PDV**: as abas "geral" são processadas em ordem de
   prioridade (`BI DE EQUIPAMENTOS` → `VISIBILIDADE` → `SKU-PDV`) e cada campo
   é preenchido pela **primeira** aba que realmente tiver aquela coluna — uma
   aba processada depois só completa o que falta, nunca sobrescreve um valor
   (inclusive zero legítimo, tipo "Faturamento Real = 0" numa Venda Zero) que
   uma aba melhor já preencheu. A aba de Chopeira entra como campos extras
   (`statusChopeira`, `faturamentoChopeira`, `metaChopeira`, `gapChopeira`),
   mostrados no Raio-X só quando aquele PDV tiver esse dado.
5. **Linha de TOTAL** é descartada automaticamente em cada aba.
6. **Números lidos direto da célula**, não do texto formatado — evita que uma
   coluna sem casas decimais ou com separador de milhar arredonde ou distorça
   o valor real (ver Limitações abaixo pra mais contexto).

Abra o botão **Diagnóstico** no cabeçalho pra ver, aba por aba: qual coluna foi
usada em cada campo, uma amostra de valores reais lidos (pra confirmar sem
abrir a planilha), e quantos PDVs vieram de mais de uma aba. Dá pra trocar
manualmente o mapeamento de qualquer campo em qualquer aba — só o que você
realmente mudar vira override salvo (`localStorage`); o resto continua se
adaptando automaticamente na próxima carga. Se o cabeçalho de alguma aba não
for identificado com confiança, um aviso aparece no Diagnóstico e em toast.

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
- **Arquivos grandes/complexos (dezenas de MB, muitas abas) podem levar de 20
  a 30 segundos pra processar.** Isso é o custo real de descompactar o arquivo
  e ler a tabela de strings compartilhadas do Excel (inerente ao formato, não
  dá pra pular) — mas a interface não trava nesse tempo (roda num Web Worker) e
  o overlay mostra o progresso por etapa/aba em vez de parecer travado.
- **Se a mesma aba tiver o PDV repetido em mais de uma linha** (ex.: duas
  "tabelas" coladas lado a lado na mesma aba, ou uma exportação com histórico
  de vários meses), o app mantém os valores da **primeira ocorrência** e ignora
  as demais para aquele PDV — não tenta adivinhar qual linha é "a certa". Se a
  sua base tiver esse padrão, mais vale limpar a aba de origem do que confiar
  no comportamento automático.
- Arquivo público: **não commite planilhas reais** (veja `.gitignore`).
