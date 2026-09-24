# SalesLab · Ambev — Command Center Operacional

Painel operacional 100% client-side (HTML + CSS + Vanilla JS) para consolidar a base
mestre de equipamentos/faturamento/status de SKUs por PDV. Não existe backend: o
Excel é lido, processado e descartado inteiramente na memória do navegador. Pode ser
hospedado direto no GitHub Pages.

## Rodando localmente

Basta abrir `index.html` num servidor estático (ou usar a extensão "Live Server").
Abrir via `file://` direto também funciona, mas alguns navegadores restringem
`fetch`/Workers nesse modo — preferir sempre um servidor local ou o GitHub Pages.

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
mostra qual coluna foi usada para cada campo e permite trocar manualmente — a
escolha fica salva no navegador (`localStorage`) e é reaplicada em cargas futuras.

## Limitações conhecidas

- O gráfico do Raio-X é **ilustrativo/simulado** (determinístico por PDV), não
  reflete histórico real — não há coluna de série mensal mapeada ainda.
- A tabela de Auditoria renderiza no máximo 200 linhas por vez (botão
  "Mostrar mais" carrega mais) para não travar a DOM em bases grandes.
- Arquivo público: **não commite planilhas reais** (veja `.gitignore`).
