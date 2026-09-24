# SalesLab Ops: Central Executiva Ambev

Painel operacional 100% estático (HTML + CSS + JS puro). O Excel é lido **na memória do navegador**: não há backend, banco de dados nem envio de arquivo para nenhum servidor.

## Como usar

1. Abra a página (GitHub Pages) e clique em **Carregar arquivo mestre**, ou arraste o `.xlsm`/`.xlsx`/`.xls` para a janela.
2. **Visão gerencial**: agrupa por setor, com representante, equipamentos, Giro OK, atingimento (verde a partir da meta) e GAPs. Clique num setor para ver só os PDVs dele.
3. **Auditoria de PDVs**: busca por código ou nome fantasia, filtro por status e 200 linhas por vez ("Mostrar mais" revela o resto).
4. **Auditar** abre o Raio-X do PDV: dados do cliente, faturamento esperado x real, curva de cobertura e o botão **Copiar pauta de cobrança** (pronta para colar no Teams ou no WhatsApp).
5. **Diagnóstico** (topo) mostra qual aba e qual linha de cabeçalho foram usadas, qual coluna virou cada campo e como cada valor de status foi classificado.

## Publicar no GitHub Pages

`Settings > Pages > Build and deployment > Deploy from a branch > main / (root)`.

## Funcionar em rede corporativa que bloqueia CDN

O app carrega SheetJS e Chart.js primeiro da pasta `vendor/` e só depois da CDN. Baixe uma vez e faça commit da pasta:

```powershell
mkdir vendor
Invoke-WebRequest https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js -OutFile vendor/xlsx.full.min.js
Invoke-WebRequest https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.js -OutFile vendor/chart.umd.js
```

Ícones (Remixicon) e fonte (Inter) são opcionais: se forem bloqueados, o app funciona com a fonte do sistema e os botões continuam com texto.

## Como o Excel é lido

- **Aba**: prefere a que tem "Visibilidade" ou "BI de Equipamentos" no nome; senão usa a primeira. Dá para trocar no Diagnóstico.
- **Cabeçalho**: varre as 40 primeiras linhas e escolhe a que tem várias células de texto e termos como PDV, STATUS, SETOR, GAP. Linhas de título e de data (`DATA: 23/09/2026`) são puladas.
- **Colunas**: os nomes são normalizados (sem acento, quebras de linha viram espaço) e cada campo procura o padrão mais específico primeiro, então a ordem das colunas no arquivo não importa. Se algo vier errado, escolha a coluna certa no Diagnóstico; a escolha fica salva no navegador.
- **Status**: `NOK` nunca conta como `OK`. Para incluir outras palavras do BI, edite a lista `CLASSES` no topo do `script.js`.
- Linhas de total do relatório são descartadas.

## Privacidade

O repositório é público. **Nunca faça commit das planilhas** (o `.gitignore` já bloqueia `.xlsx`, `.xlsm`, `.xls` e `.csv`). Todo texto vindo do Excel é escapado antes de ir para a tela.

## Limitações conhecidas

- A curva de cobertura do Raio-X é **simulada** (estável por PDV) e vem marcada como ilustrativa; ainda não usa histórico real.
- Cabeçalhos de dois níveis (células mescladas em duas linhas) não são combinados.
- Arquivo protegido por senha não é aberto; salve uma cópia sem senha.
