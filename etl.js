'use strict';

/**
 * ETL puro (sem DOM) — compartilhado entre a thread principal e o Web Worker de ingestão.
 * Carregado via <script src="etl.js"> na página e via importScripts('etl.js') no worker;
 * por isso expõe tudo em `self.ETL` (em vez de `const`/`let` de topo, que não vira global).
 */
(function (global) {
  const CONFIG = {
    sheetKeywordsPrimary: ['BIDEEQUIPAMENTOS', 'VISIBILIDADE', 'SKUPDV'],
    sheetKeywordsSupport: ['VISAOGERENCIAL'],
    headerScanRows: 15,
    headerKeywords: ['PDV', 'STATUS', 'SETOR'],
    minHeaderKeywordMatches: 2,
    amostrasPorCampo: 3,
    linhasPorLoteProgresso: 5000,
    // Prioridade das palavras-chave por campo (índice menor = mais prioritário).
    fieldKeywords: {
      pdv: ['PDV', 'COD', 'CLIENTE'],
      nome: ['RAZAO', 'NOME', 'FANTASIA'],
      setor: ['GV', 'SETOR', 'CODSETOR'],
      responsavel: ['SUPERCOM', 'COMERCIAL', 'RN', 'DONO'],
      statusGeral: ['STATUSPDV', 'STATUSDO', 'GIRO'],
      statusSku: ['STATUSSKU', 'SKU'],
      faturamentoReal: ['FATURAMENTOREAL', 'REAL'],
      faturamentoEsperado: ['FATURAMENTOESPERADO', 'ESPERADO'],
    },
  };

  /** Remove acentos, espaços, quebras de linha e pontuação; deixa tudo maiúsculo. */
  function normalizeKey(str) {
    return String(str ?? '')
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-zA-Z0-9]/g, '')
      .toUpperCase();
  }

  function parseNumero(valor) {
    if (valor === undefined || valor === null || valor === '') return 0;
    if (typeof valor === 'number') return valor;
    const limpo = String(valor)
      .replace(/[^\d,.-]/g, '')
      .replace(/\.(?=\d{3}(\D|$))/g, '')
      .replace(',', '.');
    const numero = parseFloat(limpo);
    return Number.isFinite(numero) ? numero : 0;
  }

  function encontrarAbaAlvo(workbook) {
    const nomes = workbook.SheetNames;
    const normalizados = nomes.map(normalizeKey);
    let idx = normalizados.findIndex((n) => CONFIG.sheetKeywordsPrimary.some((kw) => n.includes(kw)));
    if (idx === -1) idx = normalizados.findIndex((n) => CONFIG.sheetKeywordsSupport.some((kw) => n.includes(kw)));
    if (idx === -1) idx = 0;
    return nomes[idx];
  }

  /**
   * Varre as N primeiras linhas em busca da linha que parece ser o cabeçalho real (ignora título/data).
   * `confianca: 'baixa'` sinaliza que nenhuma linha bateu o mínimo de palavras-chave e caímos no fallback
   * (linha 0) — o app usa isso pra avisar o usuário em vez de assumir silenciosamente que está certo.
   */
  function encontrarLinhaCabecalho(matriz) {
    const limite = Math.min(CONFIG.headerScanRows, matriz.length);
    for (let i = 0; i < limite; i++) {
      const linha = matriz[i] || [];
      const textoLinha = linha.map(normalizeKey).join(' ');
      const matches = CONFIG.headerKeywords.filter((kw) => textoLinha.includes(kw)).length;
      if (matches >= CONFIG.minHeaderKeywordMatches) return { indice: i, confianca: 'alta' };
    }
    return { indice: 0, confianca: 'baixa' };
  }

  /** Pontua o quanto uma chave de coluna (já normalizada) combina com uma lista de palavras-chave. */
  function pontuarColuna(chaveNormalizada, keywords) {
    let melhor = null;
    keywords.forEach((kw, idx) => {
      const kwNorm = normalizeKey(kw);
      if (kwNorm && chaveNormalizada.includes(kwNorm)) {
        if (!melhor || idx < melhor.idx || (idx === melhor.idx && kwNorm.length > melhor.kwLen)) {
          melhor = { idx, kwLen: kwNorm.length };
        }
      }
    });
    return melhor;
  }

  /**
   * Mapeia campo -> nome original da coluna usando um leilão global (não greedy por campo):
   * todas as combinações (campo, coluna) são pontuadas e ordenadas por prioridade + especificidade,
   * e cada coluna só pode ser usada uma vez. Isso evita que "Status PDV" roube a coluna do campo
   * "pdv" só porque contém a substring "PDV" — o match mais específico (e de maior prioridade) vence.
   */
  function mapearColunas(headers) {
    const candidatos = [];
    Object.entries(CONFIG.fieldKeywords).forEach(([campo, keywords]) => {
      headers.forEach((header) => {
        const chaveNormalizada = normalizeKey(header);
        if (!chaveNormalizada) return;
        const pontuacao = pontuarColuna(chaveNormalizada, keywords);
        if (pontuacao) {
          candidatos.push({
            campo,
            header,
            score: pontuacao.idx,
            especificidade: pontuacao.kwLen,
            tamanhoColuna: chaveNormalizada.length,
          });
        }
      });
    });

    candidatos.sort((a, b) =>
      a.score - b.score ||
      b.especificidade - a.especificidade ||
      a.tamanhoColuna - b.tamanhoColuna
    );

    const mapa = {};
    const colunasUsadas = new Set();
    candidatos.forEach(({ campo, header }) => {
      if (mapa[campo] || colunasUsadas.has(header)) return;
      mapa[campo] = header;
      colunasUsadas.add(header);
    });
    return mapa;
  }

  /** Helper genérico (uso pontual/manual): acha o valor de uma linha cujo cabeçalho combina com keywords. */
  function extrairColuna(row, keywords) {
    let melhorChave = null;
    let melhorPontuacao = null;
    Object.keys(row).forEach((chave) => {
      const pontuacao = pontuarColuna(normalizeKey(chave), keywords);
      if (pontuacao && (!melhorPontuacao || pontuacao.idx < melhorPontuacao.idx)) {
        melhorPontuacao = pontuacao;
        melhorChave = chave;
      }
    });
    return melhorChave ? row[melhorChave] : undefined;
  }

  /**
   * @param {object} worksheet planilha do SheetJS
   * @param {object} overridesSalvos mapeamento manual salvo pelo usuário no Diagnóstico ({campo: header})
   * @param {(atual:number, total:number)=>void} [onProgresso] chamado a cada lote de linhas (barra de progresso real)
   */
  function construirRegistros(worksheet, overridesSalvos = {}, onProgresso) {
    const matriz = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    const { indice: linhaCabecalho, confianca: confiancaCabecalho } = encontrarLinhaCabecalho(matriz);
    const headers = (matriz[linhaCabecalho] || []).map((h) => String(h || '').trim()).filter(Boolean);

    const mapaAutomatico = mapearColunas(headers);
    // Overrides salvos só valem se a coluna ainda existir neste arquivo.
    const mapa = { ...mapaAutomatico };
    Object.entries(overridesSalvos).forEach(([campo, header]) => {
      if (headers.includes(header)) mapa[campo] = header;
    });

    const registros = [];
    const amostras = {};
    const totalLinhas = matriz.length - (linhaCabecalho + 1);

    for (let i = linhaCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      if (onProgresso && (i - linhaCabecalho) % CONFIG.linhasPorLoteProgresso === 0) {
        onProgresso(i - linhaCabecalho, totalLinhas);
      }
      if (!linha || linha.every((c) => String(c ?? '').trim() === '')) continue;

      const row = {};
      headers.forEach((h, idx) => { row[h] = linha[idx]; });

      const pdv = String(row[mapa.pdv] ?? '').trim();
      const nome = String(row[mapa.nome] ?? '').trim();
      if (!pdv && !nome) continue;
      if (normalizeKey(pdv).includes('TOTAL') || normalizeKey(nome).includes('TOTAL')) continue;

      Object.keys(mapa).forEach((campo) => {
        const valor = String(row[mapa[campo]] ?? '').trim();
        if (!valor) return;
        if (!amostras[campo]) amostras[campo] = [];
        if (amostras[campo].length < CONFIG.amostrasPorCampo && !amostras[campo].includes(valor)) {
          amostras[campo].push(valor);
        }
      });

      registros.push({
        pdv,
        nome,
        setor: String(row[mapa.setor] ?? '').trim() || '—',
        responsavel: String(row[mapa.responsavel] ?? '').trim() || '—',
        statusGeral: String(row[mapa.statusGeral] ?? '').trim(),
        statusSku: String(row[mapa.statusSku] ?? '').trim(),
        faturamentoReal: parseNumero(row[mapa.faturamentoReal]),
        faturamentoEsperado: parseNumero(row[mapa.faturamentoEsperado]),
      });
    }
    return { registros, mapa, mapaAutomatico, linhaCabecalho, headers, amostras, confiancaCabecalho };
  }

  /** "NOK" contém "OK" como substring — por isso o NOK/Zero precisa ser checado ANTES do OK. */
  function classificarGeral(status) {
    const s = normalizeKey(status);
    if (!s) return 'OUTRO';
    if (s.includes('NOK') || s.includes('ZERO')) return 'NOK';
    if (s.includes('OK') || s.includes('BATEU') || s.includes('OVER')) return 'OK';
    return 'OUTRO';
  }

  function classificarSku(status) {
    const s = normalizeKey(status);
    if (!s) return 'OUTRO';
    if (s.includes('GAP') || s.includes('FALTAM')) return 'GAP';
    if (s.includes('OK') || s.includes('BATEU') || s.includes('OVER')) return 'OK';
    return 'OUTRO';
  }

  function calcularKPIs(registros) {
    const total = registros.length;
    let giroOk = 0, vendaZero = 0, gapsSku = 0;
    registros.forEach((r) => {
      const classeGeral = classificarGeral(r.statusGeral);
      if (classeGeral === 'OK') giroOk++;
      else if (classeGeral === 'NOK') vendaZero++;
      if (classificarSku(r.statusSku) === 'GAP') gapsSku++;
    });
    return { total, giroOk, vendaZero, gapsSku };
  }

  function agruparPorSetor(registros) {
    const grupos = new Map();
    registros.forEach((r) => {
      const chave = r.setor || '—';
      if (!grupos.has(chave)) {
        grupos.set(chave, { setor: chave, responsavel: r.responsavel, total: 0, giroOk: 0, gaps: 0 });
      }
      const g = grupos.get(chave);
      g.total++;
      if (classificarGeral(r.statusGeral) === 'OK') g.giroOk++;
      if (classificarSku(r.statusSku) === 'GAP') g.gaps++;
    });
    const lista = Array.from(grupos.values()).map((g) => ({ ...g, atingimento: g.total ? g.giroOk / g.total : 0 }));
    lista.sort((a, b) => a.atingimento - b.atingimento); // pior -> melhor
    return lista;
  }

  global.ETL = {
    CONFIG,
    normalizeKey,
    parseNumero,
    encontrarAbaAlvo,
    encontrarLinhaCabecalho,
    pontuarColuna,
    mapearColunas,
    extrairColuna,
    construirRegistros,
    classificarGeral,
    classificarSku,
    calcularKPIs,
    agruparPorSetor,
  };
})(typeof self !== 'undefined' ? self : this);
