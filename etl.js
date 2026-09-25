'use strict';

/**
 * ETL puro (sem DOM) — compartilhado entre a thread principal e o Web Worker de ingestão.
 * Carregado via <script src="etl.js"> na página e via importScripts('etl.js') no worker;
 * por isso expõe tudo em `self.ETL` (em vez de `const`/`let` de topo, que não vira global).
 *
 * Lê VÁRIAS abas relevantes do arquivo (não só a "melhor"/maior) e mescla os registros pela
 * chave do PDV: cada campo é preenchido pela primeira aba (em ordem de prioridade) que
 * realmente tiver aquela coluna — uma aba mais pobre processada depois só completa o que
 * falta, nunca sobrescreve o que uma aba melhor já preencheu.
 */
(function (global) {
  // Campos de identidade — compartilhados entre a leitura "geral" e a de "chopeira",
  // pra que o Diagnóstico consiga mostrar nome/setor/responsável também pra quem só
  // aparece na aba de chopeira.
  const CAMPOS_IDENTIDADE = {
    pdv: { keywords: ['PDV', 'COD', 'CLIENTE'], tipo: 'texto', padrao: '' },
    nome: { keywords: ['RAZAO', 'NOME', 'FANTASIA'], tipo: 'texto', padrao: '—' },
    setor: { keywords: ['GV', 'SETOR', 'CODSETOR'], tipo: 'texto', padrao: '—' },
    responsavel: { keywords: ['SUPERCOM', 'REPRESENTANTE', 'COMERCIAL', 'RN', 'DONO'], tipo: 'texto', padrao: '—' },
  };

  const CAMPOS_GERAL = {
    ...CAMPOS_IDENTIDADE,
    statusGeral: { keywords: ['STATUSPDV', 'STATUSDO', 'GIRO'], tipo: 'texto', padrao: '' },
    statusSku: { keywords: ['STATUSSKU', 'SKU'], tipo: 'texto', padrao: '' },
    // "Faturamento PDV" é o nome real usado na planilha da Ambev pro valor realizado —
    // mais específico que "REAL" (que não aparece de fato em nenhuma coluna real observada).
    faturamentoReal: { keywords: ['FATURAMENTOPDV', 'FATURAMENTOREAL', 'REAL'], tipo: 'numero', padrao: 0 },
    faturamentoEsperado: { keywords: ['FATURAMENTOESPERADO', 'ESPERADO'], tipo: 'numero', padrao: 0 },
  };

  const CAMPOS_CHOPEIRA = {
    ...CAMPOS_IDENTIDADE,
    statusChopeira: { keywords: ['STATUS'], tipo: 'texto', padrao: '' },
    faturamentoChopeira: { keywords: ['FATURAMENTOPDV', 'FATURAMENTO'], tipo: 'numero', padrao: 0 },
    metaChopeira: { keywords: ['META'], tipo: 'numero', padrao: 0 },
    gapChopeira: { keywords: ['GAP'], tipo: 'numero', padrao: 0 },
  };

  const CAMPOS_TODOS = { ...CAMPOS_GERAL, ...CAMPOS_CHOPEIRA };

  const CONFIG = {
    // Ordem = prioridade: a primeira aba encontrada que casar com uma keyword de índice
    // menor é processada primeiro, e "ganha" a mesclagem por PDV.
    keywordsGeral: ['BIDEEQUIPAMENTOS', 'VISIBILIDADE', 'SKUPDV'],
    keywordsChopeira: ['CHOPEIRA'],
    keywordsSuporte: ['VISAOGERENCIAL'], // só usada se NENHUMA aba "geral" for encontrada
    headerScanRows: 15,
    headerKeywords: ['PDV', 'STATUS', 'SETOR'],
    minHeaderKeywordMatches: 2,
    amostrasPorCampo: 3,
    linhasPorLoteProgresso: 5000,
  };

  /** Remove acentos, espaços, quebras de linha e pontuação; deixa tudo maiúsculo. */
  function normalizeKey(str) {
    return String(str ?? '')
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-zA-Z0-9]/g, '')
      .toUpperCase();
  }

  /**
   * Fallback pra quando o valor numérico não veio direto da célula (raro — normalmente
   * `construirRegistros` já usa o número bruto da célula, sem passar por aqui). Detecta se
   * "," ou "." é o separador decimal pela posição (o último dos dois, com só 1-2 dígitos
   * depois) em vez de assumir formato brasileiro sempre — essa planilha da Ambev usa "R$
   * 2,000.00" (padrão US: vírgula de milhar, ponto decimal), não o formato BR.
   */
  function parseNumero(valor) {
    if (valor === undefined || valor === null || valor === '') return 0;
    if (typeof valor === 'number') return valor;
    let texto = String(valor).replace(/[^\d,.-]/g, '');
    if (!texto) return 0;
    const ultimaVirgula = texto.lastIndexOf(',');
    const ultimoPonto = texto.lastIndexOf('.');
    if (ultimaVirgula > ultimoPonto && texto.length - ultimaVirgula - 1 <= 2) {
      texto = texto.replace(/\./g, '').replace(',', '.'); // decimal com vírgula (BR): 1.234,56
    } else {
      texto = texto.replace(/,/g, ''); // vírgula é separador de milhar (US): 4,613 ou 1,234.56
    }
    const numero = parseFloat(texto);
    return Number.isFinite(numero) ? numero : 0;
  }

  /**
   * Acha todas as abas relevantes (não só a primeira/maior) e a ORDEM de prioridade em que
   * devem ser mescladas. Se nenhuma aba "geral" bater com as keywords principais, cai pra
   * aba de apoio (Visão Gerencial) ou a primeira aba do arquivo, pra nunca ficar sem nada.
   */
  function encontrarAbasCandidatas(workbook) {
    const nomes = workbook.SheetNames;
    const candidatasGeral = [];
    const candidatasChopeira = [];

    nomes.forEach((nome) => {
      const norm = normalizeKey(nome);
      const idxGeral = CONFIG.keywordsGeral.findIndex((kw) => norm.includes(kw));
      if (idxGeral !== -1) { candidatasGeral.push({ sheetName: nome, prioridade: idxGeral }); return; }
      const idxChopeira = CONFIG.keywordsChopeira.findIndex((kw) => norm.includes(kw));
      if (idxChopeira !== -1) candidatasChopeira.push({ sheetName: nome, prioridade: idxChopeira });
    });

    candidatasGeral.sort((a, b) => a.prioridade - b.prioridade);
    candidatasChopeira.sort((a, b) => a.prioridade - b.prioridade);

    if (!candidatasGeral.length) {
      const idxSuporte = nomes.findIndex((n) => CONFIG.keywordsSuporte.some((kw) => normalizeKey(n).includes(kw)));
      candidatasGeral.push({ sheetName: nomes[idxSuporte !== -1 ? idxSuporte : 0], prioridade: 0 });
    }

    return [
      ...candidatasGeral.map((c) => ({ ...c, papel: 'geral' })),
      ...candidatasChopeira.map((c) => ({ ...c, papel: 'chopeira' })),
    ];
  }

  /** Varre as N primeiras linhas em busca da linha que parece ser o cabeçalho real (ignora título/data). */
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
  function mapearColunas(headers, camposDef) {
    const candidatos = [];
    Object.entries(camposDef).forEach(([campo, def]) => {
      headers.forEach((header) => {
        const chaveNormalizada = normalizeKey(header);
        if (!chaveNormalizada) return;
        const pontuacao = pontuarColuna(chaveNormalizada, def.keywords);
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
   * @param {object} [camposDef] quais campos procurar e com quais keywords (padrão: CAMPOS_GERAL)
   */
  function construirRegistros(worksheet, overridesSalvos = {}, onProgresso, camposDef = CAMPOS_GERAL) {
    // Duas leituras: texto formatado (pra cabeçalho/status/nome — mais legível) e valores
    // brutos da célula (pra números — evita que o formato de exibição da célula, tipo "sem
    // casas decimais" ou "vírgula de milhar", arredonde ou distorça o valor real).
    const matrizTexto = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
    const matrizBruta = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' });
    const { indice: linhaCabecalho, confianca: confiancaCabecalho } = encontrarLinhaCabecalho(matrizTexto);
    const headers = (matrizTexto[linhaCabecalho] || []).map((h) => String(h || '').trim()).filter(Boolean);

    const mapaAutomatico = mapearColunas(headers, camposDef);
    // Overrides salvos só valem se a coluna ainda existir neste arquivo.
    const mapa = { ...mapaAutomatico };
    Object.entries(overridesSalvos).forEach(([campo, header]) => {
      if (headers.includes(header)) mapa[campo] = header;
    });

    const registros = [];
    const amostras = {};
    const totalLinhas = matrizTexto.length - (linhaCabecalho + 1);

    for (let i = linhaCabecalho + 1; i < matrizTexto.length; i++) {
      const linha = matrizTexto[i];
      const linhaBruta = matrizBruta[i];
      if (onProgresso && (i - linhaCabecalho) % CONFIG.linhasPorLoteProgresso === 0) {
        onProgresso(i - linhaCabecalho, totalLinhas);
      }
      if (!linha || linha.every((c) => String(c ?? '').trim() === '')) continue;

      const row = {};
      const rowBruta = {};
      headers.forEach((h, idx) => { row[h] = linha[idx]; rowBruta[h] = linhaBruta ? linhaBruta[idx] : undefined; });

      const pdv = String(row[mapa.pdv] ?? '').trim();
      const nome = mapa.nome ? String(row[mapa.nome] ?? '').trim() : '';
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

      const registro = {};
      Object.entries(camposDef).forEach(([campo, def]) => {
        if (def.tipo === 'numero') {
          const bruto = rowBruta[mapa[campo]];
          registro[campo] = typeof bruto === 'number' ? bruto : parseNumero(row[mapa[campo]]);
        } else {
          registro[campo] = String(row[mapa[campo]] ?? '').trim();
        }
      });
      registros.push(registro);
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

  function criarRegistroVazio() {
    const registro = { fontes: [], _preenchidos: new Set() };
    Object.entries(CAMPOS_TODOS).forEach(([campo, def]) => { registro[campo] = def.padrao; });
    return registro;
  }

  /**
   * Mescla os registros de UMA aba no mapa consolidado (chave = PDV normalizado).
   * Só preenche campos que essa aba realmente tem (via `mapa`, não pelo valor lido) e que
   * NENHUMA aba anterior já preencheu — assim uma aba processada depois nunca sobrescreve
   * um valor válido (incluindo zero legítimo, como "Faturamento Real = 0" numa Venda Zero)
   * só porque o valor "parece vazio".
   */
  function mesclarRegistros(consolidado, resultado, sheetName) {
    const camposDisponiveis = Object.keys(resultado.mapa);
    resultado.registros.forEach((r) => {
      const chave = normalizeKey(r.pdv);
      if (!chave) return;

      if (!consolidado.has(chave)) {
        const registro = criarRegistroVazio();
        camposDisponiveis.forEach((campo) => {
          registro[campo] = r[campo];
          registro._preenchidos.add(campo);
        });
        registro.fontes.push(sheetName);
        consolidado.set(chave, registro);
        return;
      }

      const existente = consolidado.get(chave);
      camposDisponiveis.forEach((campo) => {
        if (!existente._preenchidos.has(campo)) {
          existente[campo] = r[campo];
          existente._preenchidos.add(campo);
        }
      });
      if (!existente.fontes.includes(sheetName)) existente.fontes.push(sheetName);
    });
  }

  /**
   * Orquestra a leitura de TODAS as abas relevantes do workbook (não só uma) e devolve os
   * PDVs já mesclados, mais o diagnóstico de cada aba processada (pra tela de Diagnóstico).
   * @param {object} overridesSalvos formato { [nomeDaAba]: { [campo]: header } }
   */
  function processarWorkbook(workbook, overridesSalvos = {}, onProgresso) {
    const candidatas = encontrarAbasCandidatas(workbook);
    const consolidado = new Map();
    const abasProcessadas = [];

    candidatas.forEach(({ sheetName, papel }, indiceAba) => {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) return;
      const camposDef = papel === 'chopeira' ? CAMPOS_CHOPEIRA : CAMPOS_GERAL;
      const overridesAba = overridesSalvos[sheetName] || {};
      const progressoAba = onProgresso
        ? (atual, total) => onProgresso({ sheetName, papel, indiceAba, totalAbas: candidatas.length, atual, total })
        : undefined;

      const resultado = construirRegistros(worksheet, overridesAba, progressoAba, camposDef);
      mesclarRegistros(consolidado, resultado, sheetName);
      abasProcessadas.push({
        sheetName,
        papel,
        linhaCabecalho: resultado.linhaCabecalho,
        headers: resultado.headers,
        mapa: resultado.mapa,
        mapaAutomatico: resultado.mapaAutomatico,
        amostras: resultado.amostras,
        confiancaCabecalho: resultado.confiancaCabecalho,
        totalRegistros: resultado.registros.length,
      });
    });

    const registros = Array.from(consolidado.values()).map((r) => {
      const limpo = { fontes: r.fontes };
      Object.keys(CAMPOS_TODOS).forEach((campo) => { limpo[campo] = r[campo]; });
      return limpo;
    });

    const pdvsComMultiplasFontes = registros.filter((r) => r.fontes.length > 1).length;

    return { registros, abasProcessadas, pdvsComMultiplasFontes };
  }

  global.ETL = {
    CONFIG,
    CAMPOS_GERAL,
    CAMPOS_CHOPEIRA,
    CAMPOS_TODOS,
    normalizeKey,
    parseNumero,
    encontrarAbasCandidatas,
    encontrarLinhaCabecalho,
    pontuarColuna,
    mapearColunas,
    extrairColuna,
    construirRegistros,
    mesclarRegistros,
    processarWorkbook,
    classificarGeral,
    classificarSku,
    calcularKPIs,
    agruparPorSetor,
  };
})(typeof self !== 'undefined' ? self : this);
