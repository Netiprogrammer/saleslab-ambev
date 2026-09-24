'use strict';

/**
 * SalesLab · Ambev — Command Center Operacional
 * ETL 100% client-side (SheetJS) + KPIs + Visão Gerencial + Auditoria de PDVs + Raio-X (Chart.js).
 * Nenhum dado sai do navegador: tudo é lido, processado e descartado em memória (RAM).
 */

// ---------------------------------------------------------------------------
// CONFIGURAÇÃO — regras de negócio centralizadas (fácil de ajustar via Diagnóstico)
// ---------------------------------------------------------------------------
const CONFIG = {
  sheetKeywordsPrimary: ['BIDEEQUIPAMENTOS', 'VISIBILIDADE', 'SKUPDV'],
  sheetKeywordsSupport: ['VISAOGERENCIAL'],
  headerScanRows: 15,
  headerKeywords: ['PDV', 'STATUS', 'SETOR'],
  minHeaderKeywordMatches: 2,
  renderChunkSize: 200,
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
  overridesStorageKey: 'saleslab_column_overrides',
  themeStorageKey: 'saleslab_theme',
};

// Estado global da aplicação (única fonte de verdade em memória).
const state = {
  registros: [],
  pdvsFiltrados: [],
  limiteRenderAuditoria: CONFIG.renderChunkSize,
  filtroSetor: null,
  registroAtivo: null,
  diagnostico: null,
  overrides: carregarOverrides(),
  chartInstance: null,
};

// ---------------------------------------------------------------------------
// UTILITÁRIOS
// ---------------------------------------------------------------------------

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

function formatMoeda(valor) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(valor || 0);
}

/** Cria um elemento DOM sem nunca usar innerHTML com dados do Excel (evita XSS). */
function el(tag, props = {}, filhos = []) {
  const node = document.createElement(tag);
  Object.entries(props).forEach(([chave, valor]) => {
    if (chave === 'class') node.className = valor;
    else if (chave === 'dataset') Object.entries(valor).forEach(([dk, dv]) => { node.dataset[dk] = dv; });
    else if (chave.startsWith('on') && typeof valor === 'function') node.addEventListener(chave.slice(2), valor);
    else node.setAttribute(chave, valor);
  });
  (Array.isArray(filhos) ? filhos : [filhos]).forEach((filho) => {
    if (filho === null || filho === undefined) return;
    node.appendChild(typeof filho === 'string' || typeof filho === 'number' ? document.createTextNode(String(filho)) : filho);
  });
  return node;
}

function getCssVar(nome) {
  return getComputedStyle(document.documentElement).getPropertyValue(nome).trim();
}

// ---------------------------------------------------------------------------
// MOTOR DE INGESTÃO — seleção de aba, detecção de cabeçalho e fuzzy matching
// ---------------------------------------------------------------------------

function encontrarAbaAlvo(workbook) {
  const nomes = workbook.SheetNames;
  const normalizados = nomes.map(normalizeKey);
  let idx = normalizados.findIndex((n) => CONFIG.sheetKeywordsPrimary.some((kw) => n.includes(kw)));
  if (idx === -1) idx = normalizados.findIndex((n) => CONFIG.sheetKeywordsSupport.some((kw) => n.includes(kw)));
  if (idx === -1) idx = 0;
  return nomes[idx];
}

/** Varre as N primeiras linhas em busca da linha que parece ser o cabeçalho real (ignora título/data). */
function encontrarLinhaCabecalho(matriz) {
  const limite = Math.min(CONFIG.headerScanRows, matriz.length);
  for (let i = 0; i < limite; i++) {
    const linha = matriz[i] || [];
    const textoLinha = linha.map(normalizeKey).join(' ');
    const matches = CONFIG.headerKeywords.filter((kw) => textoLinha.includes(kw)).length;
    if (matches >= CONFIG.minHeaderKeywordMatches) return i;
  }
  return 0;
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

function carregarOverrides() {
  try {
    return JSON.parse(localStorage.getItem(CONFIG.overridesStorageKey)) || {};
  } catch {
    return {};
  }
}

function salvarOverrides(overrides) {
  state.overrides = overrides;
  localStorage.setItem(CONFIG.overridesStorageKey, JSON.stringify(overrides));
}

function construirRegistros(worksheet) {
  const matriz = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: '' });
  const linhaCabecalho = encontrarLinhaCabecalho(matriz);
  const headers = (matriz[linhaCabecalho] || []).map((h) => String(h || '').trim()).filter(Boolean);

  const mapaAutomatico = mapearColunas(headers);
  // Overrides salvos só valem se a coluna ainda existir neste arquivo.
  const mapa = { ...mapaAutomatico };
  Object.entries(state.overrides).forEach(([campo, header]) => {
    if (headers.includes(header)) mapa[campo] = header;
  });

  const registros = [];
  for (let i = linhaCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    if (!linha || linha.every((c) => String(c ?? '').trim() === '')) continue;

    const row = {};
    headers.forEach((h, idx) => { row[h] = linha[idx]; });

    const pdv = String(row[mapa.pdv] ?? '').trim();
    const nome = String(row[mapa.nome] ?? '').trim();
    if (!pdv && !nome) continue;
    if (normalizeKey(pdv).includes('TOTAL') || normalizeKey(nome).includes('TOTAL')) continue;

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
  return { registros, mapa, mapaAutomatico, linhaCabecalho, headers };
}

// ---------------------------------------------------------------------------
// CLASSIFICAÇÃO DE STATUS
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// PRNG DETERMINÍSTICO — curva "ilustrativa" do Raio-X (estável por PDV, sem backend)
// ---------------------------------------------------------------------------

function hashString(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (Math.imul(31, h) + str.charCodeAt(i)) | 0;
  return h >>> 0;
}

function mulberry32(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function gerarCurvaSimulada(registro) {
  const rand = mulberry32(hashString(registro.pdv || registro.nome || 'seed'));
  const base = classificarGeral(registro.statusGeral) === 'OK' ? 85 : 45;
  const meses = [];
  const valores = [];
  const hoje = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(hoje.getFullYear(), hoje.getMonth() - i, 1);
    meses.push(d.toLocaleDateString('pt-BR', { month: 'short' }).replace('.', ''));
    const ruido = (rand() - 0.5) * 20;
    valores.push(Math.max(0, Math.min(100, Math.round(base + ruido))));
  }
  return { meses, valores };
}

// ---------------------------------------------------------------------------
// RENDERIZAÇÃO
// ---------------------------------------------------------------------------

function renderKpis(kpis) {
  document.getElementById('kpiTotalEquip').textContent = kpis.total.toLocaleString('pt-BR');
  document.getElementById('kpiGiroOk').textContent = kpis.giroOk.toLocaleString('pt-BR');
  document.getElementById('kpiVendaZero').textContent = kpis.vendaZero.toLocaleString('pt-BR');
  document.getElementById('kpiGapsSku').textContent = kpis.gapsSku.toLocaleString('pt-BR');
}

function renderGerencial(grupos) {
  const tbody = document.getElementById('gerencialBody');
  tbody.textContent = '';
  if (!grupos.length) {
    tbody.appendChild(el('tr', {}, el('td', { colspan: '7', class: 'vazio' }, 'Sem dados para exibir.')));
    return;
  }
  grupos.forEach((g) => {
    const pct = g.atingimento * 100;
    const corBarra = pct >= 80 ? 'ok' : pct >= 50 ? 'atencao' : 'critico';
    const tr = el('tr', { class: 'linha-clicavel', onclick: () => irParaSetor(g.setor) }, [
      el('td', {}, g.setor),
      el('td', {}, g.responsavel),
      el('td', { class: 'num' }, String(g.total)),
      el('td', { class: 'num' }, String(g.giroOk)),
      el('td', {}, el('div', { class: 'meter' }, el('div', { class: `meter-fill ${corBarra}`, style: `width:${pct.toFixed(1)}%` }))),
      el('td', { class: 'num' }, `${pct.toFixed(1)}%`),
      el('td', { class: 'num gap' }, String(g.gaps)),
    ]);
    tbody.appendChild(tr);
  });
}

function irParaSetor(setor) {
  state.filtroSetor = setor;
  document.getElementById('searchInput').value = '';
  ativarView('auditoria');
  document.getElementById('filtroSetorAtivo').textContent = setor;
  document.getElementById('filtroSetorAtivo').closest('.filtro-setor').hidden = false;
  aplicarFiltros();
}

function limparFiltroSetor() {
  state.filtroSetor = null;
  document.getElementById('filtroSetorAtivo').closest('.filtro-setor').hidden = true;
  aplicarFiltros();
}

function renderAuditoria() {
  const tbody = document.getElementById('auditoriaBody');
  tbody.textContent = '';
  const lista = state.pdvsFiltrados;
  const visiveis = lista.slice(0, state.limiteRenderAuditoria);

  if (!visiveis.length) {
    tbody.appendChild(el('tr', {}, el('td', { colspan: '6', class: 'vazio' }, 'Nenhum PDV encontrado.')));
  }

  visiveis.forEach((r) => {
    const classe = classificarGeral(r.statusGeral);
    const tr = el('tr', {}, [
      el('td', {}, r.pdv || '—'),
      el('td', {}, r.nome || '—'),
      el('td', {}, r.setor),
      el('td', {}, el('span', { class: `badge badge-${classe.toLowerCase()}` }, r.statusGeral || '—')),
      el('td', {}, r.statusSku || '—'),
      el('td', {}, el('button', { class: 'btn-auditar', onclick: () => abrirRaioX(r) }, 'Auditar')),
    ]);
    tbody.appendChild(tr);
  });

  const contador = document.getElementById('contadorAuditoria');
  contador.textContent = `Mostrando ${visiveis.length} de ${lista.length} PDVs`;
  document.getElementById('btnMostrarMais').hidden = visiveis.length >= lista.length;
}

function aplicarFiltros() {
  const termo = normalizeKey(document.getElementById('searchInput').value);
  const statusSel = document.getElementById('statusFilter').value;
  let lista = state.registros;
  if (state.filtroSetor) lista = lista.filter((r) => r.setor === state.filtroSetor);
  if (termo) lista = lista.filter((r) => normalizeKey(r.nome).includes(termo) || normalizeKey(r.pdv).includes(termo));
  if (statusSel !== 'todos') lista = lista.filter((r) => classificarGeral(r.statusGeral) === statusSel);
  state.pdvsFiltrados = lista;
  state.limiteRenderAuditoria = CONFIG.renderChunkSize;
  renderAuditoria();
}

function renderTudo() {
  renderKpis(calcularKPIs(state.registros));
  renderGerencial(agruparPorSetor(state.registros));
  aplicarFiltros();
  document.getElementById('emptyState').hidden = state.registros.length > 0;
}

// ---------------------------------------------------------------------------
// RAIO-X (painel lateral + gráfico simulado + copiar pauta)
// ---------------------------------------------------------------------------

function abrirRaioX(registro) {
  state.registroAtivo = registro;
  document.getElementById('raioXNome').textContent = registro.nome || '—';
  document.getElementById('raioXPdv').textContent = registro.pdv || '—';
  document.getElementById('raioXSetor').textContent = registro.setor;
  document.getElementById('raioXResponsavel').textContent = registro.responsavel;
  document.getElementById('raioXStatusGeral').textContent = registro.statusGeral || '—';
  document.getElementById('raioXStatusSku').textContent = registro.statusSku || '—';
  document.getElementById('raioXFatEsperado').textContent = formatMoeda(registro.faturamentoEsperado);
  document.getElementById('raioXFatReal').textContent = formatMoeda(registro.faturamentoReal);
  renderGraficoRaioX(registro);
  document.getElementById('painelRaioX').classList.add('aberto');
  document.getElementById('overlayRaioX').hidden = false;
}

function fecharRaioX() {
  document.getElementById('painelRaioX').classList.remove('aberto');
  document.getElementById('overlayRaioX').hidden = true;
}

function renderGraficoRaioX(registro) {
  const canvas = document.getElementById('chartRaioX');
  if (!window.Chart) return;
  const { meses, valores } = gerarCurvaSimulada(registro);
  if (state.chartInstance) state.chartInstance.destroy();
  state.chartInstance = new Chart(canvas, {
    type: 'line',
    data: {
      labels: meses,
      datasets: [{
        label: 'Cobertura estimada (%)',
        data: valores,
        borderColor: getCssVar('--color-accent') || '#f5b700',
        backgroundColor: 'transparent',
        tension: 0.35,
        pointRadius: 3,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: { y: { min: 0, max: 100, ticks: { callback: (v) => v + '%' } } },
      plugins: { legend: { display: false } },
    },
  });
}

function gerarTextoPauta(r) {
  const classe = classificarGeral(r.statusGeral);
  const situacao = classe === 'NOK' ? 'Venda Zero / NOK' : classe === 'OK' ? 'Regularizado' : 'Verificar status';
  return `*[Alerta SOPI]* PDV ${r.pdv || '—'} - ${r.nome || 'Sem nome'} (${r.setor}). Situação: ${situacao}. `
    + `Faturamento esperado: ${formatMoeda(r.faturamentoEsperado)} | real: ${formatMoeda(r.faturamentoReal)}. `
    + `Verificar inaderência com o representante ${r.responsavel}.`;
}

async function copiarPauta() {
  if (!state.registroAtivo) return;
  const texto = gerarTextoPauta(state.registroAtivo);
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(texto);
      mostrarToast('Pauta copiada para a área de transferência.');
      return;
    }
    throw new Error('Clipboard API indisponível');
  } catch {
    copiarFallback(texto);
  }
}

function copiarFallback(texto) {
  const area = document.createElement('textarea');
  area.value = texto;
  area.setAttribute('readonly', '');
  area.style.position = 'fixed';
  area.style.opacity = '0';
  document.body.appendChild(area);
  area.focus();
  area.select();
  let copiado = false;
  try { copiado = document.execCommand('copy'); } catch { copiado = false; }
  document.body.removeChild(area);
  if (copiado) {
    mostrarToast('Pauta copiada (modo compatibilidade).');
  } else {
    mostrarToast('Copiar automaticamente falhou — texto exibido para copiar manualmente.', 'erro');
    abrirModalTexto(texto);
  }
}

function abrirModalTexto(texto) {
  const modal = document.getElementById('modalTexto');
  const textarea = document.getElementById('modalTextoArea');
  textarea.value = texto;
  modal.hidden = false;
  textarea.focus();
  textarea.select();
}

// ---------------------------------------------------------------------------
// DIAGNÓSTICO DE MAPEAMENTO (permite corrigir a coluna escolhida por campo)
// ---------------------------------------------------------------------------

const ROTULOS_CAMPO = {
  pdv: 'Código do PDV',
  nome: 'Nome / Razão Social',
  setor: 'Setor / GV',
  responsavel: 'Responsável',
  statusGeral: 'Status Geral (Giro)',
  statusSku: 'Status SKU',
  faturamentoReal: 'Faturamento Real',
  faturamentoEsperado: 'Faturamento Esperado',
};

function abrirDiagnostico() {
  const modal = document.getElementById('diagnosticoModal');
  const corpo = document.getElementById('diagnosticoCorpo');
  corpo.textContent = '';

  if (!state.diagnostico) {
    corpo.appendChild(el('p', { class: 'vazio' }, 'Carregue uma planilha para ver o diagnóstico de mapeamento.'));
    modal.hidden = false;
    return;
  }

  const { nomeAba, linhaCabecalho, headers, mapa } = state.diagnostico;
  corpo.appendChild(el('p', {}, [el('strong', {}, 'Aba usada: '), nomeAba]));
  corpo.appendChild(el('p', {}, [el('strong', {}, 'Linha do cabeçalho: '), String(linhaCabecalho + 1)]));

  const tabela = el('table', { class: 'tabela-diagnostico' });
  const thead = el('thead', {}, el('tr', {}, [el('th', {}, 'Campo'), el('th', {}, 'Coluna detectada'), el('th', {}, 'Trocar')]));
  const tbody = el('tbody');

  Object.entries(ROTULOS_CAMPO).forEach(([campo, rotulo]) => {
    const colunaAtual = mapa[campo] || '(não encontrada)';
    const select = el('select', { dataset: { campo } }, [
      el('option', { value: '' }, '— manter automático —'),
      ...headers.map((h) => el('option', { value: h, ...(h === mapa[campo] ? { selected: 'selected' } : {}) }, h)),
    ]);
    tbody.appendChild(el('tr', {}, [
      el('td', {}, rotulo),
      el('td', {}, colunaAtual),
      el('td', {}, select),
    ]));
  });

  tabela.appendChild(thead);
  tabela.appendChild(tbody);
  corpo.appendChild(tabela);
  modal.hidden = false;
}

function salvarDiagnostico() {
  const selects = document.querySelectorAll('#diagnosticoCorpo select');
  const overrides = {};
  selects.forEach((sel) => {
    if (sel.value) overrides[sel.dataset.campo] = sel.value;
  });
  salvarOverrides(overrides);
  if (state.worksheetAtual) reprocessarComOverrides();
  fecharDiagnostico();
  mostrarToast('Mapeamento salvo. Dados reprocessados.');
}

function fecharDiagnostico() {
  document.getElementById('diagnosticoModal').hidden = true;
}

function reprocessarComOverrides() {
  const { registros, mapa, mapaAutomatico, linhaCabecalho, headers } = construirRegistros(state.worksheetAtual);
  state.registros = registros;
  state.diagnostico = { nomeAba: state.diagnostico.nomeAba, linhaCabecalho, headers, mapa, mapaAutomatico };
  renderTudo();
  atualizarBadge(registros.length);
}

// ---------------------------------------------------------------------------
// TOAST / LOADING
// ---------------------------------------------------------------------------

function mostrarToast(mensagem, tipo = 'sucesso') {
  const container = document.getElementById('toastContainer');
  const toast = el('div', { class: `toast toast-${tipo}` }, mensagem);
  container.appendChild(toast);
  setTimeout(() => toast.classList.add('saindo'), 3200);
  setTimeout(() => toast.remove(), 3600);
}

function mostrarLoading(mensagem) {
  document.getElementById('loadingMensagem').textContent = mensagem;
  document.getElementById('loadingOverlay').hidden = false;
}

function esconderLoading() {
  document.getElementById('loadingOverlay').hidden = true;
}

function atualizarBadge(qtd) {
  const badge = document.getElementById('statusBadge');
  badge.textContent = qtd > 0 ? `Pipeline Ativo (${qtd.toLocaleString('pt-BR')} registros)` : 'Aguardando Base';
  badge.classList.toggle('badge-ativo', qtd > 0);
}

function habilitarNavegacao(ativo) {
  document.getElementById('navGerencial').disabled = !ativo;
  document.getElementById('navAuditoria').disabled = !ativo;
}

// ---------------------------------------------------------------------------
// UPLOAD / ETL
// ---------------------------------------------------------------------------

async function processarArquivo(file) {
  mostrarLoading('Descodificando matrizes binárias na RAM...');
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const nomeAba = encontrarAbaAlvo(workbook);
    const worksheet = workbook.Sheets[nomeAba];
    state.worksheetAtual = worksheet;

    const { registros, mapa, mapaAutomatico, linhaCabecalho, headers } = construirRegistros(worksheet);
    state.registros = registros;
    state.diagnostico = { nomeAba, linhaCabecalho, headers, mapa, mapaAutomatico };
    state.filtroSetor = null;
    document.getElementById('filtroSetorAtivo').closest('.filtro-setor').hidden = true;

    renderTudo();
    atualizarBadge(registros.length);
    habilitarNavegacao(true);
    mostrarToast(`Base carregada: ${registros.length.toLocaleString('pt-BR')} registros na aba "${nomeAba}".`);
  } catch (err) {
    console.error(err);
    mostrarToast('Erro ao processar o arquivo: ' + err.message, 'erro');
  } finally {
    esconderLoading();
  }
}

// ---------------------------------------------------------------------------
// NAVEGAÇÃO / TEMA / INICIALIZAÇÃO
// ---------------------------------------------------------------------------

function ativarView(nomeView) {
  document.querySelectorAll('.view').forEach((v) => v.classList.toggle('ativa', v.dataset.view === nomeView));
  document.querySelectorAll('.nav-btn').forEach((b) => b.classList.toggle('ativo', b.dataset.view === nomeView));
}

function aplicarTema(tema) {
  document.documentElement.setAttribute('data-theme', tema);
  localStorage.setItem(CONFIG.themeStorageKey, tema);
  const icone = document.getElementById('iconeTema');
  if (icone) icone.className = tema === 'light' ? 'ri-moon-line' : 'ri-sun-line';
}

function initTema() {
  aplicarTema(localStorage.getItem(CONFIG.themeStorageKey) || 'dark');
  document.getElementById('btnTheme').addEventListener('click', () => {
    const atual = document.documentElement.getAttribute('data-theme') === 'light' ? 'dark' : 'light';
    aplicarTema(atual);
    if (state.registroAtivo) renderGraficoRaioX(state.registroAtivo);
  });
}

function initUpload() {
  const input = document.getElementById('fileInput');
  const btn = document.getElementById('btnUpload');
  btn.addEventListener('click', () => input.click());
  input.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) processarArquivo(file);
    input.value = ''; // permite escolher o mesmo arquivo de novo e ainda disparar 'change'
  });

  ['dragenter', 'dragover'].forEach((evt) => window.addEventListener(evt, (e) => { e.preventDefault(); document.body.classList.add('arrastando'); }));
  ['dragleave', 'drop'].forEach((evt) => window.addEventListener(evt, (e) => { e.preventDefault(); document.body.classList.remove('arrastando'); }));
  window.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) processarArquivo(file);
  });
}

function initNavegacao() {
  document.querySelectorAll('.nav-btn').forEach((btn) => {
    btn.addEventListener('click', () => ativarView(btn.dataset.view));
  });
}

function initFiltros() {
  document.getElementById('searchInput').addEventListener('input', aplicarFiltros);
  document.getElementById('statusFilter').addEventListener('change', aplicarFiltros);
  document.getElementById('btnMostrarMais').addEventListener('click', () => {
    state.limiteRenderAuditoria += CONFIG.renderChunkSize;
    renderAuditoria();
  });
  document.getElementById('btnLimparFiltroSetor').addEventListener('click', limparFiltroSetor);
}

function initRaioX() {
  document.getElementById('btnFecharRaioX').addEventListener('click', fecharRaioX);
  document.getElementById('overlayRaioX').addEventListener('click', fecharRaioX);
  document.getElementById('btnCopiarPauta').addEventListener('click', copiarPauta);
}

function initDiagnostico() {
  document.getElementById('btnDiagnostico').addEventListener('click', abrirDiagnostico);
  document.getElementById('btnFecharDiagnostico').addEventListener('click', fecharDiagnostico);
  document.getElementById('btnSalvarDiagnostico').addEventListener('click', salvarDiagnostico);
  document.getElementById('btnFecharModalTexto').addEventListener('click', () => { document.getElementById('modalTexto').hidden = true; });
}

document.addEventListener('DOMContentLoaded', () => {
  initTema();
  initUpload();
  initNavegacao();
  initFiltros();
  initRaioX();
  initDiagnostico();

  const prontas = window.bibliotecasProntas || Promise.resolve();
  prontas.then(() => {
    if (!window.XLSX) mostrarToast('SheetJS não carregou (vendor/ e CDN indisponíveis). Verifique a rede.', 'erro');
    if (!window.Chart) mostrarToast('Chart.js não carregou — o gráfico do Raio-X ficará indisponível.', 'erro');
  });
});
