'use strict';

/**
 * SalesLab · Ambev — Command Center Operacional
 * Camada de UI: orquestra o Web Worker de ingestão (worker.js/etl.js), renderiza
 * KPIs, Visão Gerencial, Auditoria de PDVs e o Raio-X (Chart.js). Nenhum dado sai
 * do navegador — tudo é lido, processado e descartado em memória (RAM).
 */

const APP_CONFIG = {
  renderChunkSize: 200,
  overridesStorageKey: 'saleslab_column_overrides',
  themeStorageKey: 'saleslab_theme',
  kpiSnapshotStorageKey: 'saleslab_kpi_snapshot',
};

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

const ETAPAS_LOADING = {
  lendo: 'Lendo o arquivo...',
  'detectando-aba': 'Detectando a aba correta...',
  'mapeando-colunas': 'Mapeando colunas e montando registros...',
  'processando-linhas': null, // mensagem dinâmica (mostra progresso de linhas)
  'calculando-indicadores': 'Calculando KPIs e agrupamentos por setor...',
};

// Estado global da aplicação (única fonte de verdade em memória).
const state = {
  registros: [],
  pdvsFiltrados: [],
  limiteRenderAuditoria: APP_CONFIG.renderChunkSize,
  filtroSetor: null,
  registroAtivo: null,
  diagnostico: null,
  overrides: carregarOverrides(),
  chartInstance: null,
  bufferAtual: null,
  worker: null,
  ordenacaoAuditoria: { campo: null, direcao: 1 },
  pilhaFocusTrap: [], // suporta overlays aninhados (ex.: Raio-X aberto + modal de cópia manual por cima)
  carregando: false,
};

// ---------------------------------------------------------------------------
// UTILITÁRIOS DE UI
// ---------------------------------------------------------------------------

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

function carregarOverrides() {
  try {
    return JSON.parse(localStorage.getItem(APP_CONFIG.overridesStorageKey)) || {};
  } catch {
    return {};
  }
}

function salvarOverrides(overrides) {
  state.overrides = overrides;
  localStorage.setItem(APP_CONFIG.overridesStorageKey, JSON.stringify(overrides));
}

/** Traduz erros técnicos (do parser/worker) em mensagens que fazem sentido pra quem não é dev. */
function mensagemAmigavel(erroTecnico) {
  const msg = String(erroTecnico || '');
  if (/Unsupported file|zip|central directory|not a valid/i.test(msg)) {
    return 'O arquivo não parece ser um Excel válido (.xlsx/.xlsm/.xls). Confira se o download não foi interrompido.';
  }
  if (/SheetJS não carregou/i.test(msg)) {
    return 'Não conseguimos carregar o motor de leitura de planilhas (rede bloqueou o vendor/ e a CDN). Veja o README sobre a pasta vendor/.';
  }
  return `Não conseguimos processar o arquivo. Detalhe técnico: ${msg}`;
}

// ---------------------------------------------------------------------------
// CLASSIFICAÇÃO / KPIs / AGRUPAMENTO (delegados ao ETL compartilhado com o worker)
// ---------------------------------------------------------------------------

const classificarGeral = (status) => ETL.classificarGeral(status);
const classificarSku = (status) => ETL.classificarSku(status);

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
// RENDERIZAÇÃO — KPIs (com delta vs. última carga) e ícones de status
// ---------------------------------------------------------------------------

const ICONES_STATUS = { OK: 'ri-checkbox-circle-fill', NOK: 'ri-close-circle-fill', OUTRO: 'ri-question-line' };

function badgeStatus(classe, texto) {
  return el('span', { class: `badge badge-${classe.toLowerCase()}` }, [
    el('i', { class: ICONES_STATUS[classe] || ICONES_STATUS.OUTRO, 'aria-hidden': 'true' }),
    ' ' + (texto || '—'),
  ]);
}

function renderKpiDelta(elementoDelta, atual, anterior) {
  if (anterior === undefined || anterior === null) { elementoDelta.textContent = ''; elementoDelta.removeAttribute('class'); return; }
  const diff = atual - anterior;
  elementoDelta.className = 'kpi-delta' + (diff > 0 ? ' subiu' : diff < 0 ? ' desceu' : ' estavel');
  elementoDelta.textContent = diff === 0 ? '— sem alteração' : `${diff > 0 ? '▲' : '▼'} ${Math.abs(diff).toLocaleString('pt-BR')} desde a última carga`;
}

function renderKpis(kpis) {
  let anterior = null;
  try { anterior = JSON.parse(localStorage.getItem(APP_CONFIG.kpiSnapshotStorageKey)); } catch { anterior = null; }

  document.getElementById('kpiTotalEquip').textContent = kpis.total.toLocaleString('pt-BR');
  document.getElementById('kpiGiroOk').textContent = kpis.giroOk.toLocaleString('pt-BR');
  document.getElementById('kpiVendaZero').textContent = kpis.vendaZero.toLocaleString('pt-BR');
  document.getElementById('kpiGapsSku').textContent = kpis.gapsSku.toLocaleString('pt-BR');

  const pares = [
    ['kpiTotalEquipDelta', kpis.total, anterior?.total],
    ['kpiGiroOkDelta', kpis.giroOk, anterior?.giroOk],
    ['kpiVendaZeroDelta', kpis.vendaZero, anterior?.vendaZero],
    ['kpiGapsSkuDelta', kpis.gapsSku, anterior?.gapsSku],
  ];
  pares.forEach(([id, atual, ant]) => renderKpiDelta(document.getElementById(id), atual, ant));

  localStorage.setItem(APP_CONFIG.kpiSnapshotStorageKey, JSON.stringify(kpis));
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
    const tr = el('tr', { class: 'linha-clicavel', tabindex: '0', role: 'button', onclick: () => irParaSetor(g.setor), onkeydown: (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); irParaSetor(g.setor); } } }, [
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

// ---------------------------------------------------------------------------
// AUDITORIA — tabela ordenável + busca + exportação CSV
// ---------------------------------------------------------------------------

function ordenarRegistros(lista) {
  const { campo, direcao } = state.ordenacaoAuditoria;
  if (!campo) return lista;
  const copia = [...lista];
  copia.sort((a, b) => {
    const va = String(a[campo] ?? '').toLowerCase();
    const vb = String(b[campo] ?? '').toLowerCase();
    return va < vb ? -direcao : va > vb ? direcao : 0;
  });
  return copia;
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
      el('td', {}, badgeStatus(classe, r.statusGeral)),
      el('td', {}, r.statusSku || '—'),
      el('td', {}, el('button', { class: 'btn-auditar', onclick: () => abrirRaioX(r) }, 'Auditar')),
    ]);
    tbody.appendChild(tr);
  });

  const contador = document.getElementById('contadorAuditoria');
  contador.textContent = `Mostrando ${visiveis.length} de ${lista.length} PDVs`;
  document.getElementById('btnMostrarMais').hidden = visiveis.length >= lista.length;
  document.getElementById('btnExportarCsv').hidden = lista.length === 0;

  document.querySelectorAll('#tabelaAuditoria th[data-campo]').forEach((th) => {
    const indicador = th.querySelector('.indicador-ordenacao');
    if (!indicador) return;
    indicador.textContent = th.dataset.campo === state.ordenacaoAuditoria.campo
      ? (state.ordenacaoAuditoria.direcao === 1 ? '▲' : '▼')
      : '';
  });
}

function aplicarFiltros() {
  const termo = ETL.normalizeKey(document.getElementById('searchInput').value);
  const statusSel = document.getElementById('statusFilter').value;
  let lista = state.registros;
  if (state.filtroSetor) lista = lista.filter((r) => r.setor === state.filtroSetor);
  if (termo) lista = lista.filter((r) => ETL.normalizeKey(r.nome).includes(termo) || ETL.normalizeKey(r.pdv).includes(termo));
  if (statusSel !== 'todos') lista = lista.filter((r) => classificarGeral(r.statusGeral) === statusSel);
  state.pdvsFiltrados = ordenarRegistros(lista);
  state.limiteRenderAuditoria = APP_CONFIG.renderChunkSize;
  renderAuditoria();
}

function escaparCsv(valor) {
  const texto = String(valor ?? '');
  return /[",\n;]/.test(texto) ? `"${texto.replace(/"/g, '""')}"` : texto;
}

function exportarCsv() {
  const colunas = ['pdv', 'nome', 'setor', 'responsavel', 'statusGeral', 'statusSku', 'faturamentoEsperado', 'faturamentoReal'];
  const cabecalho = ['PDV', 'Nome', 'Setor', 'Responsável', 'Status Geral', 'Status SKU', 'Faturamento Esperado', 'Faturamento Real'];
  const linhas = [cabecalho.join(';')];
  state.pdvsFiltrados.forEach((r) => {
    linhas.push(colunas.map((c) => escaparCsv(r[c])).join(';'));
  });
  const blob = new Blob(['﻿' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `auditoria_pdvs_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  mostrarToast(`CSV exportado com ${state.pdvsFiltrados.length.toLocaleString('pt-BR')} PDVs.`);
}

function renderTudo(kpis, grupos) {
  renderKpis(kpis);
  renderGerencial(grupos);
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
  ativarFocusTrap(document.getElementById('painelRaioX'));
}

function fecharRaioX() {
  document.getElementById('painelRaioX').classList.remove('aberto');
  document.getElementById('overlayRaioX').hidden = true;
  desativarFocusTrap();
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
  const focoOriginal = document.activeElement; // remover o textarea temporário abaixo da tela levaria o foco pro <body>
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
  if (focoOriginal && document.body.contains(focoOriginal)) focoOriginal.focus();
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
  ativarFocusTrap(modal.querySelector('.modal-card'));
  textarea.focus();
  textarea.select();
}

function fecharModalTexto() {
  document.getElementById('modalTexto').hidden = true;
  desativarFocusTrap();
}

// ---------------------------------------------------------------------------
// DIAGNÓSTICO DE MAPEAMENTO (permite corrigir a coluna escolhida por campo)
// ---------------------------------------------------------------------------

function abrirDiagnostico() {
  const modal = document.getElementById('diagnosticoModal');
  const corpo = document.getElementById('diagnosticoCorpo');
  corpo.textContent = '';

  if (!state.diagnostico) {
    corpo.appendChild(el('p', { class: 'vazio' }, 'Carregue uma planilha para ver o diagnóstico de mapeamento.'));
    modal.hidden = false;
    ativarFocusTrap(modal.querySelector('.modal-card'));
    return;
  }

  const { nomeAba, linhaCabecalho, headers, mapa, amostras, confiancaCabecalho } = state.diagnostico;
  corpo.appendChild(el('p', {}, [el('strong', {}, 'Aba usada: '), nomeAba]));
  corpo.appendChild(el('p', {}, [el('strong', {}, 'Linha do cabeçalho: '), String(linhaCabecalho + 1)]));

  if (confiancaCabecalho === 'baixa') {
    corpo.appendChild(el('p', { class: 'aviso-diagnostico' }, [
      el('i', { class: 'ri-error-warning-line', 'aria-hidden': 'true' }),
      ' Não encontramos uma linha de cabeçalho com confiança — conferindo o mapeamento abaixo, corrija manualmente o que estiver errado.',
    ]));
  }

  const tabela = el('table', { class: 'tabela-diagnostico' });
  const thead = el('thead', {}, el('tr', {}, [el('th', {}, 'Campo'), el('th', {}, 'Coluna detectada'), el('th', {}, 'Amostra'), el('th', {}, 'Trocar')]));
  const tbody = el('tbody');

  Object.entries(ROTULOS_CAMPO).forEach(([campo, rotulo]) => {
    const colunaAtual = mapa[campo] || '(não encontrada)';
    const amostraTexto = (amostras && amostras[campo] && amostras[campo].length) ? amostras[campo].join(', ') : '—';
    const select = el('select', { dataset: { campo } }, [
      el('option', { value: '' }, '— manter automático —'),
      ...headers.map((h) => el('option', { value: h, ...(h === mapa[campo] ? { selected: 'selected' } : {}) }, h)),
    ]);
    tbody.appendChild(el('tr', {}, [
      el('td', {}, rotulo),
      el('td', {}, mapa[campo] ? colunaAtual : el('span', { class: 'texto-alerta' }, colunaAtual)),
      el('td', { class: 'amostra' }, amostraTexto),
      el('td', {}, select),
    ]));
  });

  tabela.appendChild(thead);
  tabela.appendChild(tbody);
  corpo.appendChild(tabela);
  modal.hidden = false;
  ativarFocusTrap(modal.querySelector('.modal-card'));
}

function salvarDiagnostico() {
  const selects = document.querySelectorAll('#diagnosticoCorpo select');
  const overrides = {};
  selects.forEach((sel) => {
    if (sel.value) overrides[sel.dataset.campo] = sel.value;
  });
  salvarOverrides(overrides);
  if (state.bufferAtual) {
    fecharDiagnostico();
    executarIngestao(state.bufferAtual.slice(0), 'Reprocessando com o novo mapeamento...');
  } else {
    fecharDiagnostico();
  }
}

function fecharDiagnostico() {
  document.getElementById('diagnosticoModal').hidden = true;
  desativarFocusTrap();
}

// ---------------------------------------------------------------------------
// TOAST / LOADING / BADGE
// ---------------------------------------------------------------------------

function mostrarToast(mensagem, tipo = 'sucesso') {
  const container = document.getElementById('toastContainer');
  const toast = el('div', { class: `toast toast-${tipo}`, role: tipo === 'erro' ? 'alert' : 'status' }, mensagem);
  container.appendChild(toast);
  setTimeout(() => toast.classList.add('saindo'), 4200);
  setTimeout(() => toast.remove(), 4600);
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
  if (qtd > 0) {
    const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    badge.textContent = `Pipeline Ativo (${qtd.toLocaleString('pt-BR')} registros) · carregado às ${hora}`;
  } else {
    badge.textContent = 'Aguardando Base';
  }
  badge.classList.toggle('badge-ativo', qtd > 0);
}

function habilitarNavegacao(ativo) {
  document.getElementById('navGerencial').disabled = !ativo;
  document.getElementById('navAuditoria').disabled = !ativo;
  document.getElementById('btnLimparBase').hidden = !ativo;
}

// ---------------------------------------------------------------------------
// UPLOAD / ETL (via Web Worker — a thread principal nunca trava)
// ---------------------------------------------------------------------------

function obterWorker() {
  if (!state.worker) {
    state.worker = new Worker('worker.js');
  }
  return state.worker;
}

function mensagemEtapa(msg) {
  if (msg.etapa === 'processando-linhas') {
    return `Processando linha ${msg.atual.toLocaleString('pt-BR')} de ${msg.total.toLocaleString('pt-BR')}...`;
  }
  return ETAPAS_LOADING[msg.etapa] || 'Processando...';
}

function definirCarregando(carregando) {
  state.carregando = carregando;
  document.getElementById('btnDiagnostico').disabled = carregando;
  document.getElementById('btnUpload').disabled = carregando;
  document.getElementById('btnLimparBase').disabled = carregando;
}

function executarIngestao(buffer, mensagemInicial) {
  definirCarregando(true);
  mostrarLoading(mensagemInicial);
  const worker = obterWorker();

  worker.onmessage = (evento) => {
    const msg = evento.data;
    if (msg.tipo === 'progresso') {
      mostrarLoading(mensagemEtapa(msg));
      return;
    }
    if (msg.tipo === 'erro') {
      console.error(msg.mensagem);
      mostrarToast(mensagemAmigavel(msg.mensagem), 'erro');
      esconderLoading();
      definirCarregando(false);
      return;
    }
    if (msg.tipo === 'resultado') {
      state.registros = msg.registros;
      state.diagnostico = {
        nomeAba: msg.nomeAba,
        linhaCabecalho: msg.linhaCabecalho,
        headers: msg.headers,
        mapa: msg.mapa,
        mapaAutomatico: msg.mapaAutomatico,
        amostras: msg.amostras,
        confiancaCabecalho: msg.confiancaCabecalho,
      };
      state.filtroSetor = null;
      document.getElementById('filtroSetorAtivo').closest('.filtro-setor').hidden = true;

      renderTudo(msg.kpis, msg.grupos);
      atualizarBadge(msg.registros.length);
      habilitarNavegacao(true);
      esconderLoading();
      definirCarregando(false);

      if (msg.registros.length === 0) {
        mostrarToast('Nenhum registro reconhecido nesta aba. Abra o Diagnóstico para conferir a aba e o mapeamento.', 'erro');
      } else if (msg.confiancaCabecalho === 'baixa') {
        mostrarToast('Cabeçalho não identificado com confiança — confira o mapeamento no Diagnóstico.', 'erro');
      } else if (!msg.mapa.pdv && !msg.mapa.nome) {
        mostrarToast('Não localizamos as colunas de PDV/Nome automaticamente. Corrija no Diagnóstico.', 'erro');
      } else {
        mostrarToast(`Base carregada: ${msg.registros.length.toLocaleString('pt-BR')} registros na aba "${msg.nomeAba}".`);
      }
    }
  };

  worker.onerror = (erro) => {
    console.error(erro);
    mostrarToast(mensagemAmigavel(erro.message), 'erro');
    esconderLoading();
    definirCarregando(false);
  };

  const copiaParaWorker = buffer.slice(0);
  worker.postMessage({ buffer: copiaParaWorker, overrides: state.overrides }, [copiaParaWorker]);
}

const EXTENSOES_ACEITAS = /\.(xlsx|xlsm|xls)$/i;

async function processarArquivo(file) {
  if (!EXTENSOES_ACEITAS.test(file.name)) {
    mostrarToast(`Formato "${file.name.split('.').pop()}" não suportado. Use .xlsx, .xlsm ou .xls.`, 'erro');
    return;
  }
  mostrarLoading('Lendo o arquivo...');
  try {
    const buffer = await file.arrayBuffer();
    state.bufferAtual = buffer;
    executarIngestao(buffer.slice(0), 'Descodificando matrizes binárias na RAM...');
  } catch (err) {
    console.error(err);
    mostrarToast(mensagemAmigavel(err.message), 'erro');
    esconderLoading();
  }
}

function limparBase() {
  if (!confirm('Limpar a base carregada? Isso não afeta o mapeamento salvo nem o tema.')) return;
  state.registros = [];
  state.pdvsFiltrados = [];
  state.diagnostico = null;
  state.bufferAtual = null;
  state.filtroSetor = null;
  state.registroAtivo = null;
  document.getElementById('filtroSetorAtivo').closest('.filtro-setor').hidden = true;
  renderTudo({ total: 0, giroOk: 0, vendaZero: 0, gapsSku: 0 }, []);
  atualizarBadge(0);
  habilitarNavegacao(false);
  ativarView('gerencial');
  mostrarToast('Base removida da memória.');
}

// ---------------------------------------------------------------------------
// FOCUS TRAP / FECHAMENTO PADRÃO (Esc + clique fora) PARA TODOS OS OVERLAYS
// ---------------------------------------------------------------------------

function elementosFocaveis(container) {
  return Array.from(container.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'))
    .filter((elemento) => !elemento.disabled && elemento.offsetParent !== null);
}

/**
 * Empilha o trap (não substitui um único global) para suportar overlays aninhados —
 * ex.: abrir o Raio-X, clicar em "Copiar Pauta" e cair no modal de cópia manual por cima.
 * Fechar o de cima restaura o foco e o trap do de baixo, em vez de perder a referência.
 */
function ativarFocusTrap(container) {
  const focoAnterior = document.activeElement;
  // Foca o próprio diálogo (não o primeiro campo aleatório lá dentro) — padrão de acessibilidade
  // pra modal: o leitor de tela anuncia o título do diálogo em vez de cair direto num <select> da tabela.
  if (!container.hasAttribute('tabindex')) container.setAttribute('tabindex', '-1');
  container.focus();

  const aoTeclar = (e) => {
    if (e.key !== 'Tab') return;
    const lista = elementosFocaveis(container);
    if (!lista.length) return;
    const primeiro = lista[0];
    const ultimo = lista[lista.length - 1];
    if (e.shiftKey && document.activeElement === primeiro) { e.preventDefault(); ultimo.focus(); }
    else if (!e.shiftKey && document.activeElement === ultimo) { e.preventDefault(); primeiro.focus(); }
  };
  container.addEventListener('keydown', aoTeclar);
  state.pilhaFocusTrap.push({
    focoAnterior,
    remover: () => container.removeEventListener('keydown', aoTeclar),
  });
}

function desativarFocusTrap() {
  const topo = state.pilhaFocusTrap.pop();
  if (!topo) return;
  topo.remover();
  if (topo.focoAnterior && document.body.contains(topo.focoAnterior)) topo.focoAnterior.focus();
}

/** Fecha o overlay "mais de cima" aberto no momento (Esc funciona igual em qualquer modal/painel). */
function fecharOverlayAtivo() {
  if (!document.getElementById('modalTexto').hidden) { fecharModalTexto(); return true; }
  if (!document.getElementById('diagnosticoModal').hidden) { fecharDiagnostico(); return true; }
  if (document.getElementById('painelRaioX').classList.contains('aberto')) { fecharRaioX(); return true; }
  return false;
}

function tornarFechavelPorClique(overlayEl, aoFechar) {
  overlayEl.addEventListener('click', (e) => { if (e.target === overlayEl) aoFechar(); });
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
  localStorage.setItem(APP_CONFIG.themeStorageKey, tema);
  const icone = document.getElementById('iconeTema');
  if (icone) icone.className = tema === 'light' ? 'ri-moon-line' : 'ri-sun-line';
}

function initTema() {
  aplicarTema(localStorage.getItem(APP_CONFIG.themeStorageKey) || 'dark');
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

  document.getElementById('btnLimparBase').addEventListener('click', limparBase);
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
    state.limiteRenderAuditoria += APP_CONFIG.renderChunkSize;
    renderAuditoria();
  });
  document.getElementById('btnLimparFiltroSetor').addEventListener('click', limparFiltroSetor);
  document.getElementById('btnExportarCsv').addEventListener('click', exportarCsv);

  document.querySelectorAll('#tabelaAuditoria th[data-campo]').forEach((th) => {
    th.addEventListener('click', () => {
      const campo = th.dataset.campo;
      if (state.ordenacaoAuditoria.campo === campo) {
        state.ordenacaoAuditoria.direcao *= -1;
      } else {
        state.ordenacaoAuditoria = { campo, direcao: 1 };
      }
      aplicarFiltros();
    });
  });

  // Atalho "/" foca a busca (só quando não estamos digitando em outro campo).
  document.addEventListener('keydown', (e) => {
    if (e.key !== '/' || (e.target instanceof Element && e.target.matches('input, textarea, select'))) return;
    const auditoriaAtiva = document.querySelector('.view[data-view="auditoria"]').classList.contains('ativa');
    if (!auditoriaAtiva) return;
    e.preventDefault();
    document.getElementById('searchInput').focus();
  });
}

function initRaioX() {
  document.getElementById('btnFecharRaioX').addEventListener('click', fecharRaioX);
  tornarFechavelPorClique(document.getElementById('overlayRaioX'), fecharRaioX);
  document.getElementById('btnCopiarPauta').addEventListener('click', copiarPauta);
}

function initDiagnostico() {
  document.getElementById('btnDiagnostico').addEventListener('click', abrirDiagnostico);
  document.getElementById('btnFecharDiagnostico').addEventListener('click', fecharDiagnostico);
  document.getElementById('btnSalvarDiagnostico').addEventListener('click', salvarDiagnostico);
  tornarFechavelPorClique(document.getElementById('diagnosticoModal'), fecharDiagnostico);

  document.getElementById('btnFecharModalTexto').addEventListener('click', fecharModalTexto);
  tornarFechavelPorClique(document.getElementById('modalTexto'), fecharModalTexto);
}

function initTeclasGlobais() {
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') fecharOverlayAtivo();
  });
}

function initPWA() {
  if (!('serviceWorker' in navigator)) return;
  if (window.DESATIVAR_SW_EM_DEV) return; // ligado só durante o desenvolvimento local, pra não cachear versões antigas
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch((err) => console.warn('Service Worker não registrado:', err));
  });
}

document.addEventListener('DOMContentLoaded', () => {
  initTema();
  initUpload();
  initNavegacao();
  initFiltros();
  initRaioX();
  initDiagnostico();
  initTeclasGlobais();
  initPWA();

  const prontas = window.bibliotecasProntas || Promise.resolve();
  prontas.then(() => {
    if (!window.XLSX) mostrarToast('SheetJS não carregou (vendor/ e CDN indisponíveis). Verifique a rede.', 'erro');
    if (!window.Chart) mostrarToast('Chart.js não carregou — o gráfico do Raio-X ficará indisponível.', 'erro');
  });
});
