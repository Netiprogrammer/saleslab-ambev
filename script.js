'use strict';

/* =====================================================================
   SalesLab Ops: Central Executiva Ambev
   App 100% client-side: o Excel é lido na memória do navegador e nada
   é enviado para nenhum servidor.
   ===================================================================== */

/* ---------- Configuração (ajuste aqui sem mexer no resto) ---------- */
const CONFIG = {
    META_ATINGIMENTO: 70,          // % mínimo de Giro OK por setor para ficar verde
    PAGINA: 200,                   // linhas exibidas por vez na Auditoria de PDVs
    LINHAS_BUSCA_CABECALHO: 40,    // quantas linhas do topo varrer atrás do cabeçalho
    MAX_STATUS_DIAG: 14,           // valores distintos de status listados no Diagnóstico
    CHAVE_MAPA: 'saleslab.mapa.v1' // onde o mapeamento manual de colunas é salvo
};

/* Bibliotecas: tenta primeiro a cópia local (pasta vendor/), depois a CDN.
   Assim o app continua funcionando em rede corporativa que bloqueia CDNs. */
const LIBS = {
    XLSX:  ['vendor/xlsx.full.min.js', 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js'],
    Chart: ['vendor/chart.umd.js', 'https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.js']
};

/* Termos que identificam a linha de cabeçalho (já sem acento e em maiúsculas) */
const TERMOS_CABECALHO = ['PDV', 'STATUS', 'SETOR', 'RAZAO', 'FANTASIA', 'GV', 'CLIENTE', 'SKU', 'GAP', 'GIRO', 'FATURAMENTO', 'COD', 'NOME'];

/* Campos do app e como localizá-los. Os padrões são testados contra o nome da
   coluna normalizado (sem acento, maiúsculas, quebras de linha viram espaço) e a
   ORDEM importa: o primeiro padrão é o mais específico. */
const CAMPOS = [
    { id: 'pdv',     rotulo: 'Código do PDV',
      excluir: /STATUS|GAP|FAT|NOME|RAZAO|SETOR|\bGV\b|SKU|EQUIP/,
      padroes: [/^(COD|CODIGO) (DO )?(PDV|CLIENTE)$/, /^(PDV|CLIENTE|COD|CODIGO)$/, /^(COD|CODIGO)\b/, /\bPDV\b/] },
    { id: 'nome',    rotulo: 'Nome fantasia',
      excluir: /REPRESENTANTE|SUPERVIS|VENDEDOR|\bGV\b|\bRN\b|STATUS|SETOR/,
      padroes: [/NOME FANTASIA/, /FANTASIA/, /RAZAO SOCIAL/, /RAZAO/, /^NOME/, /NOME/] },
    { id: 'setor',   rotulo: 'Setor',
      excluir: /STATUS|GAP/,
      padroes: [/^SETOR$/, /^(COD )?SETOR/, /SETOR/, /^GV$/] },
    { id: 'dono',    rotulo: 'Representante / dono',
      excluir: /STATUS|GAP/,
      padroes: [/REPRESENTANTE/, /^RN$/, /^DONO/, /SUPERVIS|SUPERCOM/, /^GV$/, /COMERCIAL/] },
    { id: 'status',  rotulo: 'Status do PDV (giro)',
      excluir: /SKU|GAP/,
      padroes: [/^STATUS PDV$/, /^STATUS( GERAL| GIRO| DO PDV)?$/, /STATUS.*(PDV|GIRO)/, /GIRO/, /SITUACAO/, /STATUS/] },
    { id: 'gap',     rotulo: 'GAP de SKU',
      excluir: null,
      padroes: [/^GAP SKU PDV$/, /GAP.*SKU|SKU.*GAP/, /^GAP/, /GAP/, /FALTAM?/] },
    { id: 'fatEsp',  rotulo: 'Faturamento esperado',
      excluir: null,
      padroes: [/FAT\w* ESPERAD/, /ESPERAD/, /META (FAT|R )/] },
    { id: 'fatReal', rotulo: 'Faturamento real',
      excluir: null,
      padroes: [/FAT\w* (REAL|PDV|ATUAL)/, /REALIZAD/, /^REAL$/, /FATURAMENTO/] }
];

/* Classificação do texto do status. A ordem importa: "NOK" é testado antes de
   "OK" (e \b garante que "NOK" nunca case com "OK"). Edite os termos se o BI
   usar outras palavras: o Diagnóstico mostra como cada valor foi classificado. */
const CLASSES = [
    { cat: 'vz',  rx: /\bVENDA ZERO\b|\bVZ\b|\bSEM VENDA\b/ },
    { cat: 'nok', rx: /\bNOK\b|\bNAO OK\b|\bNAO ADERENTE\b|\bINADERENTE\b|\bNAO BATEU\b/ },
    { cat: 'gap', rx: /\bGAP\b|\bFALTAM?\b/ },
    { cat: 'ok',  rx: /\bOK\b|\bOVER\b|\bBATEU\b|\bADERENTE\b/ }
];
const ROTULO_CAT = { ok: 'Giro OK', gap: 'GAP', vz: 'Venda zero', nok: 'NOK', outro: 'Não classificado', vazio: 'Sem status' };
const CLASSE_CAT = { ok: 'ok', gap: 'danger', vz: 'warning', nok: 'danger', outro: 'neutral', vazio: 'neutral' };

/* ---------- Utilitários ---------- */
const $ = (id) => document.getElementById(id);

const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;', '`': '&#96;' };
/* Todo texto vindo do Excel passa por aqui antes de entrar em innerHTML */
const escapeHTML = (v) => (v === null || v === undefined) ? '' : String(v).replace(/[&<>'"`]/g, (c) => ESC[c]);

const txt = (v) => (v === null || v === undefined) ? '' : String(v).trim();
const semAcento = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const simples = (s) => semAcento(s).toLowerCase();
/* "Status \nPDV" -> "STATUS PDV" | "GAP SKU/PDV" -> "GAP SKU PDV" | "Cód. PDV" -> "COD PDV" */
const normalizar = (v) => semAcento(v).toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();

const fmtInt = (n) => Number(n).toLocaleString('pt-BR');
const BRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });

/* Cede o controle ao navegador para pintar o overlay antes de um trabalho pesado.
   O timeout garante que não trava se a aba estiver em segundo plano. */
const frame = () => new Promise((resolve) => {
    let feito = false;
    const fim = () => { if (!feito) { feito = true; resolve(); } };
    requestAnimationFrame(() => setTimeout(fim, 0));
    setTimeout(fim, 120);
});

const store = {
    get(chave, padrao = null) { try { const v = localStorage.getItem(chave); return v === null ? padrao : v; } catch (_) { return padrao; } },
    set(chave, valor) { try { localStorage.setItem(chave, valor); } catch (_) { /* storage bloqueado: segue sem persistir */ } }
};

function paraNumero(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v !== 'string') return null;
    let s = v.trim().replace(/[R$\s%]/g, '');
    if (!s || !/[0-9]/.test(s)) return null;
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');   // 1.234,56
    else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');                                   // 12,5
    else if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : null;
}

function fmtValor(v) {
    const n = paraNumero(v);
    if (n !== null) return BRL.format(n);
    return txt(v) || '—';
}

const setorLabel = (s) => /^\d+$/.test(s) ? `Setor ${s}` : s;

/* ---------- Estado e referências do DOM ---------- */
const estado = {
    arquivo: '', buffer: null, abas: [], aba: '',
    cabecalhos: [], linhas: [], linhaCabecalho: 0, cabecalhoDetectado: false,
    mapa: {}, registros: [], setores: [],
    modo: 'gerencial', filtrados: [], visiveis: CONFIG.PAGINA,
    filtroSetor: null, selecionadoId: null,
    chart: null, regGrafico: null, ocupado: false
};

const dom = {
    fileInput: $('excelFileInput'), btnUpload: $('btn-upload'), btnDiag: $('btn-diag'),
    themeToggle: $('theme-toggle'), viewTitle: $('view-title'), fileStatus: $('file-status'),
    kpiTotal: $('kpi-total'), kpiOk: $('kpi-ok'), kpiVz: $('kpi-vz'), kpiGap: $('kpi-gap'),
    kpiTotalSub: $('kpi-total-sub'), kpiOkSub: $('kpi-ok-sub'), kpiVzSub: $('kpi-vz-sub'), kpiGapSub: $('kpi-gap-sub'),
    toolbar: $('toolbar-search'), busca: $('tableSearch'), filtroStatus: $('filtroStatus'), chipSetor: $('chipSetor'),
    split: $('splitView'), thead: $('tableHead'), tbody: $('tableBody'),
    footer: $('tableFooter'), count: $('tableCount'), btnMais: $('btnMais'),
    side: $('sideDetail'), detId: $('det-id'), detContent: $('det-content'),
    chartWrap: $('chartContainer'), chartNote: $('chartNote'), canvas: $('pdvChart'),
    overlay: $('loadingOverlay'), loadTitle: $('loadingTitle'), loadText: $('loadingText'), loadBar: $('loadingBar'),
    dropHint: $('dropHint'), modal: $('diagModal'), diagSheet: $('diagSheet'), diagInfo: $('diagInfo'),
    diagMap: $('diagMap'), diagStatus: $('diagStatus'), diagClose: $('diagClose'), diagOk: $('diagOk'), diagReset: $('diagReset'),
    toastStack: $('toastStack')
};

/* ---------- Feedback: toast e loading ---------- */
function toast(msg, tipo = 'info', ms = 4500) {
    const el = document.createElement('div');
    el.className = `toast ${tipo}`;
    el.textContent = msg;
    el.addEventListener('click', () => el.remove());
    dom.toastStack.appendChild(el);
    setTimeout(() => el.remove(), ms);
}

function mostrarLoading(titulo, texto, pct) {
    dom.loadTitle.textContent = titulo;
    dom.loadText.textContent = texto || '';
    if (pct === null || pct === undefined) {
        dom.loadBar.classList.add('indet');
        dom.loadBar.style.width = '';
    } else {
        dom.loadBar.classList.remove('indet');
        dom.loadBar.style.width = `${Math.round(pct * 100)}%`;
    }
    dom.overlay.classList.remove('hidden');
}
const esconderLoading = () => dom.overlay.classList.add('hidden');

/* ---------- Bibliotecas (vendor/ primeiro, CDN depois) ---------- */
const promessasLib = {};
function carregarScript(src) {
    return new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolve;
        s.onerror = () => reject(new Error(`Falha ao carregar ${src}`));
        document.head.appendChild(s);
    });
}
function garantirLib(nome) {
    if (window[nome]) return Promise.resolve(true);
    if (!promessasLib[nome]) {
        promessasLib[nome] = (async () => {
            for (const src of LIBS[nome]) {
                try { await carregarScript(src); if (window[nome]) return true; } catch (_) { /* tenta a próxima fonte */ }
            }
            return false;
        })().then((ok) => { if (!ok) delete promessasLib[nome]; return ok; });
    }
    return promessasLib[nome];
}

/* ---------- Tema ---------- */
function aplicarTema(tema) {
    document.documentElement.setAttribute('data-theme', tema);
    store.set('theme', tema);
    const icone = dom.themeToggle.querySelector('i');
    if (icone) icone.className = tema === 'light' ? 'ri-moon-line' : 'ri-sun-line';
    if (estado.regGrafico) desenharGrafico(estado.regGrafico);   // recolore o gráfico
}

/* =====================================================================
   MOTOR DE INGESTÃO (ETL em memória)
   ===================================================================== */

function lerArquivo(file, onProgress) {
    return new Promise((resolve, reject) => {
        const leitor = new FileReader();
        leitor.onprogress = (e) => { if (e.lengthComputable) onProgress(e.loaded / e.total); };
        leitor.onload = () => resolve(leitor.result);
        leitor.onerror = () => reject(leitor.error || new Error('Não foi possível ler o arquivo.'));
        leitor.readAsArrayBuffer(file);
    });
}

/* Aba preferida: a de visibilidade / BI de equipamentos. Sem ela, a primeira. */
function escolherAba(nomes) {
    const preferidas = [/VISIBILIDADE/, /BI.*EQUIP/, /SKU.*PDV|PDV.*SKU/];
    for (const rx of preferidas) {
        const achada = nomes.find((n) => rx.test(normalizar(n)));
        if (achada) return achada;
    }
    return nomes[0];
}

/* Smart Header Parser: varre o topo da aba atrás da linha que parece cabeçalho
   (várias células de texto e pelo menos 2 termos-chave). Linhas de título ou de
   data ("DATA: 23/09/2026") têm poucas células e ficam de fora. */
function detectarCabecalho(matriz) {
    const limite = Math.min(CONFIG.LINHAS_BUSCA_CABECALHO, matriz.length);
    let melhor = { idx: -1, hits: 0 };
    for (let i = 0; i < limite; i++) {
        const linha = matriz[i] || [];
        const textos = linha.filter((c) => typeof c === 'string' && c.trim() !== '');
        if (textos.length < 3) continue;
        const tokens = new Set();
        textos.forEach((t) => normalizar(t).split(' ').forEach((p) => tokens.add(p)));
        const hits = TERMOS_CABECALHO.filter((t) => tokens.has(t)).length;
        if (hits > melhor.hits) melhor = { idx: i, hits };
    }
    return melhor.hits >= 2 ? melhor.idx : -1;
}

/* Mapeamento manual salvo (por nome normalizado do cabeçalho) */
function lerMapaSalvo() {
    try { return JSON.parse(store.get(CONFIG.CHAVE_MAPA, '{}')) || {}; } catch (_) { return {}; }
}
function salvarMapaManual(campo, cabecalhoNorm) {
    const m = lerMapaSalvo();
    m[campo] = cabecalhoNorm;
    store.set(CONFIG.CHAVE_MAPA, JSON.stringify(m));
}

/* Fuzzy matching: para cada campo escolhe a coluna cujo nome casa com o padrão
   mais específico (respeitando a ordem dos padrões, não a ordem das colunas). */
function autoMapear(cabecalhos) {
    const norm = cabecalhos.map(normalizar);
    const salvo = lerMapaSalvo();
    const usados = new Set();
    const mapa = {};

    // 1) Escolhas manuais salvas têm prioridade
    for (const campo of CAMPOS) {
        if (!Object.prototype.hasOwnProperty.call(salvo, campo.id)) continue;
        if (salvo[campo.id] === '') { mapa[campo.id] = -1; continue; }
        const i = norm.indexOf(salvo[campo.id]);
        if (i >= 0) { mapa[campo.id] = i; usados.add(i); }
    }
    // 2) Detecção automática para o restante (uma coluna não é reaproveitada por outro campo)
    for (const campo of CAMPOS) {
        if (campo.id in mapa) continue;
        let melhor = { p: Infinity, i: -1 };
        norm.forEach((h, i) => {
            if (!h || usados.has(i)) return;
            if (campo.excluir && campo.excluir.test(h)) return;
            const p = campo.padroes.findIndex((rx) => rx.test(h));
            if (p >= 0 && p < melhor.p) melhor = { p, i };
        });
        mapa[campo.id] = melhor.i;
        if (melhor.i >= 0) usados.add(melhor.i);
    }
    return mapa;
}

function classificarStatus(texto) {
    const t = normalizar(texto);
    if (!t) return 'vazio';
    for (const c of CLASSES) if (c.rx.test(t)) return c.cat;
    return 'outro';
}

/* Transforma as linhas cruas em registros tipados (uma vez por ingestão) */
function construirRegistros() {
    const { linhas, mapa } = estado;
    const col = (r, campo) => { const i = mapa[campo]; return i >= 0 ? r[i] : undefined; };
    const semChaves = mapa.pdv < 0 && mapa.nome < 0;
    const ehTotal = (s) => /^(SUB ?)?TOTAL( GERAL)?$/.test(normalizar(s));
    const out = [];

    for (let n = 0; n < linhas.length; n++) {
        const r = linhas[n];
        if (!r) continue;
        const pdv = txt(col(r, 'pdv'));
        const nome = txt(col(r, 'nome'));
        if (!semChaves && !pdv && !nome) continue;      // linha sem identificação
        if (ehTotal(pdv) || ehTotal(nome)) continue;    // linha de total do relatório

        const status = txt(col(r, 'status'));
        const cat = classificarStatus(status);
        const gapRaw = txt(col(r, 'gap'));
        const gapNum = paraNumero(col(r, 'gap'));
        const gapPorColuna = gapNum !== null ? gapNum > 0 : classificarStatus(gapRaw) === 'gap';

        out.push({
            id: out.length, pdv, nome,
            setor: txt(col(r, 'setor')) || 'Sem setor',
            dono: txt(col(r, 'dono')),
            status, cat, gapRaw, gapNum,
            temGap: gapPorColuna || cat === 'gap',
            fatEspRaw: col(r, 'fatEsp'), fatRealRaw: col(r, 'fatReal'),
            busca: simples(`${pdv} ${nome}`)
        });
    }
    return out;
}

async function carregarAba(nome) {
    mostrarLoading('Processando base consolidada...', `Lendo a aba "${nome}"`, null);
    await frame();

    // Só a aba escolhida é decodificada: bem menos memória que ler o .xlsm inteiro
    const wb = XLSX.read(estado.buffer, { type: 'array', sheets: nome, dense: true, cellStyles: false, cellHTML: false });
    const ws = wb.Sheets[nome];
    if (!ws) throw new Error(`A aba "${nome}" está vazia ou não pôde ser lida.`);

    // Sem defval: linhas esparsas (não cria milhares de células vazias em planilhas "infladas")
    const matriz = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false });
    if (!matriz.length) throw new Error(`A aba "${nome}" não tem dados.`);

    const idx = detectarCabecalho(matriz);
    const posicao = idx >= 0 ? idx : 0;
    const cab = matriz[posicao] || [];

    estado.aba = nome;
    estado.cabecalhoDetectado = idx >= 0;
    estado.linhaCabecalho = posicao + 1;
    estado.cabecalhos = Array.from({ length: cab.length }, (_, i) => txt(cab[i]).replace(/\s+/g, ' ') || `Coluna ${i + 1}`);
    estado.linhas = matriz.slice(posicao + 1);
    estado.mapa = autoMapear(estado.cabecalhos);
    estado.selecionadoId = null;
    estado.regGrafico = null;
    resetarPainel();
    recalcular();
}

async function ingerir(file) {
    if (estado.ocupado) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
        toast('Formato não suportado. Use um arquivo .xlsx, .xlsm ou .xls.', 'erro');
        return;
    }
    estado.ocupado = true;
    mostrarLoading('Processando base consolidada...', 'Lendo o arquivo', 0);
    try {
        if (!(await garantirLib('XLSX'))) {
            throw new Error('A biblioteca SheetJS não carregou. Coloque xlsx.full.min.js na pasta vendor/ ou libere o acesso à CDN.');
        }
        // Uint8Array é a forma documentada do SheetJS para type: 'array' (é só uma "janela" sobre o mesmo buffer, sem cópia)
        estado.buffer = new Uint8Array(await lerArquivo(file, (p) => mostrarLoading('Processando base consolidada...', `Lendo o arquivo (${Math.round(p * 100)}%)`, p)));
        await frame();
        estado.arquivo = file.name;
        estado.abas = XLSX.read(estado.buffer, { type: 'array', bookSheets: true }).SheetNames || [];
        if (!estado.abas.length) throw new Error('O arquivo não tem abas.');
        await carregarAba(escolherAba(estado.abas));

        if (!estado.cabecalhoDetectado) {
            toast('Não achei a linha de cabeçalho automaticamente e usei a primeira linha. Abra o Diagnóstico para conferir a aba.', 'aviso', 9000);
        } else if (estado.mapa.pdv < 0 || estado.mapa.status < 0) {
            toast('Algumas colunas importantes não foram encontradas. Abra o Diagnóstico para escolher.', 'aviso', 9000);
        }
    } catch (err) {
        console.error(err);
        const msg = /password|encrypt/i.test(String(err && err.message))
            ? 'O arquivo está protegido por senha. Salve uma cópia sem senha e tente de novo.'
            : `Não foi possível processar o arquivo: ${err && err.message ? err.message : err}`;
        toast(msg, 'erro', 9000);
    } finally {
        estado.ocupado = false;
        esconderLoading();
    }
}

/* =====================================================================
   KPIs, AGREGAÇÃO E RENDERIZAÇÃO
   ===================================================================== */

function recalcular() {
    estado.registros = construirRegistros();
    renderKPIs();
    const n = estado.registros.length;
    dom.fileStatus.innerHTML = `<span class="dot green"></span> Base ativa: ${fmtInt(n)} PDVs, aba "${escapeHTML(estado.aba)}"`;
    dom.btnDiag.disabled = false;
    mudarVisao(estado.modo);
}

function renderKPIs() {
    const regs = estado.registros;
    const total = regs.length;
    let ok = 0, vz = 0, gap = 0;
    for (const r of regs) {
        if (r.cat === 'ok') ok++;
        if (r.cat === 'vz' || r.cat === 'nok') vz++;
        if (r.temGap) gap++;
    }
    const pct = (n) => total ? `${((n / total) * 100).toFixed(1).replace('.', ',')}% da base` : '\u00a0';
    dom.kpiTotal.textContent = fmtInt(total);
    dom.kpiOk.textContent = fmtInt(ok);
    dom.kpiVz.textContent = fmtInt(vz);
    dom.kpiGap.textContent = fmtInt(gap);
    dom.kpiTotalSub.textContent = total ? `aba "${estado.aba}"` : 'registros na base';
    dom.kpiOkSub.textContent = pct(ok);
    dom.kpiVzSub.textContent = pct(vz);
    dom.kpiGapSub.textContent = pct(gap);
}

function moda(mapaContagem) {
    let melhor = '', max = 0;
    for (const [k, n] of mapaContagem) if (n > max) { melhor = k; max = n; }
    return melhor;
}

function agruparPorSetor(regs) {
    const grupos = new Map();
    for (const r of regs) {
        let g = grupos.get(r.setor);
        if (!g) { g = { setor: r.setor, equip: 0, ok: 0, vz: 0, gap: 0, donos: new Map() }; grupos.set(r.setor, g); }
        g.equip++;
        if (r.cat === 'ok') g.ok++;
        if (r.cat === 'vz' || r.cat === 'nok') g.vz++;
        if (r.temGap) g.gap++;
        if (r.dono) g.donos.set(r.dono, (g.donos.get(r.dono) || 0) + 1);
    }
    return [...grupos.values()]
        .map((g) => ({ setor: g.setor, equip: g.equip, ok: g.ok, vz: g.vz, gap: g.gap, dono: moda(g.donos) || '—', ating: g.equip ? (g.ok / g.equip) * 100 : 0 }))
        .sort((a, b) => a.ating - b.ating || b.gap - a.gap);   // piores primeiro: é onde a cobrança começa
}

function renderVazio(colunas) {
    dom.tbody.innerHTML = `
        <tr><td colspan="${colunas}" class="empty-state">
            <i class="ri-file-excel-line" aria-hidden="true"></i>
            <p>Carregue o arquivo mestre para gerar a visão gerencial e a auditoria de PDVs.</p>
            <button class="upload-btn" data-upload type="button">Selecionar arquivo</button>
            <p class="hint">Você também pode arrastar o arquivo para esta janela.</p>
        </td></tr>`;
    dom.footer.classList.add('hidden');
}

function renderGerencial() {
    dom.thead.innerHTML = `
        <tr><th>Setor</th><th>Representante / dono</th><th class="num">Equipamentos</th><th class="num">Giro OK</th><th>Atingimento</th><th class="num">GAPs</th></tr>`;
    if (!estado.registros.length) { renderVazio(6); return; }

    estado.setores = agruparPorSetor(estado.registros);
    dom.tbody.innerHTML = estado.setores.map((s, i) => {
        const bate = s.ating >= CONFIG.META_ATINGIMENTO;
        return `<tr class="clicavel" data-setor-i="${i}" tabindex="0" title="Ver os PDVs deste setor">
            <td><strong>${escapeHTML(setorLabel(s.setor))}</strong></td>
            <td>${escapeHTML(s.dono)}</td>
            <td class="num">${fmtInt(s.equip)}</td>
            <td class="num">${fmtInt(s.ok)}</td>
            <td><div class="meter ${bate ? 'ok' : 'bad'}"><div class="meter-track"><div class="meter-bar" style="width:${Math.min(100, s.ating).toFixed(0)}%"></div></div><span>${s.ating.toFixed(0)}%</span></div></td>
            <td class="num"><span class="badge ${s.gap ? 'danger' : 'ok'}">${fmtInt(s.gap)} PDV${s.gap === 1 ? '' : 's'}</span></td>
        </tr>`;
    }).join('');

    dom.count.textContent = `${fmtInt(estado.setores.length)} setores, do menor para o maior atingimento. Meta: ${CONFIG.META_ATINGIMENTO}%.`;
    dom.btnMais.classList.add('hidden');
    dom.footer.classList.remove('hidden');
}

function badgeStatus(r) {
    if (r.cat !== 'outro' && r.cat !== 'vazio') return `<span class="badge ${CLASSE_CAT[r.cat]}">${ROTULO_CAT[r.cat]}</span>`;
    return r.status ? `<span class="badge neutral">${escapeHTML(r.status)}</span>` : '<span class="badge neutral">Sem status</span>';
}

const fmtGap = (r) => r.gapNum !== null ? fmtInt(r.gapNum) : (r.gapRaw || '—');

function celulaGap(r) {
    if (r.temGap) return `<span class="badge danger">${escapeHTML(fmtGap(r))}</span>`;
    return r.gapRaw ? `<span class="text-muted">${escapeHTML(fmtGap(r))}</span>` : '<span class="text-muted">—</span>';
}

function aplicarFiltros() {
    const termo = simples(dom.busca.value.trim());
    const cat = dom.filtroStatus.value;
    estado.filtrados = estado.registros.filter((r) => {
        if (termo && !r.busca.includes(termo)) return false;
        if (estado.filtroSetor !== null && r.setor !== estado.filtroSetor) return false;
        if (cat === 'todos') return true;
        if (cat === 'gap') return r.temGap;
        return r.cat === cat;
    });
    estado.visiveis = CONFIG.PAGINA;
    renderPDVs();
}

function renderPDVs() {
    dom.thead.innerHTML = `
        <tr><th>Cód. PDV</th><th>Nome fantasia</th><th>Setor</th><th>Status</th><th>GAP de SKU</th><th class="text-right">Ação</th></tr>`;
    if (!estado.registros.length) { renderVazio(6); return; }

    const lista = estado.filtrados;
    const fatia = lista.slice(0, estado.visiveis);
    if (!fatia.length) {
        dom.tbody.innerHTML = '<tr><td colspan="6" class="empty-state"><p>Nenhum PDV encontrado com esses filtros.</p></td></tr>';
    } else {
        dom.tbody.innerHTML = fatia.map((r) => `
            <tr data-id="${r.id}"${r.id === estado.selecionadoId ? ' class="selected"' : ''}>
                <td><strong>${escapeHTML(r.pdv || '—')}</strong></td>
                <td>${escapeHTML(r.nome || '—')}</td>
                <td>${escapeHTML(setorLabel(r.setor))}</td>
                <td>${badgeStatus(r)}</td>
                <td>${celulaGap(r)}</td>
                <td class="text-right"><button class="action-sm" data-auditar="${r.id}" type="button">Auditar</button></td>
            </tr>`).join('');
    }

    const mostrando = Math.min(estado.visiveis, lista.length);
    dom.count.textContent = lista.length
        ? `Mostrando ${fmtInt(mostrando)} de ${fmtInt(lista.length)} PDVs`
        : '0 PDVs';
    dom.btnMais.classList.toggle('hidden', lista.length <= estado.visiveis);
    dom.footer.classList.remove('hidden');
}

function atualizarChip() {
    const ativo = estado.filtroSetor !== null;
    dom.chipSetor.classList.toggle('hidden', !ativo);
    if (ativo) dom.chipSetor.textContent = `${setorLabel(estado.filtroSetor)} \u00d7`;
}

function mudarVisao(tipo) {
    estado.modo = tipo;
    const pdvs = tipo === 'pdvs';
    document.querySelectorAll('.nav-link[data-view]').forEach((l) => {
        const ativo = l.dataset.view === tipo;
        l.classList.toggle('active', ativo);
        if (ativo) l.setAttribute('aria-current', 'page'); else l.removeAttribute('aria-current');
    });
    dom.viewTitle.textContent = pdvs ? 'Auditoria analítica de PDVs' : 'Visão gerencial consolidada';
    dom.toolbar.classList.toggle('hidden', !pdvs || !estado.registros.length);
    dom.side.classList.toggle('hidden', !pdvs);
    dom.split.classList.toggle('solo', !pdvs);
    atualizarChip();
    if (pdvs) aplicarFiltros(); else renderGerencial();
}

/* =====================================================================
   RAIO-X DO PDV
   ===================================================================== */

function resetarPainel() {
    dom.detId.textContent = '---';
    dom.detContent.innerHTML = '<p class="text-muted">Clique em "Auditar" em qualquer linha para ver os dados do PDV, o faturamento e a pauta de cobrança.</p>';
    dom.chartWrap.classList.add('hidden');
    dom.chartNote.classList.add('hidden');
    if (estado.chart) { estado.chart.destroy(); estado.chart = null; }
}

function blocoFaturamento(r) {
    const esp = paraNumero(r.fatEspRaw), real = paraNumero(r.fatRealRaw);
    if (esp === null || real === null || esp <= 0) return '';
    const pct = (real / esp) * 100;
    const dif = real - esp;
    return `
        <div class="detail-row"><span>Atingimento do faturamento</span><strong class="${pct >= 100 ? 'pos' : 'neg'}">${pct.toFixed(0)}%</strong></div>
        <div class="meter ${pct >= 100 ? 'ok' : 'bad'} solo-meter"><div class="meter-track"><div class="meter-bar" style="width:${Math.min(100, pct).toFixed(0)}%"></div></div></div>
        <div class="detail-row"><span>Diferença</span><strong class="${dif >= 0 ? 'pos' : 'neg'}">${escapeHTML(BRL.format(dif))}</strong></div>`;
}

function auditar(id) {
    const r = estado.registros[id];
    if (!r) return;
    estado.selecionadoId = id;
    estado.regGrafico = r;

    dom.tbody.querySelectorAll('tr.selected').forEach((tr) => tr.classList.remove('selected'));
    const linha = dom.tbody.querySelector(`tr[data-id="${id}"]`);
    if (linha) linha.classList.add('selected');

    dom.detId.textContent = `PDV ${r.pdv || '—'}`;
    dom.detContent.innerHTML = `
        <div class="detail-row"><span>Cliente</span><strong>${escapeHTML(r.nome || '—')}</strong></div>
        <div class="detail-row"><span>Setor</span><strong>${escapeHTML(setorLabel(r.setor))}</strong></div>
        <div class="detail-row"><span>Representante</span><strong>${escapeHTML(r.dono || '—')}</strong></div>
        <div class="detail-row"><span>Status geral</span>${badgeStatus(r)}</div>
        <div class="detail-row"><span>GAP de SKU</span><strong>${escapeHTML(fmtGap(r))}</strong></div>
        <div class="detail-sep"></div>
        <div class="detail-row"><span>Faturamento esperado</span><strong>${escapeHTML(fmtValor(r.fatEspRaw))}</strong></div>
        <div class="detail-row"><span>Faturamento real</span><strong>${escapeHTML(fmtValor(r.fatRealRaw))}</strong></div>
        ${blocoFaturamento(r)}
        <button class="upload-btn w-full" id="btnPauta" type="button"><i class="ri-clipboard-line" aria-hidden="true"></i> Copiar pauta de cobrança</button>`;
    $('btnPauta').addEventListener('click', () => copiarPauta(r));

    desenharGrafico(r);
}

/* Curva simulada, mas estável: a mesma semente (código do PDV) gera sempre o mesmo desenho */
function hashTexto(s) {
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return h >>> 0;
}
function mulberry32(a) {
    return function () {
        a |= 0; a = (a + 0x6D2B79F5) | 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}
function serieSimulada(r, n = 6) {
    const rnd = mulberry32(hashTexto(`${r.pdv}|${r.nome}`));
    const base = { ok: 88, gap: 58, nok: 42, vz: 20, outro: 65, vazio: 65 }[r.cat];
    const lim = (v) => Math.max(5, Math.min(100, v));
    let v = lim(base + (rnd() - 0.5) * 30);
    const pontos = [];
    for (let i = 0; i < n; i++) {
        v = lim(v + (base - v) * 0.35 + (rnd() - 0.5) * 14);
        pontos.push(Math.round(v));
    }
    return pontos;
}
function rotulosMeses(n = 6) {
    const hoje = new Date();
    return Array.from({ length: n }, (_, i) => {
        const d = new Date(hoje.getFullYear(), hoje.getMonth() - (n - 1 - i), 1);
        return d.toLocaleDateString('pt-BR', { month: 'short' }).replace('.', '');
    });
}
function corComAlpha(cor, alpha) {
    const m = /^#([0-9a-f]{6})$/i.exec(cor);
    if (!m) return cor;
    const n = parseInt(m[1], 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

async function desenharGrafico(r) {
    dom.chartWrap.classList.remove('hidden');
    dom.chartNote.classList.remove('hidden');
    dom.chartNote.textContent = 'Curva ilustrativa: ainda não usa histórico real do arquivo.';

    if (!(await garantirLib('Chart'))) {
        dom.chartWrap.classList.add('hidden');
        dom.chartNote.textContent = 'Gráfico indisponível: o Chart.js não carregou (veja a pasta vendor/).';
        return;
    }
    if (estado.chart) { estado.chart.destroy(); estado.chart = null; }
    if (window.Chart.getChart) { const antigo = window.Chart.getChart(dom.canvas); if (antigo) antigo.destroy(); }

    const css = getComputedStyle(document.documentElement);
    const cor = (v) => css.getPropertyValue(v).trim();
    const reduzir = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    estado.chart = new window.Chart(dom.canvas.getContext('2d'), {
        type: 'line',
        data: {
            labels: rotulosMeses(),
            datasets: [{
                label: 'Cobertura de SKU (simulada)',
                data: serieSimulada(r),
                borderColor: cor('--primary'),
                backgroundColor: corComAlpha(cor('--primary'), 0.15),
                fill: true, tension: 0.35, pointRadius: 3, borderWidth: 2
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            animation: reduzir ? false : { duration: 350 },
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => `${c.parsed.y}%` } } },
            scales: {
                y: { min: 0, max: 100, ticks: { color: cor('--text-muted'), callback: (v) => `${v}%` }, grid: { color: cor('--border') } },
                x: { ticks: { color: cor('--text-muted') }, grid: { display: false } }
            }
        }
    });
}

/* ---------- Pauta de cobrança ---------- */
function montarPauta(r) {
    const saudacao = r.dono ? `Olá, ${r.dono}!` : 'Olá!';
    return [
        'PAUTA DE COBRANÇA: GAP DE SKU',
        `${saudacao} Precisamos alinhar o seguinte PDV:`,
        '',
        `PDV: ${r.pdv || '—'} - ${r.nome || '—'}`,
        `Setor: ${setorLabel(r.setor)}`,
        `Status: ${r.status || '—'}`,
        `GAP de SKU: ${fmtGap(r)}`,
        `Faturamento: esperado ${fmtValor(r.fatEspRaw)} | real ${fmtValor(r.fatRealRaw)}`,
        '',
        'Pode confirmar o plano de ação e o prazo para regularizar? Obrigado!'
    ].join('\n');
}

async function copiarTexto(texto) {
    try {
        if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(texto); return true; }
    } catch (_) { /* cai no plano B */ }
    try {   // plano B para contextos sem Clipboard API (ex.: arquivo aberto via file://)
        const ta = document.createElement('textarea');
        ta.value = texto;
        ta.setAttribute('readonly', '');
        ta.style.cssText = 'position:fixed;top:-1000px;opacity:0';
        document.body.appendChild(ta);
        ta.select();
        const ok = document.execCommand('copy');
        ta.remove();
        return ok;
    } catch (_) { return false; }
}

async function copiarPauta(r) {
    const ok = await copiarTexto(montarPauta(r));
    if (ok) toast('Pauta copiada. É só colar no Teams ou no WhatsApp.', 'ok');
    else toast('Não consegui copiar automaticamente. Selecione e copie o texto manualmente.', 'erro', 7000);
}

/* =====================================================================
   DIAGNÓSTICO (transparência do parser + ajuste manual de colunas)
   ===================================================================== */

function exemploColuna(i) {
    if (i < 0) return '';
    const limite = Math.min(estado.linhas.length, 200);
    for (let n = 0; n < limite; n++) {
        const v = txt(estado.linhas[n] && estado.linhas[n][i]);
        if (v) return v.length > 40 ? `${v.slice(0, 40)}...` : v;
    }
    return '';
}

function renderDiagStatus() {
    const cont = new Map();
    for (const r of estado.registros) {
        const chave = r.status || '(vazio)';
        const e = cont.get(chave) || { n: 0, cat: r.cat };
        e.n++;
        cont.set(chave, e);
    }
    const linhas = [...cont].sort((a, b) => b[1].n - a[1].n);
    const topo = linhas.slice(0, CONFIG.MAX_STATUS_DIAG);
    dom.diagStatus.innerHTML = topo.map(([valor, e]) => `
        <tr><td>${escapeHTML(valor)}</td><td><span class="badge ${CLASSE_CAT[e.cat]}">${ROTULO_CAT[e.cat]}</span></td><td class="num">${fmtInt(e.n)}</td></tr>`).join('')
        + (linhas.length > topo.length ? `<tr><td colspan="3" class="text-muted">+ ${linhas.length - topo.length} valores menos frequentes</td></tr>` : '')
        || '<tr><td colspan="3" class="text-muted">Nenhum registro.</td></tr>';
}

function renderDiagnostico() {
    dom.diagSheet.innerHTML = estado.abas.map((a) => `<option value="${escapeHTML(a)}"${a === estado.aba ? ' selected' : ''}>${escapeHTML(a)}</option>`).join('');
    const cab = estado.cabecalhoDetectado
        ? `Cabeçalho detectado na ${estado.linhaCabecalho}ª linha com conteúdo.`
        : 'Cabeçalho não detectado: usei a primeira linha.';
    dom.diagInfo.textContent = `Arquivo: ${estado.arquivo}. ${cab} ${fmtInt(estado.linhas.length)} linhas lidas, ${fmtInt(estado.registros.length)} PDVs válidos.`;

    dom.diagMap.innerHTML = CAMPOS.map((c) => {
        const sel = estado.mapa[c.id];
        const opcoes = ['<option value="-1">Não usar</option>']
            .concat(estado.cabecalhos.map((h, i) => `<option value="${i}"${i === sel ? ' selected' : ''}>${escapeHTML(h)}</option>`)).join('');
        return `<tr>
            <td><strong>${c.rotulo}</strong></td>
            <td><select class="select" data-campo="${c.id}">${opcoes}</select></td>
            <td class="text-muted" data-exemplo="${c.id}">${escapeHTML(exemploColuna(sel)) || '—'}</td>
        </tr>`;
    }).join('');
    renderDiagStatus();
}

function abrirDiagnostico() {
    if (!estado.abas.length) return;
    renderDiagnostico();
    dom.modal.classList.remove('hidden');
    dom.diagOk.focus();
}
function fecharDiagnostico() {
    dom.modal.classList.add('hidden');
    dom.btnDiag.focus();
}

/* =====================================================================
   EVENTOS
   ===================================================================== */

dom.themeToggle.addEventListener('click', () => {
    aplicarTema(document.documentElement.getAttribute('data-theme') === 'light' ? 'dark' : 'light');
});

dom.btnUpload.addEventListener('click', () => dom.fileInput.click());
dom.fileInput.addEventListener('change', (e) => {
    const arquivo = e.target.files && e.target.files[0];
    e.target.value = '';      // permite escolher o mesmo arquivo de novo
    if (arquivo) ingerir(arquivo);
});

document.querySelectorAll('.nav-link[data-view]').forEach((a) => {
    a.addEventListener('click', (e) => {
        e.preventDefault();
        if (a.dataset.view !== estado.modo) estado.filtroSetor = null;
        mudarVisao(a.dataset.view);
    });
});

/* Tabela: delegação de eventos (nada de onclick inline com dados do Excel) */
dom.tbody.addEventListener('click', (e) => {
    if (e.target.closest('[data-upload]')) { dom.fileInput.click(); return; }
    const btn = e.target.closest('[data-auditar]');
    if (btn) { auditar(Number(btn.dataset.auditar)); return; }
    const linha = e.target.closest('tr[data-setor-i]');
    if (linha) irParaSetor(Number(linha.dataset.setorI));
});
dom.tbody.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const linha = e.target.closest('tr[data-setor-i]');
    if (linha) irParaSetor(Number(linha.dataset.setorI));
});
function irParaSetor(i) {
    const s = estado.setores[i];
    if (!s) return;
    estado.filtroSetor = s.setor;
    mudarVisao('pdvs');
}

let temporizadorBusca = null;
dom.busca.addEventListener('input', () => {
    clearTimeout(temporizadorBusca);
    temporizadorBusca = setTimeout(() => { if (estado.modo === 'pdvs') aplicarFiltros(); }, 150);
});
dom.filtroStatus.addEventListener('change', () => { if (estado.modo === 'pdvs') aplicarFiltros(); });
dom.chipSetor.addEventListener('click', () => { estado.filtroSetor = null; atualizarChip(); aplicarFiltros(); });
dom.btnMais.addEventListener('click', () => { estado.visiveis += CONFIG.PAGINA; renderPDVs(); });

/* Diagnóstico */
dom.btnDiag.addEventListener('click', abrirDiagnostico);
dom.diagClose.addEventListener('click', fecharDiagnostico);
dom.diagOk.addEventListener('click', fecharDiagnostico);
dom.modal.addEventListener('click', (e) => { if (e.target === dom.modal) fecharDiagnostico(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !dom.modal.classList.contains('hidden')) fecharDiagnostico(); });

dom.diagMap.addEventListener('change', (e) => {
    const sel = e.target.closest('select[data-campo]');
    if (!sel) return;
    const campo = sel.dataset.campo;
    const idx = Number(sel.value);
    estado.mapa[campo] = idx;
    salvarMapaManual(campo, idx >= 0 ? normalizar(estado.cabecalhos[idx]) : '');
    recalcular();
    const celula = dom.diagMap.querySelector(`[data-exemplo="${campo}"]`);
    if (celula) celula.textContent = exemploColuna(idx) || '—';
    renderDiagStatus();
    dom.diagInfo.textContent = dom.diagInfo.textContent.replace(/[\d.]+ PDVs válidos/, `${fmtInt(estado.registros.length)} PDVs válidos`);
});

dom.diagSheet.addEventListener('change', async () => {
    if (!dom.diagSheet.value) return;
    try {
        await carregarAba(dom.diagSheet.value);
        renderDiagnostico();
    } catch (err) {
        console.error(err);
        toast(`Não foi possível ler essa aba: ${err.message}`, 'erro', 8000);
    } finally {
        esconderLoading();
    }
});

dom.diagReset.addEventListener('click', () => {
    store.set(CONFIG.CHAVE_MAPA, '{}');
    estado.mapa = autoMapear(estado.cabecalhos);
    recalcular();
    renderDiagnostico();
    toast('Detecção automática restaurada.', 'ok');
});

/* Arrastar e soltar em qualquer lugar da janela */
const temArquivo = (e) => !!e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
let profundidadeDrag = 0;
window.addEventListener('dragenter', (e) => {
    if (!temArquivo(e)) return;
    e.preventDefault();
    profundidadeDrag++;
    dom.dropHint.classList.remove('hidden');
});
window.addEventListener('dragleave', (e) => {
    if (!temArquivo(e)) return;
    profundidadeDrag = Math.max(0, profundidadeDrag - 1);
    if (!profundidadeDrag) dom.dropHint.classList.add('hidden');
});
window.addEventListener('dragover', (e) => { if (temArquivo(e)) e.preventDefault(); });
window.addEventListener('drop', (e) => {
    if (!temArquivo(e)) return;
    e.preventDefault();
    profundidadeDrag = 0;
    dom.dropHint.classList.add('hidden');
    const arquivo = e.dataTransfer.files && e.dataTransfer.files[0];
    if (arquivo) ingerir(arquivo);
});

/* Rede de segurança: nenhum erro fica silencioso */
window.addEventListener('unhandledrejection', (e) => {
    console.error(e.reason);
    toast(`Erro inesperado: ${e.reason && e.reason.message ? e.reason.message : e.reason}`, 'erro', 8000);
});
window.addEventListener('error', (e) => {
    console.error(e.error || e.message);
    toast(`Erro inesperado: ${e.message}`, 'erro', 8000);
});

/* ---------- Início ---------- */
aplicarTema(store.get('theme', 'dark'));
mudarVisao('gerencial');
garantirLib('XLSX');    // pré-carrega em segundo plano para o primeiro upload ser imediato
