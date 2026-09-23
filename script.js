const escapeHTML = (str) => {
    if (!str) return '';
    return str.toString().replace(/[&<>'"]/g, 
        tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag])
    );
};

// Gerenciamento de Tema
const themeToggle = document.getElementById('theme-toggle');
const rootElement = document.documentElement;
const currentTheme = localStorage.getItem('theme') || 'dark';
rootElement.setAttribute('data-theme', currentTheme);
updateThemeIcon(currentTheme);

themeToggle.addEventListener('click', () => {
    const newTheme = rootElement.getAttribute('data-theme') === 'light' ? 'dark' : 'light';
    rootElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    updateThemeIcon(newTheme);
    if(chartInstance) chartInstance.update();
});

function updateThemeIcon(theme) {
    themeToggle.querySelector('i').className = theme === 'light' ? 'ri-moon-line' : 'ri-sun-line';
}

let rawDataPDVs = [];
let modoAtual = 'gerencial';
let chartInstance = null;

const loadingOverlay = document.getElementById('loadingOverlay');
const fileStatus = document.getElementById('file-status');
const tableSearch = document.getElementById('tableSearch');
const toolbarSearch = document.getElementById('toolbar-search');

// Leitura Segura do Excel Mestre (Client-Side Memory Parser)
document.getElementById('excelFileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    loadingOverlay.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Varre dinamicamente procurando a aba de equipamentos
            let targetSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('BI') || n.toUpperCase().includes('PDV') || n.toUpperCase().includes('EQUIPAMENTOS')) || workbook.SheetNames[0];
            
            rawDataPDVs = XLSX.utils.sheet_to_json(workbook.Sheets[targetSheet], { defval: "" });

            loadingOverlay.classList.add('hidden');
            fileStatus.innerHTML = `<span class="dot green"></span> Pipeline Ativo (${rawDataPDVs.length} rows)`;

            calcularKPIs(rawDataPDVs);
            mudarVisao('gerencial');

        } catch (error) {
            loadingOverlay.classList.add('hidden');
            alert("Erro crítico no Parsing Engine do Excel.");
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

function calcularKPIs(data) {
    let total = data.length;
    let okCount = 0, vzCount = 0, gapCount = 0;

    data.forEach(r => {
        const status = String(r['Status \nPDV'] || r['Status PDV'] || r['Status SKU'] || r['Status do\nPDV'] || '').toUpperCase();
        if (status.includes('OK') || status.includes('BATEU') || status.includes('OVER')) okCount++;
        if (status.includes('VENDA ZERO') || status.includes('NOK')) vzCount++;
        if (status.includes('GAP')) gapCount++;
    });

    document.getElementById('kpi-total').innerText = total.toLocaleString('pt-BR');
    document.getElementById('kpi-ok').innerText = okCount.toLocaleString('pt-BR');
    document.getElementById('kpi-vz').innerText = vzCount.toLocaleString('pt-BR');
    document.getElementById('kpi-gap').innerText = gapCount.toLocaleString('pt-BR');
}

function mudarVisao(tipo, event) {
    if(event) event.preventDefault();
    modoAtual = tipo;

    document.querySelectorAll('.sidebar-menu .nav-link').forEach(l => l.classList.remove('active'));

    if (tipo === 'gerencial') {
        document.getElementById('view-title').innerText = "Visão Gerencial Consolidada (Setores)";
        toolbarSearch.style.display = 'none';
        if(event) event.currentTarget.classList.add('active');
        renderizarVisaoSetorial();
    } else {
        document.getElementById('view-title').innerText = "Auditoria Analítica de PDVs";
        toolbarSearch.style.display = 'flex';
        if(event) event.currentTarget.classList.add('active');
        renderizarTabelaPDVs(rawDataPDVs);
    }
}

// Agrupamento vetorial por Setor (In-Memory Aggregation)
function renderizarVisaoSetorial() {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    thead.innerHTML = `
        <tr>
            <th>Setor</th>
            <th>Responsável</th>
            <th>Equipamentos</th>
            <th>Giro OK</th>
            <th>Atingimento</th>
            <th class="text-right">GAP Operacional</th>
        </tr>
    `;
    tbody.innerHTML = '';

    if (rawDataPDVs.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-state"><p>Nenhum dado carregado na memória.</p></td></tr>`;
        return;
    }

    let setoresMap = {};
    rawDataPDVs.forEach(r => {
        const setor = String(r['Setor'] || r['GV'] || r['Cod. Setor'] || 'Geral');
        const dono = String(r['Supercom'] || r['Comercial'] || r['Representante \nde vendas'] || 'Equipe');
        const status = String(r['Status \nPDV'] || r['Status PDV'] || '').toUpperCase();

        if (!setoresMap[setor]) setoresMap[setor] = { dono, equip: 0, ok: 0, gap: 0 };
        setoresMap[setor].equip++;
        if (status.includes('OK') || status.includes('OVER')) setoresMap[setor].ok++;
        if (status.includes('GAP')) setoresMap[setor].gap++;
    });

    const fragment = document.createDocumentFragment();
    for (const [setor, info] of Object.entries(setoresMap)) {
        let meta = info.equip > 0 ? Math.round((info.ok / info.equip) * 100) : 0;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>Setor ${escapeHTML(setor)}</strong></td>
            <td>${escapeHTML(info.dono)}</td>
            <td>${info.equip}</td>
            <td>${info.ok}</td>
            <td><span style="color: ${meta >= 70 ? 'var(--success)' : 'var(--danger)'}; font-weight: bold;">${meta}%</span></td>
            <td class="text-right"><span class="badge danger">${info.gap} PDVs</span></td>
        `;
        fragment.appendChild(tr);
    }
    tbody.appendChild(fragment);
}

function renderizarTabelaPDVs(data) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    thead.innerHTML = `
        <tr>
            <th>Cód. PDV</th>
            <th>Nome Fantasia</th>
            <th>Setor</th>
            <th>Status PDV</th>
            <th>Status SKU</th>
            <th class="text-right">Ação</th>
        </tr>
    `;
    tbody.innerHTML = '';

    const sliceData = data.slice(0, 200);
    const fragment = document.createDocumentFragment();

    sliceData.forEach(r => {
        const pdv = escapeHTML(r['PDV'] || r['Cód. PDV'] || '---');
        const nome = escapeHTML(r['Nome Fantasia'] || r['Razão'] || '---');
        const setor = escapeHTML(r['Setor'] || r['GV'] || '---');
        const statusPDV = escapeHTML(String(r['Status \nPDV'] || r['Status PDV'] || 'Normal'));
        const statusSKU = escapeHTML(String(r['Status SKU'] || r['Status/SKU'] || 'OK'));

        let badgeClass = 'ok';
        if (statusPDV.toUpperCase().includes('GAP') || statusSKU.toUpperCase().includes('GAP')) badgeClass = 'danger';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><strong>${pdv}</strong></td>
            <td>${nome}</td>
            <td>Setor ${setor}</td>
            <td><span class="badge ${badgeClass}">${statusPDV}</span></td>
            <td>${statusSKU}</td>
            <td class="text-right">
                <button class="action-sm" onclick='selecionarPDV(${JSON.stringify(r)})'>Auditar</button>
            </td>
        `;
        fragment.appendChild(tr);
    });
    tbody.appendChild(fragment);
}

tableSearch.addEventListener('input', (e) => {
    if (modoAtual !== 'pdvs') return;
    const term = e.target.value.toLowerCase();
    const filtered = rawDataPDVs.filter(r => {
        const pdv = String(r['PDV'] || r['Cód. PDV'] || '');
        const nome = String(r['Nome Fantasia'] || r['Razão'] || '').toLowerCase();
        return pdv.includes(term) || nome.includes(term);
    });
    renderizarTabelaPDVs(filtered);
});

function selecionarPDV(r) {
    const pdv = escapeHTML(r['PDV'] || r['Cód. PDV'] || '---');
    const nome = escapeHTML(r['Nome Fantasia'] || r['Razão'] || '---');
    const setor = escapeHTML(r['Setor'] || r['GV'] || '---');
    const status = escapeHTML(String(r['Status \nPDV'] || r['Status PDV'] || 'OK'));
    
    const fatEsperado = r['Faturamento Esperado\nSOPI'] || r['Faturamento Esperado'] || 'R$ 0,00';
    const fatPDV = r['Faturamento PDV\nSOPI'] || r['Faturamento PDV'] || 'R$ 0,00';

    document.getElementById('det-id').innerText = `PDV: ${pdv}`;
    document.getElementById('det-content').innerHTML = `
        <div class="detail-row"><span>Cliente:</span> <strong>${nome}</strong></div>
        <div class="detail-row"><span>Setor:</span> <strong>Setor ${setor}</strong></div>
        <div class="detail-row"><span>Status:</span> <strong>${status}</strong></div>
        <div class="detail-row"><span>Fat. Esperado:</span> <strong style="color: var(--warning);">${fatEsperado}</strong></div>
        <div class="detail-row"><span>Fat. Real:</span> <strong style="color: var(--success);">${fatPDV}</strong></div>
        <div class="mt-4">
            <button class="action-sm w-full" style="background: var(--primary); color: white; border: none; padding: 0.5rem;" onclick="copiarPautaTeams('${pdv}', '${nome}', '${setor}')">
                <i class="ri-clipboard-line"></i> Copiar Pauta p/ Teams
            </button>
        </div>
    `;

    document.getElementById('chartContainer').classList.remove('hidden');
    renderizarGraficoLateral(8);
}

function copiarPautaTeams(pdv, nome, setor) {
    const texto = `*[Pauta Operacional SOPI]* \nOlá! Verificação necessária no PDV ${pdv} - ${nome} (Setor ${setor}). Constatada inaderência nos indicadores de sortimento.`;
    navigator.clipboard.writeText(texto);
    alert("Pauta de cobrança copiada para a área de transferência! Cole diretamente no Teams do RN.");
}

function limparCacheLocal() {
    localStorage.clear();
    location.reload();
}
