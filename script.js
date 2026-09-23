// Prevenção XSS
const escapeHTML = (str) => {
    if (!str) return '';
    return str.toString().replace(/[&<>'"]/g, 
        tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag])
    );
};

// Gerenciamento de Tema (Light/Dark) com localStorage
const themeToggle = document.getElementById('theme-toggle');
const rootElement = document.documentElement;

const currentTheme = localStorage.getItem('theme') || 'light';
rootElement.setAttribute('data-theme', currentTheme);
updateThemeIcon(currentTheme);

themeToggle.addEventListener('click', () => {
    const newTheme = rootElement.getAttribute('data-theme') === 'light' ? 'dark' : 'light';
    rootElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    updateThemeIcon(newTheme);
    if(chartInstance) chartInstance.update(); // Atualiza cores do gráfico
});

function updateThemeIcon(theme) {
    const icon = themeToggle.querySelector('i');
    icon.className = theme === 'light' ? 'ri-moon-line' : 'ri-sun-line';
}

// Mobile Menu Toggle
document.getElementById('mobile-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.add('open');
});
document.getElementById('mobile-close').addEventListener('click', () => {
    document.getElementById('sidebar').classList.remove('open');
});

// Banco de Dados em Memória
let baseOperacional = [];
let chartInstance = null;

const searchInput = document.getElementById('pdvSearch');
const clearBtn = document.getElementById('clearSearch');
const loadingOverlay = document.getElementById('loading-overlay');
const systemStatus = document.getElementById('system-status');
const emptyWorkspace = document.getElementById('empty-workspace');
const mainGrid = document.getElementById('main-grid');

// Leitura do Ficheiro Excel (.xlsm / .xlsx)
document.getElementById('excelFileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().match(/\.(xlsx|xlsm|xls)$/)) {
        alert("Aviso de Segurança: Formato inválido. Por favor, importe a base oficial do BI.");
        return;
    }

    loadingOverlay.classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Busca Inteligente de Aba
            let abaPrincipal = workbook.SheetNames.find(nome => nome.includes('BI de Equipamentos') || nome.includes('SKU-PDV')) || workbook.SheetNames[0];

            baseOperacional = XLSX.utils.sheet_to_json(workbook.Sheets[abaPrincipal], { defval: "" });

            // UI Feedback
            loadingOverlay.classList.add('hidden');
            searchInput.disabled = false;
            searchInput.placeholder = "Digite o código do PDV ou Razão Social...";
            
            systemStatus.innerHTML = `<span class="pulse-dot green"></span> Base Ativa (${baseOperacional.length} registros)`;
            emptyWorkspace.classList.remove('hidden');
            emptyWorkspace.innerHTML = `<div class="empty-icon"><i class="ri-search-eye-line"></i></div><h3>Sistema Sincronizado</h3><p>Base lida com sucesso. Utilize a barra de pesquisa acima para analisar um PDV específico.</p>`;

            searchInput.focus();

        } catch (error) {
            loadingOverlay.classList.add('hidden');
            alert("Erro na descodificação: O arquivo pode estar corrompido.");
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
});

// Mecanismo de Busca
searchInput.addEventListener('input', (e) => {
    const term = e.target.value.trim();
    if (term.length > 0) {
        clearBtn.classList.remove('hidden');
    } else {
        clearBtn.classList.add('hidden');
        resetDashboard();
    }
});

searchInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') executarBuscaProfunda(searchInput.value.trim());
});

clearBtn.addEventListener('click', () => {
    searchInput.value = '';
    clearBtn.classList.add('hidden');
    resetDashboard();
});

// Lógica Relacional e Atualização do DOM
function executarBuscaProfunda(termoBusca) {
    if (baseOperacional.length === 0) return;

    const pdvEncontrado = baseOperacional.find(row => {
        const pdvStr = String(row['PDV'] || row['Cód. PDV'] || row['Código Cliente'] || '');
        const nomeStr = String(row['Nome Fantasia'] || row['Razão'] || row['Razao Social'] || '').toLowerCase();
        return pdvStr === termoBusca || nomeStr.includes(termoBusca.toLowerCase());
    });

    if (!pdvEncontrado) {
        alert("PDV não localizado na base ativa atual.");
        return;
    }

    renderizarPDV(pdvEncontrado);
}

function renderizarPDV(dados) {
    emptyWorkspace.classList.add('hidden');
    mainGrid.classList.remove('hidden');

    const pdv = escapeHTML(dados['PDV'] || dados['Cód. PDV'] || dados['Código Cliente'] || '---');
    const nome = escapeHTML(dados['Nome Fantasia'] || dados['Razão'] || dados['Razao Social'] || '---');
    const setor = escapeHTML(dados['Setor'] || dados['Cod. Setor'] || dados['GV'] || '---');
    const status = escapeHTML(String(dados['Status SKU'] || dados['Status PDV'] || dados['Status'] || 'OK').toUpperCase());
    
    let gapField = Object.keys(dados).find(k => k.toUpperCase().includes('GAP'));
    const gap = parseInt(dados[gapField] || '0');
    
    const isGap = status.includes("GAP") || status.includes("NOK") || gap > 0;
    const gapNum = isNaN(gap) ? 0 : Math.abs(gap);
    const skuAtual = isGap ? Math.max(0, 10 - gapNum) : (10 + gapNum); 

    // Update Profile Card
    document.getElementById('pdv-id-display').innerText = pdv;
    document.getElementById('pdv-title').innerText = nome.substring(0, 25) + (nome.length > 25 ? '...' : '');
    document.getElementById('det-setor').innerText = `Setor ${setor}`;
    document.getElementById('det-status').innerText = status;
    document.getElementById('det-status').style.color = isGap ? 'var(--danger)' : 'var(--success)';
    document.getElementById('det-gap').innerText = isGap ? `${gapNum} SKU(s)` : 'Meta Atingida';
    
    const btnCobrar = document.getElementById('btn-cobrar');
    btnCobrar.style.display = isGap ? 'flex' : 'none';
    btnCobrar.onclick = () => alert(`Integração: Rota do Setor ${setor} notificada sobre PDV ${pdv}.`);

    // Update AI Card
    const aiText = document.getElementById('ai-text');
    if (isGap) {
        aiText.innerHTML = `O cliente <strong>${nome}</strong> apresenta inaderência crítica. É necessária a positivação de <strong>${gapNum} SKUs</strong> para atingir o target de diversificação. Ação recomendada junto ao vendedor do Setor ${setor}.`;
    } else {
        aiText.innerHTML = `O cliente <strong>${nome}</strong> apresenta excelente saúde de sortimento. Nenhuma intervenção na geladeira SOPI é necessária no momento.`;
    }

    renderizarGraficoEvolucao(skuAtual);
}

function renderizarGraficoEvolucao(skuAtual) {
    const ctx = document.getElementById('evolutionChart')?.getContext('2d');
    if (!ctx) return;
    if (chartInstance) chartInstance.destroy();

    const mockHistory = [ Math.max(skuAtual - 3, 2), Math.max(skuAtual - 1, 3), skuAtual + 1, skuAtual - 2, skuAtual, skuAtual ];
    const isDark = rootElement.getAttribute('data-theme') === 'dark';
    const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
    const textColor = isDark ? '#9ca3af' : '#64748b';

    chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: ['M-5', 'M-4', 'M-3', 'M-2', 'M-1', 'Atual'],
            datasets: [{
                label: 'SKUs Comprados',
                data: mockHistory,
                borderColor: '#3b82f6',
                backgroundColor: 'rgba(59, 130, 246, 0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.4
            }, {
                label: 'Meta (10)',
                data: [10, 10, 10, 10, 10, 10],
                borderColor: '#f59e0b',
                borderDash: [5, 5],
                pointRadius: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, max: 15, grid: { color: gridColor }, ticks: { color: textColor } },
                x: { grid: { display: false }, ticks: { color: textColor } }
            }
        }
    });
}

function resetDashboard() {
    mainGrid.classList.add('hidden');
    emptyWorkspace.classList.remove('hidden');
}