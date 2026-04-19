// --- CONFIGURAÇÕES ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbzrbc6xqFhpqRw2U9_1T4_rhscRJWTWlQPsCFH_5JM5Kedlq-DJj5IPpTkG3m9zcaHB2Q/exec"; 

const USUARIOS = [
    {login: "mateus", senha: "123", nome: "Mateus Lima", role: "digitador"},
    {login: "crislene", senha: "456", nome: "Crislene", role: "gestor"}
];

let usuarioAtual = null;
let abaAtual = "Digitadas";
let listas = { "Digitadas": [], "Recebimento": [], "Adiantamento": [] };

// --- CORREÇÃO DE DATA (Evita o erro de um dia a menos nos inputs) ---
function formatarDataParaExibir(dataISO) {
    if (!dataISO) return "---";
    const partes = dataISO.split('-'); // Espera AAAA-MM-DD
    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

// --- TEMA (DARK MODE) ---
function toggleTheme() {
    const b = document.body;
    const isDark = b.getAttribute('data-theme') === 'dark';
    b.setAttribute('data-theme', isDark ? 'light' : 'dark');
    document.getElementById('themeIcon').className = isDark ? 'ph ph-moon' : 'ph ph-sun';
}

// --- LOGIN ---
function realizarLogin() {
    const u = document.getElementById('userInput').value.toLowerCase();
    const s = document.getElementById('passInput').value;
    const user = USUARIOS.find(x => x.login === u && x.senha === s);

    if (user) {
        usuarioAtual = user;
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('mainHeader').style.display = 'flex';
        document.getElementById('app').style.display = 'block';
        document.getElementById('mainFooter').style.display = 'block';
        document.getElementById('userDisplay').innerText = usuarioAtual.nome.toUpperCase();
        
        switchTab('Digitadas'); // Inicia na aba de digitação
    } else {
        alert("Acesso negado: Usuário ou senha incorretos.");
    }
}

function logout() { location.reload(); }

// --- DASHBOARD: INDICADORES E MONITORAMENTO ---
async function carregarEstatisticas() {
    // Ativa o Loading e esconde o conteúdo
    document.getElementById('dash-loading').style.display = 'flex';
    document.getElementById('dash-content').style.display = 'none';

    try {
        const res = await fetch(URL_SCRIPT);
        const data = await res.json();

        // 1. Preenche Indicadores Principais (Soma das médias e Top Forn)
        document.getElementById('setorMediaGeral').innerText = data.statsSetor.mediaGeral;
        document.getElementById('setorForn').innerText = data.statsSetor.topForn;

        // 2. Preenche Monitoramento Global de Adiantamentos
        const tbodyAdi = document.querySelector("#tabelaMonitorAdi tbody");
        tbodyAdi.innerHTML = "";
        
        if (data.adiantamentosSetor && data.adiantamentosSetor.length > 0) {
            const hoje = new Date();
            hoje.setHours(0, 0, 0, 0);

            // Ordena por data (vencimento mais próximo primeiro)
            data.adiantamentosSetor.sort((a, b) => new Date(a.venc) - new Date(b.venc));

            data.adiantamentosSetor.forEach(adi => {
                const dataVenc = new Date(adi.venc);
                const diffDias = Math.ceil((dataVenc - hoje) / (1000 * 60 * 60 * 24));
                
                let classeStatus = "prazo-ok";
                let textoStatus = "No Prazo";

                if (diffDias < 0) {
                    classeStatus = "prazo-vencido";
                    textoStatus = "⚠️ VENCIDO";
                } else if (diffDias <= 7) {
                    classeStatus = "prazo-urgente";
                    textoStatus = "⏳ URGENTE";
                }

                tbodyAdi.innerHTML += `
                    <tr>
                        <td><b>${adi.responsavel}</b></td>
                        <td>${adi.nf}</td>
                        <td>${adi.fornecedor}</td>
                        <td>${new Date(adi.venc).toLocaleDateString('pt-BR', {timeZone: 'UTC'})}</td>
                        <td>R$ ${parseFloat(adi.valor).toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                        <td><span class="status-prazo ${classeStatus}">${textoStatus}</span></td>
                    </tr>`;
            });
        } else {
            tbodyAdi.innerHTML = "<tr><td colspan='6' style='text-align:center'>Nenhum adiantamento ativo.</td></tr>";
        }

        // 3. Preenche Tabela de Gestão (Somente Crislene verá)
        const tbodyG = document.querySelector("#tabelaGestao tbody");
        tbodyG.innerHTML = data.statsGestor.map(c => `
            <tr>
                <td><b>${c.nome}</b></td>
                <td>${c.total}</td>
                <td style="color:var(--accent); font-weight:800">${c.media}</td>
                <td>${c.pico}</td>
                <td>${c.atividades}</td>
            </tr>
        `).join('');

        // Finaliza o Loading
        document.getElementById('dash-loading').style.display = 'none';
        document.getElementById('dash-content').style.display = 'block';

    } catch (e) {
        console.error("Erro ao carregar Dashboard:", e);
        document.getElementById('dash-loading').innerHTML = "<p style='color:var(--danger)'>Erro ao sincronizar dados. Verifique a URL do Script.</p>";
    }
}

// --- NAVEGAÇÃO ENTRE ABAS ---
function switchTab(aba) {
    abaAtual = aba;
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + aba.toLowerCase()).classList.add('active');

    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        // Mostra detalhes de gestão apenas se for Gestora
        document.getElementById('dash-gestor').style.display = (usuarioAtual.role === 'gestor') ? 'block' : 'none';
        carregarEstatisticas();
    } else {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'block';
        document.getElementById('tabTitle').innerText = "Lote de Notas: " + aba;
        
        // Ajusta campos específicos
        document.getElementById('fieldLote').style.display = (aba === 'Adiantamento') ? 'none' : 'flex';
        document.getElementById('fieldNumAdi').style.display = (aba === 'Adiantamento') ? 'flex' : 'none';
        
        configurarStatusCard(aba);
        configurarTabelaHeader(aba);
        atualizarTabela();
    }
}

// --- BUSCA GLOBAL NO HISTÓRICO ---
async function buscarNoBanco() {
    const q = document.getElementById('inputBusca').value;
    if (q.length < 3) return alert("Digite pelo menos 3 caracteres para a busca.");

    const btn = document.querySelector('.btn-search');
    btn.innerText = "PESQUISANDO...";
    
    try {
        const res = await fetch(`${URL_SCRIPT}?search=${q}&tab=${abaAtual}`);
        const results = await res.json();
        const tbody = document.querySelector("#tabelaResultados tbody");
        tbody.innerHTML = "";

        if (results.length === 0) {
            tbody.innerHTML = "<tr><td colspan='6' style='text-align:center'>Nenhum registro encontrado no histórico.</td></tr>";
        } else {
            results.forEach(r => {
                const dataFormatada = r.data ? new Date(r.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'}) : "---";
                const valorFormatado = r.valor ? `R$ ${parseFloat(r.valor).toLocaleString('pt-BR', {minimumFractionDigits: 2})}` : "---";

                tbody.innerHTML += `
                    <tr>
                        <td><b>${r.responsavel}</b></td>
                        <td>${dataFormatada}</td>
                        <td><b>${r.nf}</b></td>
                        <td>${r.fornecedor}</td>
                        <td>${valorFormatado}</td>
                        <td><span style="font-size:10px; font-weight:800; color:var(--accent)">${r.status}</span></td>
                    </tr>`;
            });
        }
        document.getElementById('searchModal').style.display = 'block';
    } catch (e) {
        alert("Erro ao consultar o banco de dados.");
    } finally {
        btn.innerText = "PESQUISAR";
    }
}

// --- DINÂMICA DE FORMULÁRIOS ---
function configurarStatusCard(aba) {
    const sc = document.getElementById('fieldStatus');
    if (aba === 'Digitadas') {
        sc.innerHTML = `
            <div class="status-row">
                <span class="status-label">SISTEMA:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gSis" value="Oracle"> ORACLE</label>
                    <label class="radio-item"><input type="radio" name="gSis" value="MV"> MV</label>
                </div>
            </div>
            <div class="status-row" style="margin-top:15px">
                <span class="status-label">PROCESSO:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gPro" value="Manual"> MANUAL</label>
                    <label class="radio-item"><input type="radio" name="gPro" value="Reprocessada"> REPROCESSADA</label>
                </div>
            </div>`;
    } else if (aba === 'Recebimento') {
        sc.innerHTML = `
            <div class="status-row">
                <span class="status-label">LOGÍSTICA:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gLog" value="Encaminhado"> 🚚 ENVIADO</label>
                    <label class="radio-item"><input type="radio" name="gLog" value="Aguardando"> 📦 AGUARDANDO</label>
                </div>
            </div>`;
    } else {
        sc.innerHTML = `<p style="font-size:11px; font-weight:700; color:var(--text-muted)">MODO ADIANTAMENTO: FOCO EM VENCIMENTO</p>`;
    }
}

function configurarTabelaHeader(aba) {
    const head = document.getElementById('tableHeader');
    head.innerHTML = `<th>NF</th><th>FORNECEDOR</th><th>VENCIMENTO</th><th>VALOR</th><th>AÇÃO</th>`;
}

// --- LÓGICA DE LOTE ---
function adicionarNota() {
    const nf = document.getElementById('f_nf').value;
    if (!nf) return alert("O número da NF é obrigatório!");
    
    let status = "";
    if (abaAtual === 'Digitadas') {
        const s = document.querySelector('input[name="gSis"]:checked');
        const p = document.querySelector('input[name="gPro"]:checked');
        if (!s || !p) return alert("Selecione o Sistema e o Processo!");
        status = `${s.value} | ${p.value}`;
    } else if (abaAtual === 'Recebimento') {
        const l = document.querySelector('input[name="gLog"]:checked');
        if (!l) return alert("Selecione o status logístico!");
        status = l.value;
    }

    const nota = {
        destino: abaAtual,
        responsavel: usuarioAtual.nome,
        data: document.getElementById('f_data').value,
        nf: nf,
        fornecedor: document.getElementById('f_fornecedor').value,
        razaoSocial: document.getElementById('f_lote').value, // Reaproveitando para Razão no Lote
        vencimento: document.getElementById('f_vencimento').value,
        valor: document.getElementById('f_valor').value,
        setor: document.getElementById('f_setor').value || "GERAL",
        possuiLote: document.getElementById('f_lote').value,
        numAdiantamento: document.getElementById('f_num_adi').value,
        statusDigitacao: status,
        situacaoMaterial: status
    };

    listas[abaAtual].push(nota);
    atualizarTabela();
    
    // Limpa apenas o campo de NF e foca nele para o próximo lançamento
    document.getElementById('f_nf').value = "";
    document.getElementById('f_nf').focus();
}

function atualizarTabela() {
    const tbody = document.querySelector("#tabelaDados tbody");
    tbody.innerHTML = "";
    listas[abaAtual].forEach((n, i) => {
        tbody.innerHTML += `
            <tr>
                <td><b>${n.nf}</b></td>
                <td>${n.fornecedor}</td>
                <td>${formatarDataParaExibir(n.vencimento)}</td>
                <td>R$ ${parseFloat(n.valor).toLocaleString('pt-BR', {minimumFractionDigits: 2})}</td>
                <td>
                    <button onclick="removerNota(${i})" style="border:none; background:none; cursor:pointer; color:var(--danger)">
                        <i class="ph ph-trash" style="font-size:20px"></i>
                    </button>
                </td>
            </tr>`;
    });
    document.getElementById('areaAcoes').style.display = listas[abaAtual].length > 0 ? 'grid' : 'none';
}

function removerNota(i) {
    listas[abaAtual].splice(i, 1);
    atualizarTabela();
}

// --- SINCRONIZAÇÃO E PROTOCOLO ---
async function enviarTudo() {
    const btn = document.getElementById('btnEnviar');
    btn.disabled = true;
    btn.innerText = "SINCRONIZANDO...";

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify(listas[abaAtual])
        });
        
        alert("🚀 Lote enviado com sucesso para o Google Sheets!");
        listas[abaAtual] = [];
        atualizarTabela();
    } catch (e) {
        alert("Erro ao enviar dados. Verifique sua conexão.");
    } finally {
        btn.disabled = false;
        btn.innerText = "🚀 SINCRONIZAR";
    }
}

function copiarProtocolo() {
    let texto = `*PROTOCOLO DE ENTREGA - ${usuarioAtual.nome.toUpperCase()}*\n`;
    texto += `Data: ${new Date().toLocaleDateString()}\n\n`;
    
    listas[abaAtual].forEach(n => {
        texto += `NF: ${n.nf} | Fornecedor: ${n.fornecedor} | Valor: R$ ${n.valor}\n`;
    });

    navigator.clipboard.writeText(texto).then(() => {
        alert("Protocolo copiado para a área de transferência!");
    });
}

function fecharModal(id) { document.getElementById(id).style.display = 'none'; }