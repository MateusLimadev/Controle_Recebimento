// --- CONFIGURAÇÕES ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbzrbc6xqFhpqRw2U9_1T4_rhscRJWTWlQPsCFH_5JM5Kedlq-DJj5IPpTkG3m9zcaHB2Q/exec"; // URL GERADA NO APPS SCRIPT
const USUARIOS = [
    {login: "mateus", senha: "123", nome: "Mateus Lima", role: "digitador"},
    {login: "crislene", senha: "456", nome: "Crislene", role: "gestor"}
];

let usuarioAtual = null; let abaAtual = "Digitadas"; let listas = { "Digitadas": [], "Recebimento": [], "Adiantamento": [] };

// FUNÇÃO PARA CORRIGIR DATA (Manual para evitar erro de fuso horário)
function formatarDataEntrada(dataISO) {
    if (!dataISO) return "---";
    const partes = dataISO.split('-');
    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function toggleTheme() {
    const b = document.body; b.setAttribute('data-theme', b.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
}

function realizarLogin() {
    const u = document.getElementById('userInput').value.toLowerCase();
    const s = document.getElementById('passInput').value;
    const user = USUARIOS.find(x => x.login === u && x.senha === s);
    if (user) {
        usuarioAtual = user; document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('mainHeader').style.display = 'flex'; document.getElementById('app').style.display = 'block';
        document.getElementById('userDisplay').innerText = usuarioAtual.nome.toUpperCase();
        carregarEstatisticas(); switchTab('Digitadas');
    } else alert("Usuário ou senha inválidos.");
}

async function carregarEstatisticas() {
    try {
        const res = await fetch(URL_SCRIPT); const data = await res.json();
        // Carrega indicadores do Setor (baseados na Planilha de Indicadores)
        document.getElementById('setorMediaGeral').innerText = data.statsSetor.mediaGeral;
        document.getElementById('setorForn').innerText = data.statsSetor.topForn;

        const tbody = document.querySelector("#tabelaGestao tbody");
        tbody.innerHTML = data.statsGestor.map(c => `
            <tr>
                <td><b>${c.nome}</b></td>
                <td>${c.total}</td>
                <td style="color:var(--accent); font-weight:800">${c.media}</td>
                <td>${c.pico}</td>
                <td>${c.atividades}</td>
            </tr>
        `).join('');
    } catch (e) { console.log("Erro ao carregar Dashboard."); }
}

function switchTab(aba) {
    abaAtual = aba; document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + aba.toLowerCase()).classList.add('active');
    
    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('dash-gestor').style.display = (usuarioAtual.role === 'gestor') ? 'block' : 'none';
        carregarEstatisticas();
    } else {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'block';
        document.getElementById('tabTitle').innerText = "Lote de Notas " + aba;
        document.getElementById('fieldLote').style.display = (aba === 'Adiantamento') ? 'none' : 'flex';
        document.getElementById('fieldNumAdi').style.display = (aba === 'Adiantamento') ? 'flex' : 'none';
        configurarStatusCard(aba);
        document.getElementById('tableHeader').innerHTML = `<th>NF</th><th>Forn.</th><th>Razão</th><th>Venc.</th><th>Valor</th><th>Ação</th>`;
        atualizarTabela();
    }
}

async function buscarNoBanco() {
    const q = document.getElementById('inputBusca').value; if (q.length < 3) return alert("Mínimo 3 caracteres.");
    const btn = document.querySelector('.btn-search'); btn.innerText = "PESQUISANDO...";
    
    try {
        const res = await fetch(`${URL_SCRIPT}?search=${q}&tab=${abaAtual}`);
        const results = await res.json();
        const tbody = document.querySelector("#tabelaResultados tbody");
        tbody.innerHTML = "";

        results.forEach(r => {
            const dataF = r.data ? new Date(r.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'}) : "---";
            const valorF = r.valor ? `R$ ${parseFloat(r.valor).toLocaleString('pt-BR', {minimumFractionDigits: 2})}` : "---";
            tbody.innerHTML += `<tr><td><b>${r.responsavel}</b></td><td>${dataF}</td><td><b>${r.nf}</b></td><td>${r.fornecedor}</td><td>${valorF}</td><td><span style="font-size:10px; font-weight:800; color:var(--accent)">${r.status}</span></td></tr>`;
        });
        document.getElementById('searchModal').style.display = 'block';
    } catch (e) { alert("Erro na busca."); }
    finally { btn.innerText = "PESQUISAR"; }
}

function configurarStatusCard(aba) {
    const sc = document.getElementById('fieldStatus');
    if (aba === 'Digitadas') {
        sc.innerHTML = `<div class="status-row"><span class="status-label">Sistema:</span><div class="radio-group"><label class="radio-item"><input type="radio" name="gSis" value="Oracle"> ORACLE</label><label class="radio-item"><input type="radio" name="gSis" value="MV"> MV</label></div></div><div class="status-row" style="margin-top:10px"><span class="status-label">Entrada:</span><div class="radio-group"><label class="radio-item"><input type="radio" name="gPro" value="Manual"> MANUAL</label><label class="radio-item"><input type="radio" name="gPro" value="Reprocessada"> REPROCESSADA</label></div></div>`;
    } else if (aba === 'Recebimento') {
        sc.innerHTML = `<div class="status-row"><span class="status-label">Logística:</span><div class="radio-group"><label class="radio-item"><input type="radio" name="gLog" value="Encaminhado"> 🚚 ENVIADO</label><label class="radio-item"><input type="radio" name="gLog" value="Aguardando"> 📦 AGUARDA</label></div></div>`;
    } else { sc.innerHTML = `<p style="font-size:11px; font-weight:700">MODO ADIANTAMENTO</p>`; }
}

function adicionarNota() {
    const nf = document.getElementById('f_nf').value; if (!nf) return alert("NF obrigatória!");
    const nota = {
        destino: abaAtual, responsavel: usuarioAtual.nome, 
        data: document.getElementById('f_data').value,
        nf: nf, fornecedor: document.getElementById('f_fornecedor').value,
        razaoSocial: document.getElementById('f_razao').value,
        vencimento: document.getElementById('f_vencimento').value,
        valor: document.getElementById('f_valor').value,
        setor: document.getElementById('f_setor').value || "GERAL",
        possuiLote: document.getElementById('f_lote').value,
        numAdiantamento: document.getElementById('f_num_adi').value,
        statusDigitacao: abaAtual === 'Digitadas' ? (document.querySelector('input[name="gSis"]:checked')?.value + " | " + document.querySelector('input[name="gPro"]:checked')?.value) : "",
        situacaoMaterial: abaAtual === 'Recebimento' ? document.querySelector('input[name="gLog"]:checked')?.value : ""
    };
    listas[abaAtual].push(nota); atualizarTabela();
    document.getElementById('f_nf').value = ""; document.getElementById('f_nf').focus();
}

function atualizarTabela() {
    const tbody = document.querySelector("#tabelaDados tbody"); tbody.innerHTML = "";
    listas[abaAtual].forEach((n, i) => {
        tbody.innerHTML += `<tr><td><b>${n.nf}</b></td><td>${n.fornecedor}</td><td>${n.razaoSocial}</td><td>${formatarDataEntrada(n.vencimento)}</td><td>R$ ${n.valor}</td><td><button onclick="removerNota(${i})"><i class="ph ph-trash"></i></button></td></tr>`;
    });
    document.getElementById('areaAcoes').style.display = listas[abaAtual].length > 0 ? 'grid' : 'none';
}

async function enviarTudo() {
    const btn = document.querySelector('.btn-send'); btn.disabled = true; btn.innerText = "SINC...";
    try {
        await fetch(URL_SCRIPT, { method: 'POST', mode: 'no-cors', body: JSON.stringify(listas[abaAtual]) });
        alert("🚀 Lote sincronizado!"); listas[abaAtual] = []; atualizarTabela();
    } catch(e) { alert("Erro!"); }
    finally { btn.disabled = false; btn.innerText = "🚀 SINCRONIZAR"; }
}

function removerNota(i) { listas[abaAtual].splice(i, 1); atualizarTabela(); }
function fecharModal(id) { document.getElementById(id).style.display = 'none'; }
function logout() { location.reload(); }
function copiarProtocolo() {
    let t = `*PROTOCOLO - ${usuarioAtual.nome.toUpperCase()}*\n`;
    listas[abaAtual].forEach(n => t += `${n.nf} | ${n.fornecedor} | R$ ${n.valor}\n`);
    navigator.clipboard.writeText(t).then(() => alert("Copiado!"));
}