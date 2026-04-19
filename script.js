// --- CONFIGURAÇÃO ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbzrbc6xqFhpqRw2U9_1T4_rhscRJWTWlQPsCFH_5JM5Kedlq-DJj5IPpTkG3m9zcaHB2Q/exec";

let usuarioAtual = null;
let loginAtual = null;
let abaAtual = "Digitadas";
let listas = { "Digitadas": [], "Recebimento": [], "Adiantamento": [] };

// Garante que o alerta só dispara uma vez por sessão (ao fazer login)
let alertaJaExibido = false;

// --- UTILITÁRIOS ---

function parseValor(valor) {
    const num = parseFloat(String(valor ?? '0').replace(',', '.'));
    return isNaN(num) ? 0 : num;
}

function formatarValor(valor) {
    if (valor === null || valor === undefined || valor === "") return "---";
    const num = parseFloat(String(valor).replace(',', '.'));
    return isNaN(num) ? "---" : num.toLocaleString('pt-BR', {minimumFractionDigits: 2});
}

function formatarDataBR(iso) {
    if (!iso) return "---";
    const partes = iso.split('-');
    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function toggleTheme() {
    const b = document.body;
    const isDark = b.getAttribute('data-theme') === 'dark';
    b.setAttribute('data-theme', isDark ? 'light' : 'dark');
    document.getElementById('themeIcon').className = isDark ? 'ph ph-moon' : 'ph ph-sun';
}

function logout() { location.reload(); }

// --- BOTÃO DE REFRESH DINÂMICO ---
async function refreshData() {
    const icon = document.getElementById('refreshIcon');
    icon.classList.add('rotating');
    await carregarEstatisticas();
    setTimeout(() => icon.classList.remove('rotating'), 1000);
}

// --- MENSAGENS DE ERRO INLINE ---
function mostrarErro(elementId, mensagem) {
    const el = document.getElementById(elementId);
    if (!el) return;
    el.innerText = mensagem;
    el.style.display = 'block';
    setTimeout(() => el.style.display = 'none', 3500);
}

// =========================================================
// SOM TIPO MSN MESSENGER (Web Audio API)
// =========================================================
function tocarSomMSN() {
    try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();

        // Arpejo ascendente em três notas — feel de notificação dos anos 2000
        const notas = [
            { freq: 1046.5, t: 0.00, dur: 0.18 },   // C6
            { freq: 1318.5, t: 0.19, dur: 0.18 },   // E6
            { freq: 1760.0, t: 0.38, dur: 0.35 },   // A6 — nota longa de encerramento
        ];

        notas.forEach(({ freq, t, dur }) => {
            const osc  = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.connect(gain);
            gain.connect(ctx.destination);

            osc.type = 'sine';
            osc.frequency.value = freq;

            const now = ctx.currentTime + t;
            gain.gain.setValueAtTime(0, now);
            gain.gain.linearRampToValueAtTime(0.28, now + 0.008);       // ataque rápido
            gain.gain.exponentialRampToValueAtTime(0.001, now + dur);   // decay suave

            osc.start(now);
            osc.stop(now + dur + 0.05);
        });
    } catch (err) {
        console.warn('Som indisponível no navegador:', err);
    }
}

// =========================================================
// POPUP DE ALERTA DE ADIANTAMENTOS
// =========================================================
function exibirAlertaAdiantamento(lista) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    // Filtra apenas vencidos e urgentes (≤ 7 dias)
    const criticos = lista.filter(adi => {
        const diff = Math.ceil((new Date(adi.venc) - hoje) / (1000 * 60 * 60 * 24));
        return diff <= 7;
    }).sort((a, b) => new Date(a.venc) - new Date(b.venc));

    if (criticos.length === 0) return; // sem alertas, não exibe nada

    // Contadores
    const numVencidas = criticos.filter(a => {
        return Math.ceil((new Date(a.venc) - hoje) / (1000 * 60 * 60 * 24)) < 0;
    }).length;
    const numUrgentes = criticos.length - numVencidas;

    document.getElementById('numVencidas').innerText = numVencidas;
    document.getElementById('numUrgentes').innerText = numUrgentes;

    // Oculta contador zerado para não poluir
    document.getElementById('cntVencido').style.display = numVencidas > 0 ? 'flex' : 'none';
    document.getElementById('cntUrgente').style.display = numUrgentes > 0 ? 'flex' : 'none';

    // Preenche tabela
    const tbody = document.querySelector('#tabelaAlertaAdi tbody');
    tbody.innerHTML = '';
    criticos.forEach(adi => {
        const diff = Math.ceil((new Date(adi.venc) - hoje) / (1000 * 60 * 60 * 24));
        let cls, txt;
        if (diff < 0) {
            cls = 'prazo-vencido';
            txt = `⚠️ VENCIDA há ${Math.abs(diff)} dia(s)`;
        } else if (diff === 0) {
            cls = 'prazo-vencido';
            txt = '⚠️ VENCE HOJE';
        } else {
            cls = 'prazo-urgente';
            txt = `⏳ ${diff} dia(s)`;
        }
        tbody.innerHTML += `
            <tr>
                <td><b>${adi.nf}</b></td>
                <td>${adi.fornecedor}</td>
                <td>${new Date(adi.venc).toLocaleDateString('pt-BR', {timeZone: 'UTC'})}</td>
                <td>R$ ${formatarValor(adi.valor)}</td>
                <td><span class="status-prazo ${cls}">${txt}</span></td>
            </tr>`;
    });

    // Exibe o modal e toca o som
    document.getElementById('modalAlertaAdi').style.display = 'flex';
    tocarSomMSN();
    alertaJaExibido = true;
}

// --- SISTEMA DE LOGIN ---

async function realizarLogin() {
    const u = document.getElementById('userInput').value.toLowerCase().trim();
    const s = document.getElementById('passInput').value.trim();

    if (!u || !s) return mostrarErro('loginErro', '⚠️ Preencha usuário e senha.');

    const btn = document.querySelector('.btn-login-final');
    btn.disabled = true;
    btn.innerText = "VERIFICANDO...";

    try {
        const res = await fetch(`${URL_SCRIPT}?action=login&login=${encodeURIComponent(u)}&senha=${encodeURIComponent(s)}`);
        const data = await res.json();

        if (data.ok) {
            loginAtual = u;
            if (data.primeiroAcesso) {
                document.getElementById('loginScreen').style.display = 'none';
                document.getElementById('primeiroAcessoScreen').style.display = 'flex';
                document.getElementById('nomeBoasVindas').innerText = data.nome.split(' ')[0];
            } else {
                entrarNoSistema(data);
            }
        } else {
            mostrarErro('loginErro', '⚠️ Usuário ou senha incorretos.');
        }
    } catch (err) {
        mostrarErro('loginErro', '⚠️ Erro ao conectar com o servidor.');
    } finally {
        btn.disabled = false;
        btn.innerText = "ENTRAR NO SISTEMA";
    }
}

// --- TROCA DE SENHA (PRIMEIRO ACESSO) ---

async function confirmarNovaSenha() {
    const nova     = document.getElementById('novaSenha').value.trim();
    const confirma = document.getElementById('confirmaSenha').value.trim();

    if (!nova || nova.length < 6) return mostrarErro('trocaErro', '⚠️ A senha deve ter pelo menos 6 caracteres.');
    if (nova !== confirma)        return mostrarErro('trocaErro', '⚠️ As senhas não coincidem.');

    const btn = document.getElementById('btnConfirmarSenha');
    btn.disabled = true;
    btn.innerText = "SALVANDO...";

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=trocarSenha&login=${encodeURIComponent(loginAtual)}&novaSenha=${encodeURIComponent(nova)}`);
        const data = await res.json();

        if (data.ok) {
            document.getElementById('primeiroAcessoScreen').style.display = 'none';
            entrarNoSistema(data);
        } else {
            mostrarErro('trocaErro', '⚠️ Erro ao salvar senha. Tente novamente.');
        }
    } catch (err) {
        mostrarErro('trocaErro', '⚠️ Erro ao conectar com o servidor.');
    } finally {
        btn.disabled = false;
        btn.innerText = "SALVAR E ENTRAR";
    }
}

function entrarNoSistema(data) {
    alertaJaExibido = false; // reseta o flag ao entrar
    usuarioAtual = { nome: data.nome, role: data.role };
    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('primeiroAcessoScreen').style.display = 'none';
    document.getElementById('mainHeader').style.display = 'flex';
    document.getElementById('app').style.display = 'block';
    document.getElementById('userNameHeader').innerText = usuarioAtual.nome;
    switchTab('Dashboard');
}

// --- EVENTOS DE TECLADO E MODAL ---
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('passInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('userInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('confirmaSenha').addEventListener('keydown', e => { if (e.key === 'Enter') confirmarNovaSenha(); });

    // Fecha modais clicando fora
    ['searchModal', 'modalAlertaAdi'].forEach(id => {
        document.getElementById(id).addEventListener('click', function(e) {
            if (e.target === this) fecharModal(id);
        });
    });
});

// --- DASHBOARD ---

async function carregarEstatisticas() {
    document.getElementById('dash-loading').style.display = 'flex';
    document.getElementById('dash-content').style.display = 'none';

    try {
        const res  = await fetch(URL_SCRIPT);
        const data = await res.json();

        document.getElementById('setorMediaGeral').innerText = data.statsSetor.mediaGeral;
        document.getElementById('setorForn').innerText       = data.statsSetor.topForn;

        const tbodyAdi = document.querySelector("#tabelaMonitorAdi tbody");
        tbodyAdi.innerHTML = "";

        if (data.adiantamentosSetor && data.adiantamentosSetor.length > 0) {
            const hoje = new Date();
            hoje.setHours(0, 0, 0, 0);

            let adiantamentosParaExibir = data.adiantamentosSetor;
            if (usuarioAtual.role === "digitador") {
                adiantamentosParaExibir = data.adiantamentosSetor.filter(adi => adi.responsavel === usuarioAtual.nome);
            }

            adiantamentosParaExibir.sort((a, b) => new Date(a.venc) - new Date(b.venc));

            if (adiantamentosParaExibir.length === 0) {
                tbodyAdi.innerHTML = "<tr><td colspan='6' style='text-align:center'>Nenhum adiantamento pendente.</td></tr>";
            } else {
                adiantamentosParaExibir.forEach(adi => {
                    const venc = new Date(adi.venc);
                    const diff = Math.ceil((venc - hoje) / (1000 * 60 * 60 * 24));
                    let cls = "prazo-ok", txt = "No Prazo";
                    if (diff < 0)      { cls = "prazo-vencido"; txt = "⚠️ VENCIDO"; }
                    else if (diff <= 7) { cls = "prazo-urgente"; txt = "⏳ URGENTE"; }
                    tbodyAdi.innerHTML += `
                        <tr>
                            <td><b>${adi.responsavel}</b></td>
                            <td>${adi.nf}</td>
                            <td>${adi.fornecedor}</td>
                            <td>${new Date(adi.venc).toLocaleDateString('pt-BR', {timeZone: 'UTC'})}</td>
                            <td>R$ ${formatarValor(adi.valor)}</td>
                            <td><span class="status-prazo ${cls}">${txt}</span></td>
                        </tr>`;
                });
            }

            // ── ALERTA DE LOGIN: dispara apenas uma vez por sessão ──
            if (!alertaJaExibido) {
                exibirAlertaAdiantamento(adiantamentosParaExibir);
            }

        } else {
            document.querySelector("#tabelaMonitorAdi tbody").innerHTML =
                "<tr><td colspan='6' style='text-align:center'>Nenhum adiantamento pendente.</td></tr>";
        }

        if (usuarioAtual.role === 'gestor') {
            const tbodyG = document.querySelector("#tabelaGestao tbody");
            tbodyG.innerHTML = data.statsGestor.map(c => `
                <tr>
                    <td><b>${c.nome}</b></td>
                    <td>${c.total}</td>
                    <td style="color:var(--accent); font-weight:800">${c.media}</td>
                    <td>${c.pico}</td>
                    <td>${c.atividades}</td>
                </tr>`).join('');
        }

        document.getElementById('dash-loading').style.display = 'none';
        document.getElementById('dash-content').style.display = 'block';
    } catch (e) {
        document.getElementById('dash-loading').innerHTML =
            "<p style='color:var(--danger)'>Erro ao carregar dados do Google Sheets.</p>";
    }
}

// --- NAVEGAÇÃO ---

function switchTab(aba) {
    abaAtual = aba;
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + aba.toLowerCase()).classList.add('active');

    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('dash-gestor').style.display = (usuarioAtual.role === 'gestor') ? 'block' : 'none';
        carregarEstatisticas();
    } else {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'block';
        document.getElementById('tabTitle').innerText = "Lote de Notas: " + aba;
        document.getElementById('fieldNumAdi').style.display = (aba === 'Adiantamento') ? 'flex' : 'none';
        document.getElementById('fieldLote').style.display   = (aba === 'Adiantamento') ? 'none' : 'flex';
        configurarStatusCard(aba);
        configurarTableHeader();
        atualizarTabela();
    }
}

// --- LÓGICA DE FORMULÁRIOS ---

function configurarStatusCard(aba) {
    const sc = document.getElementById('fieldStatus');
    if (aba === 'Digitadas') {
        sc.innerHTML = `
            <div class="status-row">
                <span class="status-label">Sistema:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gSis" value="Oracle"> ORACLE</label>
                    <label class="radio-item"><input type="radio" name="gSis" value="MV"> MV</label>
                </div>
            </div>
            <div class="status-row">
                <span class="status-label">Processo:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gPro" value="Manual"> MANUAL</label>
                    <label class="radio-item"><input type="radio" name="gPro" value="Reprocessada"> REPROCESSADA</label>
                </div>
            </div>`;
    } else if (aba === 'Recebimento') {
        sc.innerHTML = `
            <div class="status-row">
                <span class="status-label">Logística:</span>
                <div class="radio-group">
                    <label class="radio-item"><input type="radio" name="gLog" value="Encaminhado"> 🚚 ENVIADO</label>
                    <label class="radio-item"><input type="radio" name="gLog" value="Aguardando"> 📦 AGUARDA</label>
                </div>
            </div>`;
    } else {
        sc.innerHTML = `<p style="font-size:11px; font-weight:700; color:var(--text-muted)">MODO ADIANTAMENTO ATIVO</p>`;
    }
}

function configurarTableHeader() {
    const head = document.getElementById('tableHeader');
    head.innerHTML = `<th>NF</th><th>FORNECEDOR</th><th>RAZÃO</th><th>VENC.</th><th>VALOR</th><th>AÇÃO</th>`;
}

function adicionarNota() {
    const nf = document.getElementById('f_nf').value;
    if (!nf) return alert("Número da NF é obrigatório!");

    const nota = {
        destino:         abaAtual,
        responsavel:     usuarioAtual.nome,
        data:            document.getElementById('f_data').value,
        nf:              nf,
        fornecedor:      document.getElementById('f_fornecedor').value,
        razaoSocial:     document.getElementById('f_razao').value,
        vencimento:      document.getElementById('f_vencimento').value,
        valor:           document.getElementById('f_valor').value,
        setor:           document.getElementById('f_setor').value || "GERAL",
        possuiLote:      document.getElementById('f_lote').value,
        numAdiantamento: document.getElementById('f_num_adi').value,
        statusDigitacao: abaAtual === 'Digitadas'
            ? (document.querySelector('input[name="gSis"]:checked')?.value + " | " + document.querySelector('input[name="gPro"]:checked')?.value)
            : ""
    };

    listas[abaAtual].push(nota);
    atualizarTabela();
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
                <td>${n.razaoSocial}</td>
                <td>${formatarDataBR(n.vencimento)}</td>
                <td>R$ ${formatarValor(n.valor)}</td>
                <td><button onclick="removerNota(${i})" style="border:none;background:none;cursor:pointer;color:var(--danger)"><i class="ph ph-trash" style="font-size:20px"></i></button></td>
            </tr>`;
    });
    document.getElementById('areaAcoes').style.display = listas[abaAtual].length > 0 ? 'grid' : 'none';
}

function removerNota(i) {
    listas[abaAtual].splice(i, 1);
    atualizarTabela();
}

// --- SINCRONIZAÇÃO ---

async function enviarTudo() {
    const btn = document.getElementById('btnEnviar');
    btn.disabled = true;
    btn.innerText = "SINCRONIZANDO...";
    try {
        await fetch(URL_SCRIPT, { method: 'POST', mode: 'no-cors', body: JSON.stringify(listas[abaAtual]) });
        alert("🚀 Lote enviado com sucesso!");
        listas[abaAtual] = [];
        atualizarTabela();
    } catch (e) {
        alert("Erro ao conectar com o Google Sheets.");
    } finally {
        btn.disabled = false;
        btn.innerText = "🚀 SINCRONIZAR COM PLANILHA";
    }
}

function copiarProtocolo() {
    let t = `*PROTOCOLO - ${usuarioAtual.nome.toUpperCase()}*\n`;
    listas[abaAtual].forEach(n => t += `NF: ${n.nf} | ${n.fornecedor} | R$ ${formatarValor(n.valor)}\n`);
    navigator.clipboard.writeText(t).then(() => alert("Copiado para a área de transferência!"));
}

// --- BUSCA GLOBAL ---

async function buscarNoBanco() {
    const q   = document.getElementById('inputBusca').value;
    const btn = document.querySelector('.btn-search-global');
    if (q.length < 3) return alert("Digite ao menos 3 caracteres.");

    const originalContent = btn.innerHTML;
    btn.classList.add('loading');
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> BUSCANDO...';

    try {
        const res     = await fetch(`${URL_SCRIPT}?search=${encodeURIComponent(q)}&tab=${abaAtual}`);
        const results = await res.json();
        const tbody   = document.querySelector("#tabelaResultados tbody");
        tbody.innerHTML = "";

        if (results.length === 0) {
            tbody.innerHTML = "<tr><td colspan='6' style='text-align:center'>Nenhum registro encontrado.</td></tr>";
        } else {
            results.forEach(r => {
                const dF = r.data ? new Date(r.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'}) : "---";
                tbody.innerHTML += `
                    <tr>
                        <td><b>${r.responsavel}</b></td>
                        <td>${dF}</td>
                        <td><b>${r.nf}</b></td>
                        <td>${r.fornecedor}</td>
                        <td>R$ ${formatarValor(r.valor)}</td>
                        <td>${r.status || 'FINALIZADO'}</td>
                    </tr>`;
            });
        }
        document.getElementById('searchModal').style.display = 'flex';
    } catch (e) {
        alert("Erro na busca global. Verifique a conexão com a planilha.");
    } finally {
        btn.classList.remove('loading');
        btn.innerHTML = originalContent;
    }
}

function fecharModal(id) { document.getElementById(id).style.display = 'none'; }