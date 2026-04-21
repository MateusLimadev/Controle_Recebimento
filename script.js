// --- CONFIGURAÇÃO ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbzrbc6xqFhpqRw2U9_1T4_rhscRJWTWlQPsCFH_5JM5Kedlq-DJj5IPpTkG3m9zcaHB2Q/exec";

let usuarioAtual = null;
let loginAtual = null;
let abaAtual = "Digitadas";
let listas = { "Digitadas": [], "Recebimento": [], "Adiantamento": [] };

// Garante que cada alerta só dispara uma vez por sessão (ao fazer login)
let alertaAdiJaExibido = false;
let alertaProjJaExibido = false;
let alertaJaExibido = false; // mantido por compatibilidade
let alertaProjPendente = false; // projeção carregou mas adi ainda estava aberto

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
    alertaAdiJaExibido = true;
    alertaJaExibido = true; // compatibilidade
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
    alertaJaExibido = false;
    alertaAdiJaExibido = false;
    alertaProjJaExibido = false;
    alertaProjPendente = false;
    usuarioAtual = { nome: data.nome, role: data.role, comprador: data.comprador === true || data.comprador === 'true' };
    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('primeiroAcessoScreen').style.display = 'none';
    document.getElementById('mainHeader').style.display = 'flex';
    document.getElementById('app').style.display = 'block';
    document.getElementById('userNameHeader').innerText = usuarioAtual.nome;

    const isAdmin    = usuarioAtual.role === 'administrador';
    const verProjecao = usuarioAtual.comprador || isAdmin;

    document.getElementById('btn-projecao').style.display      = verProjecao ? 'inline-flex' : 'none';
    document.getElementById('carrinhoHeaderBtn').style.display = verProjecao ? 'flex'        : 'none';
    document.getElementById('btn-admin').style.display         = isAdmin     ? 'inline-flex' : 'none';

    switchTab('Dashboard');
}

// --- EVENTOS DE TECLADO E MODAL ---
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('passInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('userInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('confirmaSenha').addEventListener('keydown', e => { if (e.key === 'Enter') confirmarNovaSenha(); });

    // Fecha modais clicando fora
    ['searchModal', 'modalAlertaAdi', 'modalAlertaProj', 'modalConfirmaSaida', 'modalUsuario', 'modalDeletarUsuario'].forEach(id => {
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

        if (data.adiantamentosSetor && data.adiantamentosSetor.length > 0) {
            const hoje = new Date();
            hoje.setHours(0, 0, 0, 0);

            let adiantamentosParaExibir = data.adiantamentosSetor;
            if (usuarioAtual.role === "digitador") {
                adiantamentosParaExibir = data.adiantamentosSetor.filter(adi => adi.responsavel === usuarioAtual.nome);
            }
            adiantamentosParaExibir.sort((a, b) => new Date(a.venc) - new Date(b.venc));

            // Salva no cache global
            adiantamentosCarregados = adiantamentosParaExibir;

            // ── ALERTA DE LOGIN: dispara apenas uma vez por sessão ──
            if (!alertaAdiJaExibido) {
                exibirAlertaAdiantamento(adiantamentosParaExibir);
            }
        } else {
            adiantamentosCarregados = [];
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

        // Carrega dados de projeção em background
        await carregarProjecaoBackground();
    } catch (e) {
        document.getElementById('dash-loading').innerHTML =
            "<p style='color:var(--danger)'>Erro ao carregar dados do Google Sheets.</p>";
    }
}

// --- RENDERIZAÇÃO DA TABELA DE ADIANTAMENTOS (aba Adiantamento) ---

function renderizarTabelaAdiantamentos(lista) {
    const tbody = document.querySelector("#tabelaMonitorAdi tbody");
    tbody.innerHTML = "";

    if (!lista || lista.length === 0) {
        tbody.innerHTML = "<tr><td colspan='7' style='text-align:center; padding:20px;'>Nenhum adiantamento pendente.</td></tr>";
        return;
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    lista.forEach(adi => {
        const diff = Math.ceil((new Date(adi.venc) - hoje) / (1000 * 60 * 60 * 24));
        let cls = "prazo-ok", txt = "No Prazo";
        if (diff < 0)       { cls = "prazo-vencido"; txt = "⚠️ VENCIDO"; }
        else if (diff <= 7) { cls = "prazo-urgente"; txt = "⏳ URGENTE"; }

        tbody.innerHTML += `
            <tr id="row-adi-${adi.nf}">
                <td><b>${adi.responsavel}</b></td>
                <td>${adi.nf}</td>
                <td>${adi.fornecedor}</td>
                <td>${new Date(adi.venc).toLocaleDateString('pt-BR', {timeZone: 'UTC'})}</td>
                <td>R$ ${formatarValor(adi.valor)}</td>
                <td><span class="status-prazo ${cls}">${txt}</span></td>
                <td style="text-align:center">
                    <button class="btn-saida-adi"
                        title="Registrar saída / quitar adiantamento"
                        onclick="deletarAdiantamento('${adi.nf}', '${adi.responsavel}')">
                        <i class="ph ph-door-open"></i>
                    </button>
                </td>
            </tr>`;
    });
}

async function carregarAdiantamentos() {
    const secao = document.getElementById('monitorAdiantamentosSection');
    secao.style.display = 'block';

    // Se já temos os dados em cache, renderiza direto
    if (adiantamentosCarregados.length > 0) {
        renderizarTabelaAdiantamentos(adiantamentosCarregados);
        return;
    }

    // Senão, busca da planilha
    document.getElementById('adi-loading').style.display = 'flex';
    try {
        const res  = await fetch(URL_SCRIPT);
        const data = await res.json();

        let lista = data.adiantamentosSetor || [];
        if (usuarioAtual.role === "digitador") {
            lista = lista.filter(a => a.responsavel === usuarioAtual.nome);
        }
        lista.sort((a, b) => new Date(a.venc) - new Date(b.venc));
        adiantamentosCarregados = lista;
        renderizarTabelaAdiantamentos(lista);
    } catch (e) {
        document.querySelector("#tabelaMonitorAdi tbody").innerHTML =
            "<tr><td colspan='7' style='text-align:center; color:var(--danger)'>Erro ao carregar dados.</td></tr>";
    } finally {
        document.getElementById('adi-loading').style.display = 'none';
    }
}
async function carregarProjecaoBackground() {
    try {
        const nomeUsuario = usuarioAtual.nome.split(' ')[0].toUpperCase(); // pega primeiro nome em maiúscula
        const res = await fetch(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`);
        const data = await res.json();

        if (!data.erro) {
            dadosProjecao = data.itens || [];
            
            // Exibe alerta de projeção apenas uma vez no login
            if (!alertaProjJaExibido && data.zeradosTotal > 0) {
                const modalAdiAberto = document.getElementById('modalAlertaAdi').style.display === 'flex';
                if (!modalAdiAberto) {
                    // Modal de adi já foi fechado (ou nunca abriu): dispara agora
                    exibirAlertaProjecao(dadosProjecao);
                } else {
                    // Modal de adi ainda está aberto: marca pendente para disparar ao fechar
                    alertaProjPendente = true;
                }
            }
        }
    } catch (e) {
        console.warn('Erro ao carregar projeção em background:', e);
    }
}

// --- VARIÁVEL GLOBAL PARA ARMAZENAR DADOS DE PROJEÇÃO ---
let dadosProjecao = [];
let filtroProjecaoAtual = 'todos';
let ordenacaoCobertura = 'asc';
let paginaAtualProjecao = 1;
const ITENS_POR_PAGINA = 100;
let itensFiltradosProjecao = [];

// Cache de adiantamentos (compartilhado entre dashboard e aba Adiantamento)
let adiantamentosCarregados = [];

// --- CARRINHO DE COMPRAS ---
let carrinho = []; // { codigo, descricao, cobertura, statusTexto }

// --- CARREGAMENTO E FILTRO DE PROJEÇÃO ---

async function carregarProjecao() {
    // Se dados já foram carregados em background, mostra logo
    if (dadosProjecao.length > 0) {
        filtroProjecaoAtual = 'todos';
        atualizarTabelaProjecao(dadosProjecao);
        document.getElementById('proj-loading').style.display = 'none';
        document.getElementById('proj-content').style.display = 'block';
        return;
    }

    document.getElementById('proj-loading').style.display = 'flex';
    document.getElementById('proj-content').style.display = 'none';

    try {
        const nomeUsuario = usuarioAtual.nome.split(' ')[0].toUpperCase();
        const res = await fetch(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`);
        const data = await res.json();

        if (data.erro) {
            document.getElementById('proj-loading').innerHTML = 
                `<p style='color:var(--danger)'>Erro: ${data.erro}</p>`;
            return;
        }

        dadosProjecao = data.itens || [];
        
        filtroProjecaoAtual = 'todos';
        atualizarTabelaProjecao(dadosProjecao);

        document.getElementById('proj-loading').style.display = 'none';
        document.getElementById('proj-content').style.display = 'block';
    } catch (e) {
        document.getElementById('proj-loading').innerHTML = 
            "<p style='color:var(--danger)'>Erro ao carregar dados de projeção.</p>";
    }
}

function ordenarPorCobertura() {
    // Alterna entre ascendente e descendente
    ordenacaoCobertura = ordenacaoCobertura === 'asc' ? 'desc' : 'asc';

    // Atualiza o ícone
    const icon = document.getElementById('iconOrdenacao');
    if (ordenacaoCobertura === 'asc') {
        icon.className = 'ph ph-sort-ascending';
    } else {
        icon.className = 'ph ph-sort-descending';
    }

    // Ordena os dados atuais e redesenha
    let itensParaExibir = dadosProjecao;

    // Aplica o filtro atual
    if (filtroProjecaoAtual === 'rp') {
        itensParaExibir = dadosProjecao.filter(i => i.temRP);
    } else if (filtroProjecaoAtual === 'cd') {
        itensParaExibir = dadosProjecao.filter(i => i.saldoCD > 0);
    } else if (filtroProjecaoAtual === 'empenho') {
        itensParaExibir = dadosProjecao.filter(i => i.temEmpenho);
    } else if (filtroProjecaoAtual === 'zerado') {
        itensParaExibir = dadosProjecao.filter(i => i.zeradoSemCobertura);
    }

    // Ordena por cobertura
    itensParaExibir.sort((a, b) => {
        if (ordenacaoCobertura === 'asc') {
            return a.cobertura - b.cobertura;
        } else {
            return b.cobertura - a.cobertura;
        }
    });

    atualizarTabelaProjecao(itensParaExibir);
}

function filtrarProjecao(tipo) {
    // Atualiza botão ativo
    document.querySelectorAll('.filtro-btn').forEach(b => b.classList.remove('filtro-ativo'));
    event.target.closest('.filtro-btn').classList.add('filtro-ativo');

    filtroProjecaoAtual = tipo;
    let filtrado = dadosProjecao;

    if (tipo === 'rp') {
        filtrado = dadosProjecao.filter(i => i.temRP);
    } else if (tipo === 'cd') {
        filtrado = dadosProjecao.filter(i => i.saldoCD > 0);
    } else if (tipo === 'empenho') {
        filtrado = dadosProjecao.filter(i => i.temEmpenho);
    } else if (tipo === 'zerado') {
        filtrado = dadosProjecao.filter(i => i.zeradoSemCobertura);
    }

    // Aplica ordenação atual
    filtrado.sort((a, b) => {
        if (ordenacaoCobertura === 'asc') {
            return a.cobertura - b.cobertura;
        } else {
            return b.cobertura - a.cobertura;
        }
    });

    atualizarTabelaProjecao(filtrado);
}

function atualizarTabelaProjecao(itens) {
    itensFiltradosProjecao = itens;
    paginaAtualProjecao = 1;
    renderizarPaginaProjecao();
}

function renderizarPaginaProjecao() {
    const tbody = document.querySelector('#tabelaProjecao tbody');
    tbody.innerHTML = '';

    const totalItens = itensFiltradosProjecao.length;
    const totalPaginas = Math.ceil(totalItens / ITENS_POR_PAGINA);

    if (totalItens === 0) {
        tbody.innerHTML = "<tr><td colspan='7' style='text-align:center; padding: 20px;'>Nenhum item encontrado para este filtro.</td></tr>";
        document.getElementById('projExibindo').innerText = 'Exibindo: 0 itens';
        document.getElementById('projPaginacao').innerHTML = '';
        return;
    }

    const inicio = (paginaAtualProjecao - 1) * ITENS_POR_PAGINA;
    const fim = Math.min(inicio + ITENS_POR_PAGINA, totalItens);
    const itensDaPagina = itensFiltradosProjecao.slice(inicio, fim);

    itensDaPagina.forEach(item => {
        let statusBadges = '';
        let temAlgumStatus = false;

        if (item.temRP) {
            statusBadges += '<span class="badge-status rp-sim"><i class="ph ph-check-circle"></i> RP</span> ';
            temAlgumStatus = true;
        }
        if (item.saldoCD > 0) {
            statusBadges += '<span class="badge-status cd-sim"><i class="ph ph-check-circle"></i> CD</span> ';
            temAlgumStatus = true;
        }
        if (item.temEmpenho) {
            statusBadges += '<span class="badge-status empenho-sim"><i class="ph ph-check-circle"></i> EMP</span> ';
            temAlgumStatus = true;
        }
        if (item.zeradoSemCobertura) {
            statusBadges += '<span class="badge-status zerado"><i class="ph ph-warning-diamond"></i> CRÍTICO</span>';
            temAlgumStatus = true;
        }

        if (!temAlgumStatus) {
            statusBadges = '<span class="badge-status" style="background: rgba(234, 88, 12, 0.15); color: #ea580c;"><i class="ph ph-shopping-cart"></i> COMPRAR</span>';
        }

        // Texto limpo do status para o carrinho/CSV
        let statusTexto = 'COMPRAR';
        if (item.zeradoSemCobertura)    statusTexto = 'CRÍTICO';
        else if (item.temRP)            statusTexto = 'RP';
        else if (item.saldoCD > 0)      statusTexto = 'CD';
        else if (item.temEmpenho)       statusTexto = 'EMPENHO';

        tbody.innerHTML += `
            <tr>
                <td><b>${item.codigo}</b></td>
                <td>${item.descricao}</td>
                <td>${item.cobertura || 0}</td>
                <td>${item.saldoCD || 0}</td>
                <td>${item.temRP ? '✓' : '✗'}</td>
                <td>${item.temEmpenho ? '✓' : '✗'}</td>
                <td>${statusBadges}</td>
                <td style="text-align:center;">
                    <button class="btn-add-carrinho ${carrinho.some(c => c.codigo === item.codigo) ? 'no-carrinho' : ''}"
                        data-codigo="${item.codigo}"
                        data-descricao="${item.descricao.replace(/"/g, '&quot;')}"
                        data-cobertura="${item.cobertura || 0}"
                        data-status="${statusTexto}"
                        title="${carrinho.some(c => c.codigo === item.codigo) ? 'Remover do carrinho' : 'Adicionar ao carrinho'}"
                        onclick="toggleCarrinhoItem(this)">
                        <i class="${carrinho.some(c => c.codigo === item.codigo) ? 'ph ph-check-circle' : 'ph ph-shopping-cart'}"></i>
                    </button>
                </td>
            </tr>`;
    });

    document.getElementById('projTotal').innerText = `Total: ${dadosProjecao.length} itens`;
    document.getElementById('projExibindo').innerText = `Exibindo: ${inicio + 1}–${fim} de ${totalItens} itens`;

    // Renderiza controles de paginação
    renderizarControlesPaginacao(totalPaginas);
}

function renderizarControlesPaginacao(totalPaginas) {
    const container = document.getElementById('projPaginacao');
    if (!container) return;

    if (totalPaginas <= 1) {
        container.innerHTML = '';
        return;
    }

    let html = '';

    // Botão anterior
    html += `<button class="pag-btn ${paginaAtualProjecao === 1 ? 'pag-disabled' : ''}" 
                onclick="irParaPagina(${paginaAtualProjecao - 1})" 
                ${paginaAtualProjecao === 1 ? 'disabled' : ''}>
                <i class="ph ph-caret-left"></i>
             </button>`;

    // Páginas
    for (let i = 1; i <= totalPaginas; i++) {
        // Exibe sempre: primeira, última, atual e as 2 vizinhas
        if (
            i === 1 || i === totalPaginas ||
            (i >= paginaAtualProjecao - 2 && i <= paginaAtualProjecao + 2)
        ) {
            html += `<button class="pag-btn ${i === paginaAtualProjecao ? 'pag-ativa' : ''}" 
                        onclick="irParaPagina(${i})">${i}</button>`;
        } else if (
            i === paginaAtualProjecao - 3 ||
            i === paginaAtualProjecao + 3
        ) {
            html += `<span class="pag-reticencias">…</span>`;
        }
    }

    // Botão próximo
    html += `<button class="pag-btn ${paginaAtualProjecao === totalPaginas ? 'pag-disabled' : ''}" 
                onclick="irParaPagina(${paginaAtualProjecao + 1})"
                ${paginaAtualProjecao === totalPaginas ? 'disabled' : ''}>
                <i class="ph ph-caret-right"></i>
             </button>`;

    container.innerHTML = html;
}

function irParaPagina(pagina) {
    const totalPaginas = Math.ceil(itensFiltradosProjecao.length / ITENS_POR_PAGINA);
    if (pagina < 1 || pagina > totalPaginas) return;
    paginaAtualProjecao = pagina;
    renderizarPaginaProjecao();
    // Scrolla suavemente para o topo da tabela
    document.querySelector('#tabelaProjecao').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// --- PAGINAÇÃO DO MODAL DE PROJEÇÃO CRÍTICA ---
let zeradosAlerta = [];
let paginaAlertaProj = 1;
const ITENS_PAG_ALERTA = 15;

function exibirAlertaProjecao(lista) {
    zeradosAlerta = lista.filter(i => i.zeradoSemCobertura);

    if (zeradosAlerta.length === 0) return;

    document.getElementById('numZerados').innerText = zeradosAlerta.length;
    paginaAlertaProj = 1;
    renderizarPaginaAlertaProj();

    document.getElementById('modalAlertaProj').style.display = 'flex';
    tocarSomMSN();
    alertaProjJaExibido = true;
}

function renderizarPaginaAlertaProj() {
    const total = zeradosAlerta.length;
    const totalPags = Math.ceil(total / ITENS_PAG_ALERTA);
    const inicio = (paginaAlertaProj - 1) * ITENS_PAG_ALERTA;
    const fim = Math.min(inicio + ITENS_PAG_ALERTA, total);

    const tbody = document.querySelector('#tabelaAlertaProj tbody');
    tbody.innerHTML = '';

    zeradosAlerta.slice(inicio, fim).forEach(item => {
        tbody.innerHTML += `
            <tr>
                <td><b>${item.codigo}</b></td>
                <td>${item.descricao}</td>
                <td>${item.cobertura || 0}</td>
                <td>${item.saldoCD || 0}</td>
            </tr>`;
    });

    // Controles de paginação dentro do modal
    let pagCtrl = document.getElementById('pagAlertaProj');
    if (!pagCtrl) {
        pagCtrl = document.createElement('div');
        pagCtrl.id = 'pagAlertaProj';
        pagCtrl.className = 'proj-paginacao';
        pagCtrl.style.marginTop = '14px';
        document.querySelector('#tabelaAlertaProj').parentElement.after(pagCtrl);
    }

    if (totalPags <= 1) {
        pagCtrl.innerHTML = '';
        return;
    }

    let html = `<button class="pag-btn ${paginaAlertaProj === 1 ? 'pag-disabled' : ''}" 
                    onclick="mudarPaginaAlerta(${paginaAlertaProj - 1})" 
                    ${paginaAlertaProj === 1 ? 'disabled' : ''}>
                    <i class="ph ph-caret-left"></i>
                </button>`;

    for (let i = 1; i <= totalPags; i++) {
        if (i === 1 || i === totalPags || (i >= paginaAlertaProj - 2 && i <= paginaAlertaProj + 2)) {
            html += `<button class="pag-btn ${i === paginaAlertaProj ? 'pag-ativa' : ''}" 
                        onclick="mudarPaginaAlerta(${i})">${i}</button>`;
        } else if (i === paginaAlertaProj - 3 || i === paginaAlertaProj + 3) {
            html += `<span class="pag-reticencias">…</span>`;
        }
    }

    html += `<button class="pag-btn ${paginaAlertaProj === totalPags ? 'pag-disabled' : ''}" 
                onclick="mudarPaginaAlerta(${paginaAlertaProj + 1})"
                ${paginaAlertaProj === totalPags ? 'disabled' : ''}>
                <i class="ph ph-caret-right"></i>
             </button>`;

    // Info de página
    html += `<span style="font-size:11px; font-weight:700; color:var(--text-muted); margin-left:8px;">
                Página ${paginaAlertaProj} de ${totalPags}
             </span>`;

    pagCtrl.innerHTML = html;
}

function mudarPaginaAlerta(pagina) {
    const totalPags = Math.ceil(zeradosAlerta.length / ITENS_PAG_ALERTA);
    if (pagina < 1 || pagina > totalPags) return;
    paginaAlertaProj = pagina;
    renderizarPaginaAlertaProj();
}

// --- NAVEGAÇÃO ---

function switchTab(aba) {
    abaAtual = aba;
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    const btnEl = document.getElementById('btn-' + aba.toLowerCase());
    if (btnEl) btnEl.classList.add('active');

    const monitorSec = document.getElementById('monitorAdiantamentosSection');
    if (monitorSec) monitorSec.style.display = 'none';

    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'none';
        const isGestorOuAdmin = usuarioAtual.role === 'gestor' || usuarioAtual.role === 'administrador';
        document.getElementById('dash-gestor').style.display = isGestorOuAdmin ? 'block' : 'none';
        carregarEstatisticas();
    } else if (aba === 'Projecao') {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'block';
        document.getElementById('view-admin').style.display = 'none';
        const nomeUsuario = usuarioAtual.nome.split(' ')[0];
        document.querySelector('#view-projecao .header-tab h2').innerText = `Projeção de Compras — ${nomeUsuario}`;
        carregarProjecao();
    } else if (aba === 'Admin') {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'block';
        carregarUsuarios();
    } else {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'block';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'none';
        document.getElementById('tabTitle').innerText = "Lote de Notas: " + aba;
        document.getElementById('fieldNumAdi').style.display = (aba === 'Adiantamento') ? 'flex' : 'none';
        document.getElementById('fieldLote').style.display   = (aba === 'Adiantamento') ? 'none' : 'flex';
        configurarStatusCard(aba);
        configurarTableHeader();
        atualizarTabela();
        if (aba === 'Adiantamento') carregarAdiantamentos();
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

// =========================================================
// CARRINHO DE COMPRAS
// =========================================================

function abrirCarrinho() {
    document.getElementById('carrinhoOverlay').style.display = 'block';
    document.getElementById('carrinhoDrawer').classList.add('aberto');
    renderizarCarrinho();
}

function fecharCarrinho() {
    document.getElementById('carrinhoOverlay').style.display = 'none';
    document.getElementById('carrinhoDrawer').classList.remove('aberto');
}

function toggleCarrinhoItem(btn) {
    const codigo     = btn.dataset.codigo;
    const descricao  = btn.dataset.descricao;
    const cobertura  = parseFloat(btn.dataset.cobertura) || 0;
    const statusTexto = btn.dataset.status;

    const idx = carrinho.findIndex(i => i.codigo === codigo);
    if (idx >= 0) {
        carrinho.splice(idx, 1);
    } else {
        carrinho.push({ codigo, descricao, cobertura, statusTexto });
    }
    atualizarContadorCarrinho();
    // Atualiza visual do botão na tabela sem re-renderizar tudo
    const noCarrinho = carrinho.some(i => i.codigo === codigo);
    btn.classList.toggle('no-carrinho', noCarrinho);
    btn.title = noCarrinho ? 'Remover do carrinho' : 'Adicionar ao carrinho';
    btn.querySelector('i').className = noCarrinho ? 'ph ph-check-circle' : 'ph ph-shopping-cart';
}

function atualizarContadorCarrinho() {
    const contador = document.getElementById('carrinhoContador');
    if (carrinho.length > 0) {
        contador.innerText = carrinho.length;
        contador.style.display = 'flex';
    } else {
        contador.style.display = 'none';
    }
}

function renderizarCarrinho() {
    const vazio   = document.getElementById('carrinhoVazio');
    const lista   = document.getElementById('carrinhoLista');
    const footer  = document.getElementById('carrinhoFooter');
    const qtdH    = document.getElementById('carrinhoQtdHeader');

    qtdH.innerText = `${carrinho.length} ${carrinho.length === 1 ? 'item' : 'itens'}`;

    if (carrinho.length === 0) {
        vazio.style.display  = 'flex';
        lista.style.display  = 'none';
        footer.style.display = 'none';
        return;
    }

    vazio.style.display  = 'none';
    lista.style.display  = 'block';
    footer.style.display = 'flex';

    const tbody = document.getElementById('carrinhoTbody');
    tbody.innerHTML = '';
    carrinho.forEach((item, idx) => {
        tbody.innerHTML += `
            <tr>
                <td><b>${item.codigo}</b></td>
                <td style="max-width:160px; font-size:11px;">${item.descricao}</td>
                <td style="text-align:center; font-weight:700;">${item.cobertura ?? '—'}</td>
                <td>${item.statusTexto ?? '—'}</td>
                <td>
                    <button class="carrinho-remover-btn" onclick="removerDoCarrinho(${idx})" title="Remover">
                        <i class="ph ph-trash"></i>
                    </button>
                </td>
            </tr>`;
    });
}

function removerDoCarrinho(idx) {
    const codigo = carrinho[idx].codigo;
    carrinho.splice(idx, 1);
    atualizarContadorCarrinho();
    renderizarCarrinho();
    // Atualiza visual do botão na tabela
    const btn = document.querySelector(`.btn-add-carrinho[data-codigo="${codigo}"]`);
    if (btn) {
        btn.classList.remove('no-carrinho');
        btn.title = 'Adicionar ao carrinho';
        btn.querySelector('i').className = 'ph ph-shopping-cart';
    }
}

function limparCarrinho() {
    if (!confirm('Deseja limpar toda a lista de compras?')) return;
    carrinho = [];
    atualizarContadorCarrinho();
    renderizarCarrinho();
    // Resetar todos os botões da tabela
    document.querySelectorAll('.btn-add-carrinho.no-carrinho').forEach(btn => {
        btn.classList.remove('no-carrinho');
        btn.title = 'Adicionar ao carrinho';
        btn.querySelector('i').className = 'ph ph-shopping-cart';
    });
}

async function enviarCarrinho() {
    if (carrinho.length === 0) return;
    const btn = document.querySelector('.carrinho-enviar-btn');
    const obs = document.getElementById('carrinhoObs').value.trim();

    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> ENVIANDO...';

    // Nome da aba: "Carrinho_PrimeiroNome"
    const primeiroNome = usuarioAtual.nome.split(' ')[0];
    const abaDestino   = `Carrinho_${primeiroNome}`;

    const payload = carrinho.map(item => ({
        codigo:      item.codigo,
        descricao:   item.descricao,
        cobertura:   item.cobertura,
        status:      item.statusTexto,
        responsavel: usuarioAtual.nome,
        observacao:  obs,
        data:        new Date().toLocaleDateString('pt-BR'),
        aba:         abaDestino
    }));

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify({ tipo: 'carrinho', itens: payload })
        });
        tocarSomMSN();
        alert(`✅ ${carrinho.length} item(s) enviado(s) para a planilha "${abaDestino}"!`);
        carrinho = [];
        atualizarContadorCarrinho();
        renderizarCarrinho();
        document.getElementById('carrinhoObs').value = '';
        document.querySelectorAll('.btn-add-carrinho.no-carrinho').forEach(btn => {
            btn.classList.remove('no-carrinho');
            btn.title = 'Adicionar ao carrinho';
            btn.querySelector('i').className = 'ph ph-shopping-cart';
        });
        fecharCarrinho();
    } catch (e) {
        alert('Erro ao enviar. Verifique a conexão com a planilha.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-paper-plane-tilt"></i> ENVIAR PARA PLANILHA';
    }
}

// =========================================================
// EXPORTAR CARRINHO COMO CSV
// =========================================================
function exportarCarrinhoCSV() {
    if (carrinho.length === 0) return;

    const hoje = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    const nomeArquivo = `lista_compras_${usuarioAtual.nome.split(' ')[0].toLowerCase()}_${hoje}.csv`;
    const obs = document.getElementById('carrinhoObs').value.trim();

    const cabecalho = ['CÓDIGO', 'DESCRIÇÃO', 'COBERTURA', 'STATUS', 'OBSERVAÇÃO'];
    const linhas = carrinho.map(item => [
        `"${item.codigo}"`,
        `"${item.descricao.replace(/"/g, '""')}"`,
        `"${item.cobertura ?? ''}"`,
        `"${item.statusTexto ?? ''}"`,
        `"${obs.replace(/"/g, '""')}"`
    ]);

    const csvContent = [cabecalho.join(';'), ...linhas.map(l => l.join(';'))].join('\n');

    // BOM para Excel abrir com acentos corretos
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href     = url;
    link.download = nomeArquivo;
    link.click();
    URL.revokeObjectURL(url);

    // Limpa o carrinho após exportar
    carrinho = [];
    atualizarContadorCarrinho();
    renderizarCarrinho();
    document.getElementById('carrinhoObs').value = '';
    document.querySelectorAll('.btn-add-carrinho.no-carrinho').forEach(btn => {
        btn.classList.remove('no-carrinho');
        btn.title = 'Adicionar ao carrinho';
        btn.querySelector('i').className = 'ph ph-shopping-cart';
    });
    fecharCarrinho();
    tocarSomMSN();
}

let _saidaNFPendente = null;
let _saidaResponsavelPendente = null;

function deletarAdiantamento(nf, responsavel) {
    // Guarda os dados para usar ao confirmar
    _saidaNFPendente = nf;
    _saidaResponsavelPendente = responsavel;

    // Preenche o modal com os dados da linha
    document.getElementById('saidaNF').innerText = nf;
    document.getElementById('saidaResponsavel').innerText = responsavel;

    // Abre o modal
    document.getElementById('modalConfirmaSaida').style.display = 'flex';
}

async function confirmarSaida() {
    const nf = _saidaNFPendente;
    const responsavel = _saidaResponsavelPendente;

    if (!nf || !responsavel) return;

    // Fecha modal imediatamente
    fecharModal('modalConfirmaSaida');
    _saidaNFPendente = null;
    _saidaResponsavelPendente = null;

    // Remove do cache global também
    adiantamentosCarregados = adiantamentosCarregados.filter(a => a.nf.toString() !== nf.toString());

    // Remove visualmente com animação
    const row = document.getElementById(`row-adi-${nf}`);
    if (row) {
        row.style.transition = 'opacity 0.35s ease, transform 0.35s ease';
        row.style.opacity = '0';
        row.style.transform = 'translateX(30px)';
        setTimeout(() => row.remove(), 380);
    }

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify({
                tipo: 'deletarAdiantamento',
                nf: nf,
                responsavel: responsavel
            })
        });
        tocarSomMSN();
    } catch (e) {
        console.warn('Erro ao deletar adiantamento:', e);
    }
}

// =========================================================
// BUSCA NA ABA PROJEÇÃO
// =========================================================

function buscarNaProjecao() {
    const termo = document.getElementById('inputBuscaProjecao').value.trim().toLowerCase();
    const btnLimpar = document.getElementById('btnLimparBuscaProj');
    btnLimpar.style.display = termo.length > 0 ? 'flex' : 'none';

    if (termo.length === 0) {
        // Volta ao filtro atual sem busca
        aplicarFiltroAtual();
        return;
    }

    const filtrado = dadosProjecao.filter(i =>
        i.codigo.toLowerCase().includes(termo) ||
        i.descricao.toLowerCase().includes(termo)
    );

    // Aplica o filtro de aba por cima da busca
    let resultado = filtrado;
    if (filtroProjecaoAtual === 'rp')      resultado = filtrado.filter(i => i.temRP);
    else if (filtroProjecaoAtual === 'cd') resultado = filtrado.filter(i => i.saldoCD > 0);
    else if (filtroProjecaoAtual === 'empenho') resultado = filtrado.filter(i => i.temEmpenho);
    else if (filtroProjecaoAtual === 'zerado')  resultado = filtrado.filter(i => i.zeradoSemCobertura);

    atualizarTabelaProjecao(resultado);
}

function limparBuscaProjecao() {
    document.getElementById('inputBuscaProjecao').value = '';
    document.getElementById('btnLimparBuscaProj').style.display = 'none';
    aplicarFiltroAtual();
}

function aplicarFiltroAtual() {
    let itens = dadosProjecao;
    if (filtroProjecaoAtual === 'rp')      itens = dadosProjecao.filter(i => i.temRP);
    else if (filtroProjecaoAtual === 'cd') itens = dadosProjecao.filter(i => i.saldoCD > 0);
    else if (filtroProjecaoAtual === 'empenho') itens = dadosProjecao.filter(i => i.temEmpenho);
    else if (filtroProjecaoAtual === 'zerado')  itens = dadosProjecao.filter(i => i.zeradoSemCobertura);

    itens.sort((a, b) => ordenacaoCobertura === 'asc' ? a.cobertura - b.cobertura : b.cobertura - a.cobertura);
    atualizarTabelaProjecao(itens);
}

function fecharModal(id) {
    document.getElementById(id).style.display = 'none';

    // Ao fechar o alerta de adiantamento, dispara o de projeção se necessário
    if (id === 'modalAlertaAdi' && !alertaProjJaExibido) {
        const zerados = dadosProjecao.filter(i => i.zeradoSemCobertura);
        if (zerados.length > 0) {
            setTimeout(() => exibirAlertaProjecao(dadosProjecao), 300);
        } else if (dadosProjecao.length === 0) {
            alertaProjPendente = true;
        }
    }
}

// =========================================================
// ADMINISTRAÇÃO DE USUÁRIOS
// =========================================================

let _usuarioEditando = null; // login do usuário sendo editado (null = criação)
let _loginParaDeletar = null;

const CARGO_LABEL = {
    digitador:      { txt: 'Digitador',     cls: 'badge-cargo digitador' },
    gestor:         { txt: 'Gestor',         cls: 'badge-cargo gestor' },
    administrador:  { txt: 'Administrador',  cls: 'badge-cargo administrador' }
};

async function carregarUsuarios() {
    document.getElementById('admin-loading').style.display = 'flex';
    document.querySelector('#tabelaUsuarios tbody').innerHTML = '';

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=getUsuarios&solicitante=${encodeURIComponent(loginAtual)}`);
        const data = await res.json();
        if (data.erro) throw new Error(data.erro);
        renderizarTabelaUsuarios(data);
    } catch (e) {
        document.querySelector('#tabelaUsuarios tbody').innerHTML =
            `<tr><td colspan="5" style="text-align:center;color:var(--danger)">Erro ao carregar usuários.</td></tr>`;
    } finally {
        document.getElementById('admin-loading').style.display = 'none';
    }
}

function renderizarTabelaUsuarios(lista) {
    const tbody = document.querySelector('#tabelaUsuarios tbody');
    tbody.innerHTML = '';

    if (!lista.length) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center">Nenhum usuário encontrado.</td></tr>';
        return;
    }

    lista.forEach(u => {
        const cargo  = CARGO_LABEL[u.role] || { txt: u.role, cls: 'badge-cargo digitador' };
        const isComprador = u.comprador === 'true' || u.comprador === true || u.role === 'administrador';
        const badgeComprador = isComprador
            ? '<span class="badge-comprador">🛒 Comprador</span>'
            : '<span style="color:var(--text-muted); font-size:11px;">—</span>';
        const status = u.primeiroAcesso === 'true' || u.primeiroAcesso === true
            ? '<span class="badge-primeiro-acesso">⏳ Aguardando 1º acesso</span>'
            : '<span class="badge-ativo">✓ Ativo</span>';

        const isSelf = u.login === loginAtual;

        tbody.innerHTML += `
            <tr>
                <td><b>${u.nome}</b></td>
                <td style="color:var(--text-muted); font-size:12px;">${u.login}</td>
                <td><span class="${cargo.cls}">${cargo.txt}</span></td>
                <td>${badgeComprador}</td>
                <td>${status}</td>
                <td style="text-align:center;">
                    <div class="admin-acoes">
                        <button class="btn-admin-acao editar" onclick="abrirModalUsuario('${u.login}','${u.nome.replace(/'/g,"\\'")}','${u.role}',${isComprador})" title="Editar">
                            <i class="ph ph-pencil-simple"></i>
                        </button>
                        <button class="btn-admin-acao resetar" onclick="resetarSenhaUsuario('${u.login}','${u.nome.replace(/'/g,"\\'")}')'" title="Resetar senha">
                            <i class="ph ph-key"></i>
                        </button>
                        <button class="btn-admin-acao deletar" onclick="abrirModalDeletar('${u.login}','${u.nome.replace(/'/g,"\\'")}')'" title="Remover"
                            ${isSelf ? 'disabled style="opacity:0.3;cursor:not-allowed"' : ''}>
                            <i class="ph ph-trash"></i>
                        </button>
                    </div>
                </td>
            </tr>`;
    });
}

function abrirModalUsuario(login, nome, role, comprador) {
    const editando = !!login;
    _usuarioEditando = editando ? login : null;

    document.getElementById('modalUsuarioTitulo').innerHTML = editando
        ? '<i class="ph ph-pencil-simple"></i> EDITAR USUÁRIO'
        : '<i class="ph ph-user-plus"></i> NOVO USUÁRIO';

    document.getElementById('u_nome').value      = nome  || '';
    document.getElementById('u_login').value     = login || '';
    document.getElementById('u_role').value      = role  || 'digitador';
    document.getElementById('u_comprador').checked = !!comprador;

    document.getElementById('u_login').disabled            = editando;
    document.getElementById('u_senha_group').style.display = editando ? 'none' : 'flex';
    document.getElementById('u_editando_info').style.display = editando ? 'block' : 'none';
    document.getElementById('u_senha').value = '';

    document.getElementById('modalUsuario').style.display = 'flex';
}

async function salvarUsuario() {
    const nome      = document.getElementById('u_nome').value.trim();
    const login     = document.getElementById('u_login').value.trim().toLowerCase();
    const role      = document.getElementById('u_role').value;
    const senha     = document.getElementById('u_senha').value.trim();
    const comprador = document.getElementById('u_comprador').checked;
    const editando  = !!_usuarioEditando;

    if (!nome || !login) return alert('⚠️ Preencha nome e login.');
    if (!editando && senha.length < 6) return alert('⚠️ A senha inicial deve ter pelo menos 6 caracteres.');

    const btn = document.getElementById('btnSalvarUsuario');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';

    try {
        const action = editando ? 'editarUsuario' : 'criarUsuario';
        let url = `${URL_SCRIPT}?action=${action}&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(login)}&nome=${encodeURIComponent(nome)}&role=${encodeURIComponent(role)}&comprador=${comprador}`;
        if (!editando) url += `&senha=${encodeURIComponent(senha)}`;

        const res  = await fetch(url);
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro || 'Erro desconhecido');

        fecharModal('modalUsuario');
        tocarSomMSN();
        await carregarUsuarios();
    } catch (e) {
        alert('Erro ao salvar: ' + e.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-check-circle"></i> SALVAR';
    }
}

async function resetarSenhaUsuario(login, nome) {
    const nova = prompt(`Nova senha para ${nome}:`);
    if (!nova || nova.length < 6) return alert('Senha deve ter pelo menos 6 caracteres.');

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=resetarSenha&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(login)}&novaSenha=${encodeURIComponent(nova)}`);
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        tocarSomMSN();
        alert(`✅ Senha de ${nome} redefinida. O usuário deverá trocar no próximo acesso.`);
        await carregarUsuarios();
    } catch (e) {
        alert('Erro ao resetar senha: ' + e.message);
    }
}

function abrirModalDeletar(login, nome) {
    _loginParaDeletar = login;
    document.getElementById('deletarUsuarioNome').innerText  = nome;
    document.getElementById('deletarUsuarioLogin').innerText = login;
    document.getElementById('modalDeletarUsuario').style.display = 'flex';
}

async function confirmarDeletarUsuario() {
    if (!_loginParaDeletar) return;
    fecharModal('modalDeletarUsuario');

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=deletarUsuario&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(_loginParaDeletar)}`);
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        tocarSomMSN();
        await carregarUsuarios();
    } catch (e) {
        alert('Erro ao remover usuário: ' + e.message);
    } finally {
        _loginParaDeletar = null;
    }
}