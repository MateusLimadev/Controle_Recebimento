// --- CONFIGURAÇÃO ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbypMUlQb26vVFPwVemydCzcG-0nWkjC9EQvTEWGxpqPuIXnyjz0D427ZAXldtjCLjnFoA/exec";

let usuarioAtual = null;
let loginAtual   = null;
let sessaoAtual  = null;

// Logins que têm acesso a múltiplos módulos e precisam escolher ao entrar
const LOGINS_ACESSO_MULTIPLO = ['supmateus'];
let _dadosLoginPendente = null;
let modoAtual = 'suprimentos'; // 'suprimentos' | 'opme'

function entrarComoSuprimentos() {
    modoAtual = 'suprimentos';
    document.getElementById('seletorInterface').style.display = 'none';
    document.getElementById('btnEntrar').style.display = 'block';
    if (_dadosLoginPendente) {
        entrarNoSistema(_dadosLoginPendente);
        registrarLog('LOGIN', 'Acesso — Módulo Suprimentos');
        _dadosLoginPendente = null;
    }
}

function entrarComoOpme() {
    modoAtual = 'opme';
    document.getElementById('seletorInterface').style.display = 'none';
    document.getElementById('btnEntrar').style.display = 'block';
    if (_dadosLoginPendente) {
        _dadosLoginPendente._modoOpme = true;
        entrarNoSistema(_dadosLoginPendente);
        registrarLog('LOGIN', 'Acesso — Módulo OPME');
        _dadosLoginPendente = null;
    }
}

// Flags de cache — evita recarregar ao trocar de aba
// São zeradas apenas pelo botão ATUALIZAR ou F5
const _cache = {
    dashboard:    false,
    adiantamento: false,
    historico:    false,
    usuarios:     false,
};
let abaAtual = "Digitadas";
let listas = { "Digitadas": [], "Recebimento": [], "Adiantamento": [] };

// Adiciona token de sessão a qualquer URL do script
function addAuth(url) {
    if (!sessaoAtual?.token) return url;
    const sep = url.includes('?') ? '&' : '?';
    return `${url}${sep}tok=${encodeURIComponent(sessaoAtual.token)}&al=${encodeURIComponent(loginAtual || '')}`;
}

// Hash SHA-256 client-side (senha nunca viaja em texto puro)
async function hashSenha(senha) {
    const msgBuffer = new TextEncoder().encode(senha);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
    const hashArray  = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

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
    const novoTema = isDark ? 'light' : 'dark';
    b.setAttribute('data-theme', novoTema);
    const newIcon = isDark ? 'ph ph-moon' : 'ph ph-sun';
    ['themeIcon','themeIconLogin','themeIconMobile'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.className = newIcon;
    });
    if (loginAtual) localStorage.setItem('tema_' + loginAtual, novoTema);
}

function logout() { pararPolling(); location.reload(); }

// ── MOBILE NAV DRAWER ──────────────────────────────────────────
function toggleMobileNav() {
    const open = document.getElementById('mobileNavDrawer').classList.contains('open');
    open ? fecharMobileNav() : abrirMobileNav();
}

function abrirMobileNav() {
    document.getElementById('mobileNavDrawer').classList.add('open');
    document.getElementById('mobileNavOverlay').classList.add('open');
    document.getElementById('hamburgerBtn').classList.add('open');
    document.body.style.overflow = 'hidden';
    _atualizarMobileNavLinks();
}

function fecharMobileNav() {
    document.getElementById('mobileNavDrawer').classList.remove('open');
    document.getElementById('mobileNavOverlay').classList.remove('open');
    document.getElementById('hamburgerBtn').classList.remove('open');
    document.body.style.overflow = '';
}

function _atualizarMobileNavLinks() {
    // Sincroniza nome/role
    const nomeEl = document.getElementById('mobileNavNome');
    const roleEl = document.getElementById('mobileNavRole');
    if (nomeEl && usuarioAtual) nomeEl.textContent = usuarioAtual.nome || 'USUÁRIO';
    if (roleEl  && usuarioAtual) roleEl.textContent  = (usuarioAtual.role || '').toUpperCase();

    // Copia os botões visíveis do float-nav para o drawer
    const container = document.getElementById('mobileNavLinks');
    if (!container) return;
    container.innerHTML = '';

    const navMap = [
        { id: 'fn-dashboard',       tab: 'Dashboard',      label: 'Dashboard',        icon: 'ph-chart-line' },
        { id: 'fn-digitadas',       tab: 'Digitadas',      label: 'Digitadas',        icon: 'ph-article' },
        { id: 'fn-recebimento',     tab: 'Recebimento',    label: 'Recebidas',        icon: 'ph-truck' },
        { id: 'fn-adiantamento',    tab: 'Adiantamento',   label: 'Adiantamento',     icon: 'ph-clock-counter-clockwise' },
        { id: 'fn-projecao',        tab: 'Projecao',       label: 'Projeção',         icon: 'ph-package' },
        { id: 'fn-protocolos-opme', tab: 'ProtocolosOpme', label: 'Protocolar Notas', icon: 'ph-clipboard-text' },
        { id: 'fn-protocolos-sup',  tab: 'ProtocolosSup',  label: 'Protocolos OPME',  icon: 'ph-clipboard-text' },
        { id: 'fn-admin',           tab: 'Admin',          label: 'Administração',    icon: 'ph-shield-check', admin: true },
        { id: 'fn-projecao-opme', tab: 'ProjecaoOPME',   label: 'Projeção OPME',    icon: 'ph-chart-bar' },
        { id: 'fn-admin-opme',      tab: 'AdminOpme',      label: 'Admin OPME',       icon: 'ph-shield-check', admin: true },
    ];

    navMap.forEach(item => {
        const el = document.getElementById(item.id);
        if (!el || el.style.display === 'none') return;
        const isActive = el.classList.contains('active');
        const btn = document.createElement('button');
        btn.className = `mobile-nav-link${isActive ? ' active' : ''}${item.admin ? ' admin' : ''}`;
        btn.innerHTML = `<i class="ph ${item.icon}"></i><span>${item.label}</span>`;
        btn.onclick = () => { switchTab(item.tab); fecharMobileNav(); };
        container.appendChild(btn);
    });
}

// --- BOTÃO DE REFRESH DINÂMICO ---
async function refreshData() {
    const icon = document.getElementById('refreshIcon');
    icon.classList.add('rotating');

    // Zera todas as flags de cache — força novo fetch em tudo
    Object.keys(_cache).forEach(k => _cache[k] = false);

    try {
        await carregarBlacklistCache();
        await carregarEstatisticas();
        _cache.dashboard = true;

        if (abaAtual === 'Projecao' || abaAtual.startsWith('Projecao_')) {
            dadosProjecao = [];
            compradorCarregado = '';
            await carregarProjecao(abaAtual.startsWith('Projecao_') ? abaAtual.replace('Projecao_', '') : null);

        } else if (abaAtual === 'Adiantamento') {
            adiantamentosCarregados = [];
            await carregarAdiantamentos();
            _cache.adiantamento = true;

        } else if (abaAtual === 'Admin') {
            await carregarUsuarios();

        } else if (abaAtual === 'Digitadas' || abaAtual === 'Recebimento') {
            _histNotas = [];
            await carregarHistoricoNotas();
            _cache.historico = true;

        } else if (abaAtual === 'Admin') {
            await carregarUsuarios();
            _cache.usuarios = true;
        }

    } finally {
        setTimeout(() => icon.classList.remove('rotating'), 600);
    }
}

// =========================================================
// TOAST
// =========================================================
function mostrarToast(mensagem, tipo = 'info', duracao = 4000) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    const icons = { success: 'ph-check-circle', error: 'ph-x-circle', warning: 'ph-warning', info: 'ph-info' };
    toast.className = `toast toast-${tipo}`;
    toast.innerHTML = `<i class="ph ${icons[tipo] || 'ph-info'}"></i><span>${mensagem}</span>`;
    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('visivel'));
    setTimeout(() => {
        toast.classList.remove('visivel');
        setTimeout(() => toast.remove(), 400);
    }, duracao);
}

// =========================================================
// PESQUISA POR PERÍODO
// =========================================================
async function buscarPorPeriodo() {
    const inicio = document.getElementById('periodoInicio').value;
    const fim    = document.getElementById('periodoFim').value;

    if (!inicio || !fim) { mostrarToast('Selecione a data de início e fim.', 'warning'); return; }
    if (inicio > fim)    { mostrarToast('A data de início deve ser anterior à data fim.', 'warning'); return; }

    const btnBuscar = document.querySelector('.btn-periodo-buscar');
    btnBuscar.innerHTML = '<i class="ph ph-circle-notch rotating"></i> BUSCANDO...';
    btnBuscar.disabled  = true;

    const primeiroNome = usuarioAtual.nome.split(' ')[0];
    const nomeCompleto = encodeURIComponent(usuarioAtual.nome);

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=buscarPorPeriodo&aba=${encodeURIComponent(abaAtual)}&primeiroNome=${encodeURIComponent(primeiroNome)}&nomeCompleto=${nomeCompleto}&inicio=${encodeURIComponent(inicio)}&fim=${encodeURIComponent(fim)}`));
        const data = await res.json();

        const tbody  = document.querySelector('#tabelaResultadoPeriodo tbody');
        const result = document.getElementById('resultadoPeriodo');
        const label  = document.getElementById('labelResultadoPeriodo');

        tbody.innerHTML = '';

        if (!data.length) {
            tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--text-muted);">Nenhum registro encontrado nesse período.</td></tr>`;
        } else {
            data.forEach(r => {
                // Datas já vêm formatadas do servidor (DD/MM/YYYY)
                tbody.innerHTML += `
                    <tr>
                        <td><b>${r.responsavel || '—'}</b></td>
                        <td>${r.data || '—'}</td>
                        <td><b>${r.nf || '—'}</b></td>
                        <td>${r.fornecedor || '—'}</td>
                        <td style="font-size:11px;">${r.razaoSocial || '—'}</td>
                        <td>${r.vencimento || '—'}</td>
                        <td>R$ ${formatarValor(r.valor)}</td>
                        <td style="font-size:11px;">${r.setor || '—'}</td>
                    </tr>`;
            });
        }

        const ini  = new Date(inicio + 'T00:00:00').toLocaleDateString('pt-BR');
        const fim2 = new Date(fim    + 'T00:00:00').toLocaleDateString('pt-BR');
        label.textContent = `${data.length} registro(s) encontrado(s) entre ${ini} e ${fim2}`;
        result.style.display = 'block';
        document.getElementById('btnLimparPeriodo').style.display = 'inline-flex';
        mostrarToast(`${data.length} registros encontrados.`, data.length ? 'success' : 'info');
    } catch (e) {
        mostrarToast('Erro ao buscar registros. Tente novamente.', 'error');
    } finally {
        btnBuscar.innerHTML = '<i class="ph ph-magnifying-glass"></i> PESQUISAR';
        btnBuscar.disabled  = false;
    }
}

function limparPesquisaPeriodo() {
    document.getElementById('periodoInicio').value = '';
    document.getElementById('periodoFim').value    = '';
    document.getElementById('resultadoPeriodo').style.display  = 'none';
    document.getElementById('btnLimparPeriodo').style.display  = 'none';
    document.querySelector('#tabelaResultadoPeriodo tbody').innerHTML = '';
}
let modoFullscreen = false;
function toggleFullscreen() {
    modoFullscreen = !modoFullscreen;
    document.body.classList.toggle('fullscreen-mode', modoFullscreen);
    document.getElementById('fullscreenIcon').className = modoFullscreen ? 'ph ph-arrows-in' : 'ph ph-arrows-out';
    document.getElementById('btnFullscreen').title = modoFullscreen ? 'Sair da tela cheia' : 'Modo tela cheia';
}

// =========================================================
// BADGE DE ADIANTAMENTOS URGENTES NA NAV
// =========================================================
function atualizarBadgeAdi(lista) {
    const hoje = new Date(); hoje.setHours(0,0,0,0);
    const urgentes = lista.filter(a => {
        const diff = Math.ceil((new Date(a.venc) - hoje) / (1000*60*60*24));
        return diff <= 7;
    }).length;
    const badge = document.getElementById('badgeAdi');
    if (urgentes > 0) {
        badge.textContent = urgentes;
        badge.style.display = 'inline-flex';
    } else {
        badge.style.display = 'none';
    }
}

// =========================================================
// EXPORTAR QUALQUER TABELA COMO CSV
// =========================================================
function exportarTabelaCSV(tabelaId, nomeArquivo) {
    const tabela = document.getElementById(tabelaId);
    if (!tabela) return;
    const linhas = tabela.querySelectorAll('tr');
    const csv = [];
    linhas.forEach(tr => {
        const cols = [...tr.querySelectorAll('th, td')].map(td => {
            const txt = td.innerText.replace(/\n/g, ' ').trim();
            return `"${txt.replace(/"/g, '""')}"`;
        });
        if (cols.length) csv.push(cols.join(';'));
    });
    if (!csv.length) return;
    const hoje = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    const blob = new Blob(['\uFEFF' + csv.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${nomeArquivo}_${hoje}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
    tocarSomMSN();
}
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
        // Envia senha em texto puro — o GS faz o hash server-side (HTTPS protege o transporte)
        const res  = await fetch(`${URL_SCRIPT}?action=login&login=${encodeURIComponent(u)}&senha=${encodeURIComponent(s)}`);
        const data = await res.json();

        if (data.ok) {
            loginAtual = u;
            sessaoAtual = {
                token:  data.token,
                expira: Date.now() + (8 * 60 * 60 * 1000)
            };
            setTimeout(() => {
                mostrarToast('⏱️ Sua sessão expirou. Faça login novamente.', 'warning', 6000);
                setTimeout(() => logout(), 3000);
            }, 8 * 60 * 60 * 1000);

            if (data.primeiroAcesso) {
                document.getElementById('loginScreen').style.display = 'none';
                document.getElementById('primeiroAcessoScreen').style.display = 'flex';
                document.getElementById('nomeBoasVindas').innerText = data.nome.split(' ')[0];
            } else if (LOGINS_ACESSO_MULTIPLO.includes(u)) {
                // Mostra seletor de interface
                _dadosLoginPendente = data;
                document.getElementById('btnEntrar').style.display = 'none';
                document.getElementById('seletorInterface').style.display = 'flex';
                btn.disabled = false;
                btn.innerText = 'ENTRAR NO SISTEMA';
                return;
            } else {
                entrarNoSistema(data);
                registrarLog('LOGIN', 'Acesso ao sistema');
            }
        } else if (data.bloqueado) {
            mostrarErro('loginErro', `🔒 Muitas tentativas. Tente novamente em ${data.restante} minuto(s).`);
        } else {
            mostrarErro('loginErro', `⚠️ Usuário ou senha incorretos. ${data.tentativas ? `(${data.tentativas}/5 tentativas)` : ''}`);
        }
    } catch (err) {
        mostrarErro('loginErro', '⚠️ Erro ao conectar com o servidor.');
    } finally {
        btn.disabled = false;
        btn.innerText = "ENTRAR NO SISTEMA";
    }
}

// --- TROCA DE SENHA (PRIMEIRO ACESSO) ---

// =========================================================
// PRIMEIRO ACESSO — FLUXO COM CÓDIGO
// =========================================================
function voltarPaStep1() {
    document.getElementById('paStep1').style.display = 'block';
    document.getElementById('paStep2').style.display = 'none';
    document.getElementById('paErro1').innerText = '';
}

async function enviarCodigoPrimeiroAcesso() {
    const email = document.getElementById('primeiroEmail').value.trim();
    if (!email || !email.includes('@')) {
        document.getElementById('paErro1').innerText = '⚠️ Informe um email válido.'; return;
    }
    const btn = document.querySelector('#paStep1 .btn-login-final');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> ENVIANDO...';
    try {
        const res  = await fetch(`${URL_SCRIPT}?action=enviarCodigoPrimeiroAcesso&login=${encodeURIComponent(loginAtual)}&email=${encodeURIComponent(email)}`);
        const data = await res.json();
        if (!data.ok) { document.getElementById('paErro1').innerText = data.erro || '⚠️ Erro ao enviar código.'; return; }
        document.getElementById('paMsgCodigo').innerHTML = `Código enviado para <b>${data.emailMask}</b>. Expira em 15 minutos.`;
        document.getElementById('paStep1').style.display = 'none';
        document.getElementById('paStep2').style.display = 'block';
        setTimeout(() => document.getElementById('paCodigo').focus(), 100);
    } catch (e) {
        document.getElementById('paErro1').innerText = '⚠️ Erro ao conectar. Tente novamente.';
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-paper-plane-tilt"></i> ENVIAR CÓDIGO DE CONFIRMAÇÃO';
    }
}

async function confirmarNovaSenha() {
    const codigo   = document.getElementById('paCodigo').value.trim();
    const nova     = document.getElementById('novaSenha').value.trim();
    const confirma = document.getElementById('confirmaSenha').value.trim();
    const email    = document.getElementById('primeiroEmail').value.trim();

    if (codigo.length !== 6) { mostrarErro('trocaErro', '⚠️ Digite o código de 6 dígitos.'); return; }
    if (nova.length < 6)     { mostrarErro('trocaErro', '⚠️ A senha deve ter pelo menos 6 caracteres.'); return; }
    if (nova !== confirma)   { mostrarErro('trocaErro', '⚠️ As senhas não coincidem.'); return; }

    const btn = document.getElementById('btnConfirmarSenha');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';
    try {
        const res  = await fetch(`${URL_SCRIPT}?action=trocarSenha&login=${encodeURIComponent(loginAtual)}&novaSenha=${encodeURIComponent(nova)}&email=${encodeURIComponent(email)}&codigo=${encodeURIComponent(codigo)}`);
        const data = await res.json();
        if (data.ok) {
            sessaoAtual = { token: data.token, expira: Date.now() + (8 * 60 * 60 * 1000) };
            entrarNoSistema(data);
            registrarLog('PRIMEIRO ACESSO', 'Senha criada e email cadastrado');
        } else {
            mostrarErro('trocaErro', data.erro || '⚠️ Código inválido ou expirado.');
        }
    } catch (err) {
        mostrarErro('trocaErro', '⚠️ Erro ao conectar com o servidor.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-check-circle"></i> SALVAR E ENTRAR';
    }
}

// =========================================================
// HISTÓRICO DE NOTAS REGISTRADAS (DIGITADAS / RECEBIDAS)
// =========================================================

const HIST_PAGE_SIZE  = 20;
let _histNotas        = [];
let _histNotasFiltrado = [];
let _histNotasPagina  = 1;

async function carregarHistoricoNotas() {
    if (abaAtual !== 'Digitadas' && abaAtual !== 'Recebimento') return;

    const section = document.getElementById('historicoNotasSection');
    section.style.display = 'block';
    document.getElementById('hist-notas-loading').style.display = 'flex';
    document.querySelector('#tabelaHistoricoNotas tbody').innerHTML = '';
    document.getElementById('histNotasPaginacao').innerHTML = '';
    document.getElementById('historicoResumo').textContent = 'Carregando...';
    document.getElementById('histBuscaInput').value = '';

    try {
        const primeiroNome = usuarioAtual.nome.split(' ')[0];
        const nomeCompleto = encodeURIComponent(usuarioAtual.nome);
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getHistoricoNotas&aba=${encodeURIComponent(abaAtual)}&primeiroNome=${encodeURIComponent(primeiroNome)}&nomeCompleto=${nomeCompleto}`));
        const data = await res.json();
        if (data.erro) throw new Error(data.erro);
        _histNotas        = data;
        _histNotasFiltrado = data;
        _histNotasPagina  = 1;
        renderizarHistoricoNotas();
    } catch (e) {
        document.getElementById('historicoResumo').textContent = 'Erro ao carregar.';
        mostrarToast('Erro ao carregar histórico: ' + e.message, 'error');
    } finally {
        document.getElementById('hist-notas-loading').style.display = 'none';
    }
}

function filtrarHistoricoNotas() {
    const q = document.getElementById('histBuscaInput').value.trim().toLowerCase();
    _histNotasFiltrado = q
        ? _histNotas.filter(n =>
            (n.nf         || '').toLowerCase().includes(q) ||
            (n.fornecedor || '').toLowerCase().includes(q) ||
            (n.razaoSocial|| '').toLowerCase().includes(q) ||
            (n.setor      || '').toLowerCase().includes(q)
          )
        : _histNotas;
    _histNotasPagina = 1;
    renderizarHistoricoNotas();
}

function renderizarHistoricoNotas() {
    const tbody     = document.querySelector('#tabelaHistoricoNotas tbody');
    const paginacao = document.getElementById('histNotasPaginacao');
    const resumo    = document.getElementById('historicoResumo');
    tbody.innerHTML    = '';
    paginacao.innerHTML = '';

    const lista = _histNotasFiltrado;

    if (!lista.length) {
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:24px;color:var(--text-muted);">Nenhuma nota encontrada.</td></tr>`;
        resumo.textContent = '0 notas';
        return;
    }

    const total     = lista.length;
    const totalPags = Math.ceil(total / HIST_PAGE_SIZE);
    const inicio    = (_histNotasPagina - 1) * HIST_PAGE_SIZE;
    const fim       = Math.min(inicio + HIST_PAGE_SIZE, total);
    const pagina    = lista.slice(inicio, fim);

    resumo.textContent = `${total} nota(s) — página ${_histNotasPagina} de ${totalPags}`;

    // Mostra coluna retirada só na aba Recebimento
    const isRecebimento = abaAtual === 'Recebimento';
    document.getElementById('thRetirada').style.display = isRecebimento ? '' : 'none';
    const colSpan = isRecebimento ? 10 : 9;

    pagina.forEach((n, idx) => {
        const editada  = n.editada === true || n.editada === 'EDITADA';
        const retirada = n.retirada === true || !!n.setorRetirada;
        const rowStyle = editada ? 'background:rgba(245,158,11,0.07);' : retirada ? 'background:rgba(2,132,199,0.05);' : '';
        const badge    = editada ? `<span style="background:#f59e0b22;color:#f59e0b;font-size:9px;font-weight:900;padding:2px 6px;border-radius:5px;margin-left:6px;">EDITADA</span>` : '';
        const absIdx   = _histNotas.indexOf(n);

        const btnRetirada = isRecebimento
            ? retirada
                ? `<td style="text-align:center;"><span style="font-size:10px;font-weight:800;color:var(--accent);">${n.setorRetirada || 'Retirado'}</span></td>`
                : `<td style="text-align:center;"><button onclick="abrirRetirada(${absIdx})"
                    style="background:transparent;border:1.5px solid var(--accent);border-radius:8px;padding:5px 8px;cursor:pointer;color:var(--accent);" title="Registrar retirada">
                    <i class="ph ph-package"></i></button></td>`
            : '';

        tbody.innerHTML += `
            <tr style="${rowStyle}">
                <td style="font-size:12px;">${n.data || '—'}</td>
                <td><b>${n.nf || '—'}</b>${badge}</td>
                <td>${n.fornecedor || '—'}</td>
                <td style="font-size:11px;">${n.razaoSocial || '—'}</td>
                <td style="font-size:12px;">${n.vencimento || '—'}</td>
                <td>R$ ${formatarValor(n.valor)}</td>
                <td style="font-size:11px;">${n.setor || '—'}</td>
                <td style="font-size:11px;">${n.tipo || '—'}</td>
                ${btnRetirada}
                <td style="text-align:center;">
                    <button onclick="abrirEdicaoNota(${absIdx})"
                        style="background:transparent;border:1.5px solid var(--border);border-radius:8px;padding:5px 8px;cursor:pointer;color:var(--accent);" title="Editar">
                        <i class="ph ph-pencil-simple"></i>
                    </button>
                </td>
            </tr>`;
    });

    if (totalPags > 1) {
        const btn = (label, page, ativo) =>
            `<button onclick="mudarPaginaHist(${page})" style="padding:6px 14px;border-radius:8px;border:1.5px solid ${ativo ? 'var(--accent)' : 'var(--border)'};background:${ativo ? 'var(--accent)' : 'transparent'};color:${ativo ? '#fff' : 'var(--text-main)'};font-weight:800;font-size:12px;cursor:${ativo ? 'default' : 'pointer'}">${label}</button>`;
        if (_histNotasPagina > 1) paginacao.innerHTML += btn('‹ Ant', _histNotasPagina - 1, false);
        for (let p = Math.max(1, _histNotasPagina - 2); p <= Math.min(totalPags, _histNotasPagina + 2); p++) {
            paginacao.innerHTML += btn(p, p, p === _histNotasPagina);
        }
        if (_histNotasPagina < totalPags) paginacao.innerHTML += btn('Prox ›', _histNotasPagina + 1, false);
    }
}

function mudarPaginaHist(p) {
    _histNotasPagina = p;
    renderizarHistoricoNotas();
    document.getElementById('historicoNotasSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── RETIRADA ──────────────────────────────────────────────

function abrirRetirada(idx) {
    const n = _histNotas[idx];
    if (!n) return;
    document.getElementById('retiradaNfLabel').textContent = n.nf;
    document.getElementById('retiradaRowIndex').value = n.rowIndex;
    document.getElementById('retiradaSetor').value = '';
    document.getElementById('modalRetirada').style.display = 'flex';
    setTimeout(() => document.getElementById('retiradaSetor').focus(), 100);
}

async function confirmarRetirada() {
    const rowIndex = parseInt(document.getElementById('retiradaRowIndex').value);
    const setor    = document.getElementById('retiradaSetor').value.trim();
    if (!setor) { mostrarToast('Informe o setor que retirou.', 'warning'); return; }

    const btn = document.querySelector('#modalRetirada .btn-saida-confirmar');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'registrarRetirada', rowIndex, setorRetirada: setor })
        });

        // Atualiza cache local
        const nota = _histNotas.find(n => n.rowIndex === rowIndex);
        if (nota) { nota.setorRetirada = setor; nota.retirada = true; }

        fecharModal('modalRetirada');
        renderizarHistoricoNotas();
        mostrarToast('Retirada registrada com sucesso!', 'success');
        registrarLog('RETIRADA', `NF ${document.getElementById('retiradaNfLabel').textContent} — Setor: ${setor}`);
    } catch (e) {
        mostrarToast('Erro ao registrar retirada.', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-package"></i> CONFIRMAR RETIRADA';
    }
}

// ── EDIÇÃO ────────────────────────────────────────────────

function abrirEdicaoNota(idx) {
    const n = _histNotas[idx];
    if (!n) return;
    document.getElementById('edit_nf').value         = n.nf || '';
    document.getElementById('edit_fornecedor').value = n.fornecedor || '';
    document.getElementById('edit_razao').value      = n.razaoSocial || '';
    document.getElementById('edit_data').value       = converterDataParaInput(n.data);
    document.getElementById('edit_vencimento').value = converterDataParaInput(n.vencimento);
    document.getElementById('edit_valor').value      = n.valor || '';
    document.getElementById('edit_setor').value      = n.setor || '';
    document.getElementById('edit_rowIndex').value   = n.rowIndex;
    document.getElementById('edit_aba').value        = abaAtual;
    document.getElementById('modalEditarNota').style.display = 'flex';
}

function converterDataParaInput(dataStr) {
    if (!dataStr) return '';
    // DD/MM/YYYY → YYYY-MM-DD
    const m = dataStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    return dataStr;
}

async function salvarEdicaoNota() {
    const rowIndex   = document.getElementById('edit_rowIndex').value;
    const aba        = document.getElementById('edit_aba').value;
    const primeiroNome = usuarioAtual.nome.split(' ')[0];

    const payload = {
        rowIndex:    parseInt(rowIndex),
        aba:         aba,
        responsavel: primeiroNome,
        nf:          document.getElementById('edit_nf').value.trim(),
        fornecedor:  document.getElementById('edit_fornecedor').value.trim(),
        razaoSocial: document.getElementById('edit_razao').value.trim(),
        data:        document.getElementById('edit_data').value,
        vencimento:  document.getElementById('edit_vencimento').value,
        valor:       document.getElementById('edit_valor').value.trim(),
        setor:       document.getElementById('edit_setor').value.trim(),
    };

    const btn = document.querySelector('#modalEditarNota .btn-salvar-usuario');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=editarNota&solicitante=${encodeURIComponent(loginAtual)}`), {
            method: 'POST',
            mode:   'no-cors',
            body:   JSON.stringify({ tipo: 'editarNota', ...payload })
        });

        // Atualiza cache local
        const nota = _histNotas.find(n => n.rowIndex === parseInt(rowIndex));
        if (nota) {
            nota.nf         = payload.nf;
            nota.fornecedor = payload.fornecedor;
            nota.razaoSocial = payload.razaoSocial;
            nota.data       = formatarDataBR(payload.data);
            nota.vencimento = formatarDataBR(payload.vencimento);
            nota.valor      = payload.valor;
            nota.setor      = payload.setor;
            nota.editada    = true;
        }

        fecharModal('modalEditarNota');
        renderizarHistoricoNotas();
        mostrarToast('✅ Nota editada com sucesso!', 'success');
        registrarLog('EDITAR NOTA', `NF ${payload.nf} — aba ${aba}`);
    } catch (e) {
        mostrarToast('Erro ao salvar edição.', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-check-circle"></i> SALVAR EDIÇÃO';
    }
}
let _resetLoginPendente = null;

function abrirEsqueciSenha() {
    document.getElementById('resetStep1').style.display = 'block';
    document.getElementById('resetStep2').style.display = 'none';
    document.getElementById('resetStep3').style.display = 'none';
    document.getElementById('resetLogin').value  = '';
    document.getElementById('resetCodigo').value = '';
    document.getElementById('resetErro1').innerText = '';
    _resetLoginPendente = null;
    document.getElementById('modalEsqueciSenha').style.display = 'flex';
}

function voltarStep1() {
    document.getElementById('resetStep1').style.display = 'block';
    document.getElementById('resetStep2').style.display = 'none';
    document.getElementById('resetErro1').innerText = '';
}

async function solicitarCodigoReset() {
    const login = document.getElementById('resetLogin').value.trim().toLowerCase();
    if (!login) { document.getElementById('resetErro1').innerText = '⚠️ Informe o login.'; return; }

    const btn = document.querySelector('#resetStep1 .btn-salvar-usuario');
    btn.disabled = true; btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> ENVIANDO...';

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=solicitarReset&login=${encodeURIComponent(login)}`);
        const data = await res.json();
        if (!data.ok) { document.getElementById('resetErro1').innerText = data.erro || '⚠️ Usuário não encontrado ou sem email cadastrado.'; return; }

        _resetLoginPendente = login;
        document.getElementById('resetEmailMsg').innerHTML = `Código enviado para <b>${data.emailMask}</b>. Expira em 15 minutos.`;
        document.getElementById('resetStep1').style.display = 'none';
        document.getElementById('resetStep2').style.display = 'block';
        setTimeout(() => document.getElementById('resetCodigo').focus(), 100);
    } catch (e) {
        document.getElementById('resetErro1').innerText = '⚠️ Erro ao enviar código. Tente novamente.';
    } finally {
        btn.disabled = false; btn.innerHTML = '<i class="ph ph-paper-plane-tilt"></i> ENVIAR CÓDIGO';
    }
}

async function validarCodigoReset() {
    const codigo = document.getElementById('resetCodigo').value.trim();
    if (codigo.length !== 6) { document.getElementById('resetErro2').innerText = '⚠️ Digite o código de 6 dígitos.'; return; }

    const btn = document.querySelector('#resetStep2 .btn-salvar-usuario');
    btn.disabled = true; btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> VALIDANDO...';

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=validarCodigo&login=${encodeURIComponent(_resetLoginPendente)}&codigo=${encodeURIComponent(codigo)}`);
        const data = await res.json();
        if (!data.ok) { document.getElementById('resetErro2').innerText = data.erro || '⚠️ Código inválido ou expirado.'; return; }
        document.getElementById('resetStep2').style.display = 'none';
        document.getElementById('resetStep3').style.display = 'block';
        setTimeout(() => document.getElementById('resetNovaSenha').focus(), 100);
    } catch (e) {
        document.getElementById('resetErro2').innerText = '⚠️ Erro ao validar código.';
    } finally {
        btn.disabled = false; btn.innerHTML = '<i class="ph ph-check-circle"></i> VALIDAR CÓDIGO';
    }
}

async function confirmarResetSenha() {
    const nova    = document.getElementById('resetNovaSenha').value.trim();
    const confirma = document.getElementById('resetConfirmaSenha').value.trim();
    if (nova.length < 6)   { document.getElementById('resetErro3').innerText = '⚠️ Senha deve ter pelo menos 6 caracteres.'; return; }
    if (nova !== confirma) { document.getElementById('resetErro3').innerText = '⚠️ As senhas não coincidem.'; return; }

    const novaHash = nova; // GS faz o hash server-side
    const codigo   = document.getElementById('resetCodigo').value.trim();
    const btn = document.querySelector('#resetStep3 .btn-salvar-usuario');
    btn.disabled = true; btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';

    try {
        const res  = await fetch(`${URL_SCRIPT}?action=resetarComCodigo&login=${encodeURIComponent(_resetLoginPendente)}&codigo=${encodeURIComponent(codigo)}&novaSenha=${encodeURIComponent(novaHash)}`);
        const data = await res.json();
        if (!data.ok) { document.getElementById('resetErro3').innerText = data.erro || '⚠️ Erro ao salvar senha.'; return; }
        fecharModal('modalEsqueciSenha');
        mostrarToast('✅ Senha redefinida com sucesso! Faça login.', 'success', 5000);
        _resetLoginPendente = null;
    } catch (e) {
        document.getElementById('resetErro3').innerText = '⚠️ Erro ao salvar. Tente novamente.';
    } finally {
        btn.disabled = false; btn.innerHTML = '<i class="ph ph-check-circle"></i> SALVAR NOVA SENHA';
    }
}

// =========================================================
// GRÁFICO PIZZA — RESUMO PROJEÇÃO
// =========================================================

function desenharPizzaProjecao(itens) {
    const canvas  = document.getElementById('projPizzaChart');
    const legenda = document.getElementById('projPizzaLegenda');
    if (!canvas || !itens || !itens.length) return;

    const total           = itens.length;
    const comSaldo        = itens.filter(i => i.saldoCD > 0).length;
    const comEmpenho      = itens.filter(i => i.temEmpenho).length;
    const comRP           = itens.filter(i => i.temRP).length;
    const compraFZ        = itens.filter(i => i.zeradoSemCobertura && i.cobertura <= 25).length;
    const zeradosCobertura = itens.filter(i => i.cobertura <= 0).length;
    const outros          = Math.max(0, total - comSaldo - comEmpenho - comRP - compraFZ);

    const fatias = [
        { label: 'Saldo CD',   valor: comSaldo,        cor: '#0284c7' },
        { label: 'Empenho',    valor: comEmpenho,      cor: '#f59e0b' },
        { label: 'RP',         valor: comRP,           cor: '#10b981' },
        { label: 'Compra FZ (Críticos)', valor: compraFZ, cor: '#ef4444' },
        { label: 'Outros',     valor: outros > 0 ? outros : 0, cor: '#94a3b8' },
    ].filter(f => f.valor > 0)
     .sort((a, b) => b.valor - a.valor);

    const ctx  = canvas.getContext('2d');
    const cx   = canvas.width  / 2;
    const cy   = canvas.height / 2;
    const r    = Math.min(cx, cy) - 8;
    let angulo = -Math.PI / 2;

    ctx.clearRect(0, 0, canvas.width, canvas.height);

    fatias.forEach(f => {
        const pct   = f.valor / total;
        const angFim = angulo + pct * 2 * Math.PI;
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, angulo, angFim);
        ctx.closePath();
        ctx.fillStyle = f.cor;
        ctx.fill();
        angulo = angFim;
    });

    // Furo central (donut)
    ctx.beginPath();
    ctx.arc(cx, cy, r * 0.52, 0, 2 * Math.PI);
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--bg-card').trim() || '#ffffff';
    ctx.fill();

    // Centro: quantidade de Compra FZ
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 22px Calibri, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(zeradosCobertura, cx, cy - 8);
    ctx.font = 'bold 9px Calibri, sans-serif';
    ctx.fillStyle = 'rgba(255,255,255,0.7)';
    ctx.fillText('Zerados', cx, cy + 10);

    // Legenda
    legenda.innerHTML = fatias.map(f => `
        <div style="display:flex;align-items:center;gap:8px;">
            <span style="width:12px;height:12px;border-radius:3px;background:${f.cor};flex-shrink:0;display:inline-block;"></span>
            <span style="color:var(--text-muted)">${f.label}</span>
            <b style="color:var(--text-main);margin-left:auto;padding-left:12px;">${f.valor}</b>
            <span style="color:var(--text-muted);font-size:11px;">(${Math.round(f.valor/total*100)}%)</span>
        </div>`).join('');
}

// ── MINI PIZZA PARA CARDS DA DASHBOARD ────────────────────
function desenharMiniPizza(canvasId, legendaId, totalId, itens) {
    const canvas  = document.getElementById(canvasId);
    const legenda = document.getElementById(legendaId);
    const totalEl = document.getElementById(totalId);
    if (!canvas || !itens) return;

    const total           = itens.length;
    const comSaldo        = itens.filter(i => i.saldoCD > 0).length;
    const comEmpenho      = itens.filter(i => i.temEmpenho).length;
    const comRP           = itens.filter(i => i.temRP).length;
    const compraFZ        = itens.filter(i => i.zeradoSemCobertura && i.cobertura <= 25).length;
    const zeradosCobertura = itens.filter(i => i.cobertura <= 0).length;
    const outros          = Math.max(0, total - comSaldo - comEmpenho - comRP - compraFZ);

    if (totalEl) totalEl.textContent = total + ' itens no total';

    const fatias = [
        { label: 'Saldo CD',   valor: comSaldo,        cor: '#0284c7' },
        { label: 'Empenho',    valor: comEmpenho,      cor: '#f59e0b' },
        { label: 'RP',         valor: comRP,           cor: '#10b981' },
        { label: 'Compra FZ (Críticos)', valor: compraFZ, cor: '#ef4444' },
        { label: 'Outros',     valor: outros,          cor: '#94a3b8' },
    ].filter(f => f.valor > 0)
     .sort((a, b) => b.valor - a.valor); // ordena por maior % primeiro

    const ctx = canvas.getContext('2d');
    const cx  = canvas.width / 2, cy = canvas.height / 2;
    const r   = Math.min(cx, cy) - 4;
    let ang   = -Math.PI / 2;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    fatias.forEach(f => {
        const end = ang + (f.valor / total) * 2 * Math.PI;
        ctx.beginPath(); ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, ang, end); ctx.closePath();
        ctx.fillStyle = f.cor; ctx.fill();
        ang = end;
    });

    const bgMini = getComputedStyle(document.documentElement).getPropertyValue('--bg-card').trim() || '#1e293b';
    ctx.beginPath(); ctx.arc(cx, cy, r * 0.52, 0, 2 * Math.PI);
    ctx.fillStyle = bgMini; ctx.fill();

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 20px Calibri, sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(zeradosCobertura, cx, cy - 7);
    ctx.font = 'bold 8px Calibri, sans-serif';
    ctx.fillStyle = 'rgba(255,255,255,0.7)';
    ctx.fillText('Zerados', cx, cy + 9);

    legenda.innerHTML = fatias.map(f =>
        `<div style="display:flex;align-items:center;gap:6px;">
            <span style="width:11px;height:11px;border-radius:3px;background:${f.cor};flex-shrink:0;"></span>
            <span style="color:var(--text-muted);flex:1;font-size:12px;">${f.label}</span>
            <b style="color:var(--text-main);font-size:13px;">${f.valor}</b>
            <span style="color:var(--text-muted);font-size:11px;min-width:32px;text-align:right;">${Math.round(f.valor/total*100)}%</span>
        </div>`
    ).join('');
}

// ── MINI GRÁFICO DE PROCESSO DE COMPRA PARA CARDS DA DASHBOARD ──────────────
function desenharMiniProcesso(canvasId, legendaId, itens, mapaStatus) {
    const canvas  = document.getElementById(canvasId);
    const legenda = document.getElementById(legendaId);
    if (!canvas || !itens) return;

    const fases = [
        { label: 'Sem processo',   cor: 'rgba(100,116,139,0.45)', chave: null },
        { label: 'Em Orçamento',   cor: '#f59e0b', chave: 'EM ORÇAMENTO' },
        { label: 'Req. Compra',    cor: '#3b82f6', chave: 'REQUISIÇÃO DE COMPRA' },
        { label: 'Ordem Compra',   cor: '#8b5cf6', chave: 'ORDEM DE COMPRA' },
        { label: 'Ag. Entrega',    cor: '#10b981', chave: 'AGUARDANDO ENTREGA DO FORNECEDOR' },
    ];

    const contagem = {};
    fases.forEach(f => { if (f.chave) contagem[f.chave] = 0; });
    Object.values(mapaStatus).forEach(sc => {
        if (contagem[sc.status] !== undefined) contagem[sc.status]++;
    });
    const comProcesso = Object.values(contagem).reduce((a, b) => a + b, 0);
    const semProcesso = Math.max(0, itens.length - comProcesso);
    const totalValores = itens.length || 1;

    const fatias = fases.map(f => ({
        label: f.label, cor: f.cor,
        valor: f.chave ? contagem[f.chave] : semProcesso
    })).filter(f => f.valor > 0);

    const ctx = canvas.getContext('2d');
    const cx  = canvas.width / 2, cy = canvas.height / 2;
    const r   = Math.min(cx, cy) - 4;
    let ang   = -Math.PI / 2;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    fatias.forEach(f => {
        const end = ang + (f.valor / totalValores) * 2 * Math.PI;
        ctx.beginPath(); ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, ang, end); ctx.closePath();
        ctx.fillStyle = f.cor; ctx.fill();
        ang = end;
    });

    const bg = getComputedStyle(document.documentElement).getPropertyValue('--bg-card').trim() || '#1e293b';
    ctx.beginPath(); ctx.arc(cx, cy, r * 0.52, 0, 2 * Math.PI);
    ctx.fillStyle = bg; ctx.fill();

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 20px Calibri, sans-serif';
    ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillText(comProcesso, cx, cy - 7);
    ctx.font = 'bold 8px Calibri, sans-serif';
    ctx.fillStyle = 'rgba(255,255,255,0.6)';
    ctx.fillText('processo', cx, cy + 9);

    legenda.innerHTML = fatias.map(f =>
        `<div style="display:flex;align-items:center;gap:6px;">
            <span style="width:9px;height:9px;border-radius:50%;background:${f.cor};flex-shrink:0;"></span>
            <span style="color:var(--text-muted);flex:1;font-size:11px;">${f.label}</span>
            <b style="color:var(--text-main);font-size:12px;">${f.valor}</b>
            <span style="color:var(--text-muted);font-size:10px;min-width:28px;text-align:right;">${Math.round(f.valor/totalValores*100)}%</span>
        </div>`
    ).join('');
}

async function carregarPizzasDashboard() {
    const wrap = document.getElementById('dashProjCards');
    if (wrap) wrap.style.display = 'block';

    try {
        // Busca a lista real de compradores para pegar os nomes corretos
        const resComp = await fetch(addAuth(`${URL_SCRIPT}?action=getCompradores`));
        const compradores = await resComp.json();
        if (!compradores || !compradores.length) return;

        // Limpa os cards existentes e reconstrói dinamicamente
        const grid = wrap.querySelector('.proj-pizza-grid');

        for (const comp of compradores) {
            const nomeUsuario    = comp.nomeUsuario;
            const canvasId       = `dashPizza_${nomeUsuario}`;
            const legendaId      = `dashLegenda_${nomeUsuario}`;
            const totalId        = `dashTotal_${nomeUsuario}`;
            const canvasProcesso = `dashProcesso_${nomeUsuario}`;
            const legendaProcesso= `dashProcessoLeg_${nomeUsuario}`;

            // Cria card se não existir
            if (!document.getElementById(canvasId) && grid) {
                const card = document.createElement('div');
                card.className = 'stat-card';
                card.style.cssText = 'padding:24px;display:flex;flex-direction:column;gap:16px;';
                card.innerHTML = `
                    <div style="display:flex;align-items:center;justify-content:space-between;">
                        <div>
                            <small style="font-weight:900;letter-spacing:1px;font-size:13px;">PROJEÇÃO — ${nomeUsuario}</small>
                            <p id="${totalId}" style="font-size:12px;color:var(--text-muted);margin-top:2px;font-weight:600;">carregando...</p>
                        </div>
                        <i class="ph ph-shopping-cart-simple" style="font-size:24px;color:var(--accent);opacity:0.6;"></i>
                    </div>
                    <div style="display:flex;gap:24px;flex-wrap:wrap;justify-content:center;">
                        <!-- Gráfico 1: Situação dos itens -->
                        <div style="display:flex;flex-direction:column;align-items:center;gap:6px;">
                            <span style="font-size:9px;font-weight:800;text-transform:uppercase;color:var(--text-muted);letter-spacing:0.5px;">SITUAÇÃO</span>
                            <div style="display:flex;align-items:center;gap:16px;">
                                <canvas id="${canvasId}" width="110" height="110" style="flex-shrink:0;"></canvas>
                                <div id="${legendaId}" style="display:flex;flex-direction:column;gap:6px;font-size:11px;flex:1;"></div>
                            </div>
                        </div>
                        <div style="width:1px;background:var(--border);align-self:stretch;"></div>
                        <!-- Gráfico 2: Processo de compra -->
                        <div style="display:flex;flex-direction:column;align-items:center;gap:6px;">
                            <span style="font-size:9px;font-weight:800;text-transform:uppercase;color:var(--text-muted);letter-spacing:0.5px;">PROCESSO</span>
                            <div style="display:flex;align-items:center;gap:16px;">
                                <canvas id="${canvasProcesso}" width="110" height="110" style="flex-shrink:0;"></canvas>
                                <div id="${legendaProcesso}" style="display:flex;flex-direction:column;gap:6px;font-size:11px;flex:1;"></div>
                            </div>
                        </div>
                    </div>`;
                grid.appendChild(card);
            }

            // Carrega dados da projeção
            try {
                const res  = await fetch(addAuth(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`));
                const data = await res.json();
                if (!data.erro && data.itens) {
                    desenharMiniPizza(canvasId, legendaId, totalId, data.itens);
                    // Carrega statusCompra para este comprador e desenha o gráfico de processo
                    const resStatus = await fetch(`${URL_SCRIPT}?action=getStatusCompra&comprador=${encodeURIComponent(nomeUsuario)}&t=${Date.now()}`);
                    const dataStatus = await resStatus.json();
                    const mapa = {};
                    (dataStatus.itens || []).forEach(i => { mapa[i.codigo] = i; });
                    desenharMiniProcesso(canvasProcesso, legendaProcesso, data.itens, mapa);
                } else {
                    const el = document.getElementById(totalId);
                    if (el) el.textContent = data.erro || 'Sem dados';
                }
            } catch(e) {
                const el = document.getElementById(totalId);
                if (el) el.textContent = 'Erro ao carregar';
            }
        }
    } catch(e) {
        console.warn('Erro ao carregar pizzas dashboard:', e);
    }
}

let _dashFsAtivo = false;
let _clockInterval = null;

function toggleDashFullscreen() {
    _dashFsAtivo = !_dashFsAtivo;
    document.body.classList.toggle('fullscreen-mode', _dashFsAtivo);
    document.body.classList.toggle('dash-fullscreen-mode', _dashFsAtivo);

    const icon  = document.getElementById('dashFullscreenIcon');
    const label = document.getElementById('dashFullscreenLabel');
    if (icon)  icon.className  = _dashFsAtivo ? 'ph ph-arrows-in'  : 'ph ph-arrows-out';
    if (label) label.textContent = _dashFsAtivo ? 'SAIR' : 'TELA CHEIA';

    if (_dashFsAtivo) {
        _clockInterval = setInterval(atualizarRelogioFs, 1000);
        atualizarRelogioFs();
        // Fullscreen API nativo (funciona em TVs e Chrome)
        if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
    } else {
        clearInterval(_clockInterval);
        if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {});
    }
}

function atualizarRelogioFs() {
    const el = document.getElementById('dashClock');
    if (!el) return;
    const now = new Date();
    el.textContent = now.toLocaleTimeString('pt-BR');
}

// Sai do fullscreen ao pressionar ESC
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && _dashFsAtivo) toggleDashFullscreen();
});

// Sai do modo ao sair do fullscreen nativo (ex: F11)
document.addEventListener('fullscreenchange', function() {
    if (!document.fullscreenElement && _dashFsAtivo) {
        _dashFsAtivo = false;
        document.body.classList.remove('dash-fullscreen-mode');
        document.body.classList.remove('fullscreen-mode');
        clearInterval(_clockInterval);
        const icon  = document.getElementById('dashFullscreenIcon');
        const label = document.getElementById('dashFullscreenLabel');
        if (icon)  icon.className   = 'ph ph-arrows-out';
        if (label) label.textContent = 'TELA CHEIA';
    }
});
function temPermissao(p) {
    return usuarioAtual.permissoes && usuarioAtual.permissoes.includes(p);
}

function entrarNoSistema(data) {
    alertaJaExibido = false;
    alertaAdiJaExibido = false;
    alertaProjJaExibido = false;
    alertaProjPendente = false;

    // Permissões: vem como string "digitador,gestor,comprador" → array
    const permsRaw = data.permissoes || data.role || 'digitador';
    usuarioAtual = {
        nome: data.nome,
        role: data.role,
        permissoes: permsRaw.toString().split(',').map(p => p.trim().toLowerCase()),
        prefixos: data.prefixos ? data.prefixos.toString().split('|').map(p => p.trim()).filter(Boolean) : []
    };

    // Registra os prefixos próprios no dicionário global
    const nomeKey = usuarioAtual.nome.split(' ')[0].toUpperCase();
    prefixosConfig[nomeKey] = usuarioAtual.prefixos;

    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('primeiroAcessoScreen').style.display = 'none';
    document.getElementById('mainHeader').style.display = 'flex';
    document.getElementById('app').style.display = 'block';
    document.getElementById('userNameHeader').innerText = usuarioAtual.nome;
    document.getElementById('floatNav').style.display = 'block';

    // Badge do setor
    const badge = document.getElementById('setorBadge');
    if (badge) {
        if (modoAtual === 'opme') {
            badge.textContent = 'ESPECIAIS';
            badge.style.display = 'inline';
            badge.style.background = 'rgba(99,102,241,0.2)';
            badge.style.color = '#818cf8';
            badge.style.border = '1px solid rgba(99,102,241,0.3)';
        } else if (LOGINS_ACESSO_MULTIPLO.includes(loginAtual)) {
            badge.textContent = 'SUPRIMENTOS';
            badge.style.display = 'inline';
            badge.style.background = 'rgba(6,182,212,0.15)';
            badge.style.color = '#06b6d4';
            badge.style.border = '1px solid rgba(6,182,212,0.25)';
        } else {
            badge.style.display = 'none';
        }
    }

    // Aplica tema salvo do usuário
    const temaSalvo = localStorage.getItem('tema_' + loginAtual);
    if (temaSalvo) {
        document.body.setAttribute('data-theme', temaSalvo);
        const icon = temaSalvo === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
        const headerIcon = document.getElementById('themeIcon');
        if (headerIcon) headerIcon.className = icon;
    }

    // Aplica tema personalizado por usuário (ex: Vitória = rosa)
    const modoTema = localStorage.getItem('tema_modo_' + loginAtual);
    if (modoTema === 'personalizado' || (!modoTema && TEMAS_PERSONALIZADOS[loginAtual?.toLowerCase()])) {
        aplicarTemaUsuario(loginAtual);
    } else {
        document.body.setAttribute('data-user', 'default');
    }
    iniciarNavAnimacoes();
    configurarNavPorUsuario();
    // Init drawer mobile com dados do usuário
    const mobileNome = document.getElementById('mobileNavNome');
    const mobileRole = document.getElementById('mobileNavRole');
    if (mobileNome) mobileNome.textContent = usuarioAtual.nome || 'USUÁRIO';
    if (mobileRole) mobileRole.textContent = (usuarioAtual.role || '').toUpperCase();
    // Redireciona para o módulo escolhido
    if (data._modoOpme) {
        switchTab('ProjecaoOPME');
    } else {
        switchTab('Dashboard');
    }
    carregarBlacklistCache();
    iniciarPolling(); // Polling em tempo real — notificações e status de protocolos
    // Recheck de adiantamentos a cada 30 minutos
    setInterval(async () => {
        const res  = await fetch(URL_SCRIPT).catch(() => null);
        if (!res) return;
        const data = await res.json().catch(() => null);
        if (!data?.adiantamentosSetor) return;
        let lista = data.adiantamentosSetor;
        if (!temPermissao('gestor') && !temPermissao('administrador') && !temPermissao('diretor'))
            lista = lista.filter(a => a.responsavel === usuarioAtual.nome);
        adiantamentosCarregados = lista;
        atualizarBadgeAdi(lista);
    }, 30 * 60 * 1000);
    // Abre modal de boas-vindas após carregar estatísticas
    setTimeout(() => exibirModalBoasVindas(), 800);
}

// =========================================================
// CONFIGURAÇÃO DO NAV POR USUÁRIO (reutilizável)
// =========================================================
function configurarNavPorUsuario() {
    const dir = document.getElementById('navProjecoesDir');
    if (dir) dir.innerHTML = '';
    const fnDir = document.getElementById('fnProjecoesDir');
    if (fnDir) fnDir.innerHTML = '';

    const isOpmeMode = modoAtual === 'opme';

    // Reset geral — esconde tudo primeiro
    ['btn-dashboard','btn-digitadas','btn-recebimento','btn-adiantamento',
     'btn-projecao','btn-admin','btn-protocolos-sup','btn-protocolos-opme',
     'fn-dashboard','fn-digitadas','fn-recebimento','fn-adiantamento',
     'fn-projecao','fn-admin','fn-protocolos-sup','fn-protocolos-opme',
     'fn-div-proj','fn-div-extra','fn-div-admin'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });

    if (isOpmeMode) {
        // Modo OPME — só mostra protocolar notas
        ['btn-protocolos-opme','fn-protocolos-opme'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = id.startsWith('fn') ? 'flex' : 'inline-flex';
        });
        // Projeção OPME — disponível para todos os esp*
        const _fnProjOpme = document.getElementById('fn-projecao-opme');
        const _fnDivProjOpme = document.getElementById('fn-div-proj-opme');
        if (_fnProjOpme) _fnProjOpme.style.display = 'flex';
        if (_fnDivProjOpme) _fnDivProjOpme.style.display = 'block';
        // Admin OPME — só para supmateus
        if (loginAtual === 'supmateus') {
            const d = document.getElementById('fn-div-admin-opme');
            const b = document.getElementById('fn-admin-opme');
            if (d) d.style.display = 'block';
            if (b) b.style.display = 'flex';
        }
        document.getElementById('carrinhoHeaderBtn').style.display = 'none';
        document.getElementById('notifHeaderBtn').style.display = 'flex';
        document.getElementById('userNameHeader').innerText = usuarioAtual.nome;
        return;
    }

    // Modo Suprimentos — restaura abas normais
    ['btn-dashboard','btn-digitadas','btn-recebimento','btn-adiantamento',
     'fn-dashboard','fn-digitadas','fn-recebimento','fn-adiantamento'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = id.startsWith('fn') ? 'flex' : 'inline-flex';
    });

    // Modo Suprimentos — comportamento normal
    const podeVerProjecao = temPermissao('comprador') || temPermissao('diretor') || temPermissao('administrador');

    const _mostrarFnProj = (mostrar, label) => {
        const fnProj = document.getElementById('fn-projecao');
        const fnDivProj = document.getElementById('fn-div-proj');
        if (fnProj) { fnProj.style.display = mostrar ? 'flex' : 'none'; if (label) fnProj.setAttribute('data-tip', label); }
        if (fnDivProj) fnDivProj.style.display = mostrar ? 'block' : 'none';
    };

    if (podeVerProjecao && !temPermissao('diretor')) {
        const primeiroNome = usuarioAtual.nome.split(' ')[0];
        document.getElementById('btnProjecaoLabel').innerText = `PROJ. ${primeiroNome.toUpperCase()}`;
        document.getElementById('btn-projecao').style.display = 'inline-flex';
        _mostrarFnProj(true, `Proj. ${primeiroNome}`);
    } else if (temPermissao('comprador') && temPermissao('diretor')) {
        const primeiroNome = usuarioAtual.nome.split(' ')[0];
        document.getElementById('btnProjecaoLabel').innerText = `PROJ. ${primeiroNome.toUpperCase()}`;
        document.getElementById('btn-projecao').style.display = 'inline-flex';
        _mostrarFnProj(true, `Proj. ${primeiroNome}`);
        criarTabsDiretor();
    } else if (temPermissao('diretor')) {
        document.getElementById('btn-projecao').style.display = 'none';
        _mostrarFnProj(false);
        criarTabsDiretor();
    } else {
        document.getElementById('btn-projecao').style.display = 'none';
        _mostrarFnProj(false);
    }

    document.getElementById('btn-dashboard').style.display   = 'inline-flex';
    document.getElementById('btn-digitadas').style.display   = 'inline-flex';
    document.getElementById('btn-recebimento').style.display = 'inline-flex';
    document.getElementById('btn-adiantamento').style.display = 'inline-flex';

    document.getElementById('carrinhoHeaderBtn').style.display = temPermissao('comprador') || temPermissao('administrador') ? 'flex' : 'none';
    document.getElementById('btn-admin').style.display = temPermissao('administrador') ? 'inline-flex' : 'none';
    document.getElementById('notifHeaderBtn').style.display = temPermissao('comprador') || temPermissao('administrador') ? 'flex' : 'none';
    document.getElementById('userNameHeader').innerText = usuarioAtual.nome;
    // Sincroniza float nav
    const fnAdmin = document.getElementById('fn-admin');
    const fnDivAdmin = document.getElementById('fn-div-admin');
    if (fnAdmin) fnAdmin.style.display = temPermissao('administrador') ? 'flex' : 'none';
    if (fnDivAdmin) fnDivAdmin.style.display = temPermissao('administrador') ? 'block' : 'none';
    configurarNavProtocolos();
    setTimeout(atualizarNavIndicador, 50);
}

// =========================================================
// MODAL BOAS-VINDAS
// =========================================================
function exibirModalBoasVindas() {
    const hora = new Date().getHours();
    const saudacao = hora < 12 ? 'Bom dia,' : hora < 18 ? 'Boa tarde,' : 'Boa noite,';
    const primeiroNome = usuarioAtual.nome.split(' ')[0];
    const hoje = new Date().toLocaleDateString('pt-BR', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
    const hojeStr = hoje.charAt(0).toUpperCase() + hoje.slice(1);

    document.getElementById('bvSaudacao').textContent = saudacao;
    document.getElementById('bvNome').textContent     = primeiroNome;
    document.getElementById('bvData').textContent     = hojeStr;

    // Monta cards de resumo
    const cards = document.getElementById('bvCards');
    cards.innerHTML = '';

    const hoje2 = new Date(); hoje2.setHours(0,0,0,0);
    const urgentes  = adiantamentosCarregados.filter(a => {
        const diff = Math.ceil((new Date(a.venc) - hoje2) / (1000*60*60*24));
        return diff <= 7 && diff >= 0;
    }).length;
    const vencidos  = adiantamentosCarregados.filter(a => new Date(a.venc) < hoje2).length;
    const totalAdi  = adiantamentosCarregados.length;

    const itensCard = [
        {
            icon: 'ph-clock-countdown',
            cor:  urgentes > 0 ? 'var(--warning)' : 'var(--success)',
            bg:   urgentes > 0 ? 'rgba(234,179,8,0.12)' : 'rgba(22,163,74,0.1)',
            valor: urgentes,
            label: 'adiantamentos vencem em até 7 dias'
        },
        {
            icon: 'ph-warning-circle',
            cor:  vencidos > 0 ? 'var(--danger)' : 'var(--success)',
            bg:   vencidos > 0 ? 'rgba(239,68,68,0.1)' : 'rgba(22,163,74,0.1)',
            valor: vencidos,
            label: 'adiantamentos vencidos'
        },
        {
            icon: 'ph-receipt',
            cor:  'var(--accent)',
            bg:   'rgba(2,132,199,0.1)',
            valor: totalAdi,
            label: 'adiantamentos em aberto no total'
        }
    ];

    // Adiciona card de cargo/permissões
    const permsLabel = usuarioAtual.permissoes.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' · ');
    itensCard.push({
        icon: 'ph-shield-check',
        cor:  '#f59e0b',
        bg:   'rgba(245,158,11,0.1)',
        valor: '',
        label: permsLabel,
        subtitulo: true
    });

    itensCard.forEach(c => {
        const div = document.createElement('div');
        div.className = 'bv-card';
        div.style.background = c.bg;
        div.style.borderColor = c.cor + '55';
        div.innerHTML = `
            <i class="ph ${c.icon}" style="font-size:28px;color:${c.cor};"></i>
            <div>
                ${c.subtitulo
                    ? `<p class="bv-card-label" style="color:${c.cor};font-size:13px;font-weight:800;">${c.label}</p>`
                    : `<p class="bv-card-valor" style="color:${c.cor};">${c.valor}</p>
                       <p class="bv-card-label">${c.label}</p>`
                }
            </div>`;
        cards.appendChild(div);
    });

    document.getElementById('modalBoasVindas').style.display = 'flex';
}

// --- EVENTOS DE TECLADO E MODAL ---
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('passInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('userInput').addEventListener('keydown', e => { if (e.key === 'Enter') realizarLogin(); });
    document.getElementById('confirmaSenha').addEventListener('keydown', e => { if (e.key === 'Enter') confirmarNovaSenha(); });

    // Fecha modais clicando fora
    ['searchModal', 'modalAlertaAdi', 'modalAlertaProj', 'modalConfirmaSaida', 'modalUsuario', 'modalDeletarUsuario', 'modalResetarSenha', 'modalBlacklist', 'modalEsqueciSenha', 'modalLimparLog', 'modalEditarNota', 'modalRetirada', 'modalSaldoSub'].forEach(id => {
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
        const res  = await fetch(addAuth(URL_SCRIPT));
        const data = await res.json();

        document.getElementById('setorMediaGeral').innerText = data.statsSetor.mediaGeral;
        document.getElementById('setorForn').innerText       = data.statsSetor.topForn;

        if (data.adiantamentosSetor && data.adiantamentosSetor.length > 0) {
            const hoje = new Date();
            hoje.setHours(0, 0, 0, 0);

            let adiantamentosParaExibir = data.adiantamentosSetor;
            if (!temPermissao('gestor') && !temPermissao('administrador') && !temPermissao('diretor')) {
                adiantamentosParaExibir = data.adiantamentosSetor.filter(adi => adi.responsavel === usuarioAtual.nome);
            }
            adiantamentosParaExibir.sort((a, b) => new Date(a.venc) - new Date(b.venc));

            // Salva no cache global
            adiantamentosCarregados = adiantamentosParaExibir;
            atualizarBadgeAdi(adiantamentosParaExibir);

            // ── ALERTA DE LOGIN: dispara apenas uma vez por sessão ──
            if (!alertaAdiJaExibido) {
                exibirAlertaAdiantamento(adiantamentosParaExibir);
            }
        } else {
            adiantamentosCarregados = [];
        }

        if (temPermissao('gestor') || temPermissao('administrador')) {
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

        await carregarProjecaoBackground();
        carregarPizzasDashboard();
    } catch (e) {
        document.getElementById('dash-loading').style.display = 'none';
        document.getElementById('dash-content').style.display = 'block';
        document.getElementById('dash-content').innerHTML =
            `<p style='color:var(--danger);padding:24px;'><i class='ph ph-warning'></i> Erro ao carregar dados do Google Sheets.</p>`;
    }
}

// =========================================================
// ABAS DE PROJEÇÃO PARA DIRETORES
// =========================================================
async function criarTabsDiretor() {
    try {
        const res   = await fetch(addAuth(`${URL_SCRIPT}?action=getCompradores&solicitante=${encodeURIComponent(loginAtual)}`));
        const lista = await res.json();

        const container = document.getElementById('navProjecoesDir');
        container.innerHTML = '';

        lista.forEach(u => {
            if (u.prefixos) {
                prefixosConfig[u.nomeUsuario] = u.prefixos.split('|').map(p => p.trim()).filter(Boolean);
            }
            if (u.nomeUsuario === usuarioAtual.nome.split(' ')[0].toUpperCase()) return;

            const abaId = `Projecao_${u.nomeUsuario}`;

            // Nav oculto
            const btn   = document.createElement('button');
            btn.className = 'nav-btn nav-btn-proj-dir';
            btn.id        = `btn-proj-${u.nomeUsuario.toLowerCase()}`;
            btn.innerHTML = `<i class="ph ph-package"></i> PROJ. ${u.nomeUsuario}`;
            btn.onclick   = () => switchTab(abaId);
            container.appendChild(btn);

            // Float nav
            const fnContainer = document.getElementById('fnProjecoesDir');
            if (fnContainer) {
                const fnBtn = document.createElement('button');
                fnBtn.className = 'fn-btn';
                fnBtn.id = `fn-proj-${u.nomeUsuario.toLowerCase()}`;
                fnBtn.setAttribute('data-tip', `Proj. ${u.nomeUsuario}`);
                fnBtn.innerHTML = `<i class="ph ph-package"></i>`;
                fnBtn.onclick = () => switchTab(abaId);
                fnContainer.appendChild(fnBtn);
            }
        });
    } catch (e) {
        console.warn('Erro ao criar abas de diretor:', e);
    }
}

const ADI_PAGE_SIZE = 15;
let _adiPagina = 1;

function renderizarTabelaAdiantamentos(lista) {
    const tbody    = document.querySelector("#tabelaMonitorAdi tbody");
    const paginacao = document.getElementById('adiPaginacao');
    tbody.innerHTML    = "";
    paginacao.innerHTML = "";

    if (!lista || lista.length === 0) {
        tbody.innerHTML = "<tr><td colspan='7' style='text-align:center; padding:20px;'>Nenhum adiantamento pendente.</td></tr>";
        return;
    }

    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const total     = lista.length;
    const totalPags = Math.ceil(total / ADI_PAGE_SIZE);
    const inicio    = (_adiPagina - 1) * ADI_PAGE_SIZE;
    const pagina    = lista.slice(inicio, Math.min(inicio + ADI_PAGE_SIZE, total));

    pagina.forEach(adi => {
        const vencDate = new Date(adi.venc);
        const diff = Math.ceil((vencDate - hoje) / (1000 * 60 * 60 * 24));
        let cls = "prazo-ok", txt = "No Prazo";
        if (diff < 0)       { cls = "prazo-vencido"; txt = "VENCIDO"; }
        else if (diff <= 7) { cls = "prazo-urgente"; txt = "URGENTE"; }

        tbody.innerHTML += `
            <tr id="row-adi-${adi.nf}">
                <td><b>${adi.responsavel}</b></td>
                <td>${adi.nf}</td>
                <td>${adi.fornecedor}</td>
                <td>${vencDate.toLocaleDateString('pt-BR', {timeZone: 'UTC'})}</td>
                <td>R$ ${formatarValor(adi.valor)}</td>
                <td><span class="status-prazo ${cls}">${txt}</span></td>
                <td style="text-align:center">
                    <button class="btn-saida-adi" title="Registrar saída"
                        onclick="deletarAdiantamento('${adi.nf}', '${adi.responsavel}')">
                        <i class="ph ph-door-open"></i>
                    </button>
                </td>
            </tr>`;
    });

    // Paginação
    if (totalPags > 1) {
        const btn = (label, p, ativo) =>
            `<button onclick="mudarPaginaAdi(${p})" style="padding:6px 14px;border-radius:8px;border:1.5px solid ${ativo ? 'var(--accent)' : 'var(--border)'};background:${ativo ? 'var(--accent)' : 'transparent'};color:${ativo ? '#fff' : 'var(--text-main)'};font-weight:800;font-size:12px;cursor:${ativo ? 'default' : 'pointer'}">${label}</button>`;
        if (_adiPagina > 1) paginacao.innerHTML += btn('‹', _adiPagina - 1, false);
        for (let p = Math.max(1, _adiPagina - 2); p <= Math.min(totalPags, _adiPagina + 2); p++) {
            paginacao.innerHTML += btn(p, p, p === _adiPagina);
        }
        if (_adiPagina < totalPags) paginacao.innerHTML += btn('›', _adiPagina + 1, false);

        paginacao.innerHTML += `<span style="font-size:12px;color:var(--text-muted);margin-left:8px;">${total} adiantamento(s)</span>`;
    }
}

function mudarPaginaAdi(p) {
    _adiPagina = p;
    renderizarTabelaAdiantamentos(adiantamentosCarregados);
    document.getElementById('monitorAdiantamentosSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
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
        const res  = await fetch(addAuth(URL_SCRIPT));
        const data = await res.json();

        let lista = data.adiantamentosSetor || [];
        if (!temPermissao('gestor') && !temPermissao('administrador') && !temPermissao('diretor')) {
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
        const res = await fetch(addAuth(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`));
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
let compradorCarregado = ''; // rastreia qual comprador está em cache
let filtroProjecaoAtual = 'todos';
let ordenacaoCobertura = 'asc';
let paginaAtualProjecao = 1;
const ITENS_POR_PAGINA = 100;
let itensFiltradosProjecao = [];

// Cache de adiantamentos (compartilhado entre dashboard e aba Adiantamento)
let adiantamentosCarregados = [];

// Ponto de compra (slider da projeção) — padrão 50 dias
let pontoCompraAtual = 50;

// Status de processo de compra por item
let statusCompraMap = {};          // { codigo: { status, dataStatus, coberturaRegistrada } }
let itemStatusCompraAtivo = null;  // item sendo editado no mini menu
let compradorProjecaoAtual = '';   // comprador da projeção sendo visualizada (ex: CRISLENE)

// Prefixos de código por usuário { 'CRIS': ['42','43'], 'ERNESTO': ['61','62',...] }
let prefixosConfig = {};
// Prefixos atualmente ativos na UI (Set)
let prefixosAtivos = new Set();

// Blacklist de códigos (Set) — carregada uma vez no login
let blacklistCodigos = new Set();

// --- CARRINHO DE COMPRAS ---
let carrinho = []; // { codigo, descricao, cobertura, statusTexto }

// --- CARREGAMENTO E FILTRO DE PROJEÇÃO ---

async function carregarProjecao(nomeUsuarioForcar) {
    const nomeUsuario = nomeUsuarioForcar || usuarioAtual.nome.split(' ')[0].toUpperCase();

    // Se já temos dados do mesmo comprador em cache, mostra sem novo fetch
    if (dadosProjecao.length > 0 && compradorCarregado === nomeUsuario) {
        filtroProjecaoAtual = 'todos';
        compradorProjecaoAtual = nomeUsuario;
        renderizarChipsPrefixos(prefixosAtivos);
        atualizarTabelaProjecao(aplicarFiltroPrefixo(dadosProjecao));
        desenharPizzaProjecao(dadosProjecao);
        renderizarGraficoProcesso();
        carregarNotificacoes();
        document.getElementById('proj-loading').style.display = 'none';
        document.getElementById('proj-content').style.display = 'block';
        return;
    }

    document.getElementById('proj-loading').style.display = 'flex';
    document.getElementById('proj-content').style.display = 'none';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`));
        const data = await res.json();

        if (data.erro) {
            document.getElementById('proj-loading').innerHTML =
                `<p style='color:var(--danger)'>Erro: ${data.erro}</p>`;
            return;
        }

        dadosProjecao  = data.itens || [];
        filtroProjecaoAtual = 'todos';
        compradorCarregado = nomeUsuario;
        // Reseta o slider para o padrão ao carregar nova projeção
        pontoCompraAtual = 50;
        const slider = document.getElementById('pontoCompraSlider');
        const label  = document.getElementById('pontoCompraLabel');
        if (slider) slider.value = 50;
        if (label)  label.textContent = '50 dias';

        // Carrega status de processo de compra para este comprador
        compradorProjecaoAtual = nomeUsuario;
        await carregarStatusCompra(nomeUsuario);

        renderizarChipsPrefixos(prefixosAtivos);
        atualizarTabelaProjecao(aplicarFiltroPrefixo(dadosProjecao));
        desenharPizzaProjecao(dadosProjecao);
        renderizarGraficoProcesso();
        carregarNotificacoes();

        document.getElementById('proj-loading').style.display = 'none';
        document.getElementById('proj-content').style.display = 'block';
    } catch (e) {
        console.error("Erro carregarProjecao:", e);
        document.getElementById('proj-loading').innerHTML =
            `<p style='color:var(--danger)'>Erro ao carregar dados de projeção.<br><small style='opacity:0.7'>${e.message || e}</small></p>`;
    }
}

// --- FILTRO DE PREFIXOS ---

function aplicarFiltroPrefixo(itens) {
    // Primeiro aplica blacklist global
    let resultado = blacklistCodigos.size > 0
        ? itens.filter(item => !blacklistCodigos.has(item.codigo.trim().toUpperCase()))
        : itens;
    // Depois aplica filtro de prefixos do usuário
    if (prefixosAtivos.size === 0) return resultado;
    return resultado.filter(item =>
        [...prefixosAtivos].some(p => item.codigo.startsWith(p))
    );
}

function renderizarChipsPrefixos(ativos) {
    const wrap  = document.getElementById('prefixosWrap');
    const chips = document.getElementById('prefixosChips');

    // Descobre todos os prefixos disponíveis para o usuário atual
    const todosOsPrefixos = ativos.size > 0 ? [...ativos] : [];
    // Também inclui prefixos que estavam originalmente configurados
    const configKey = abaAtual.startsWith('Projecao_')
        ? abaAtual.replace('Projecao_', '')
        : usuarioAtual.nome.split(' ')[0].toUpperCase();
    const prefixosOriginais = prefixosConfig[configKey] || usuarioAtual.prefixos || [];

    if (prefixosOriginais.length === 0) {
        wrap.style.display = 'none';
        return;
    }

    wrap.style.display = 'block';
    chips.innerHTML = '';
    prefixosOriginais.forEach(p => {
        const ativo = prefixosAtivos.has(p);
        const chip  = document.createElement('button');
        chip.className = `chip-prefixo ${ativo ? 'ativo' : ''}`;
        chip.textContent = p;
        chip.title = ativo ? `Ocultar itens ${p}xx` : `Mostrar itens ${p}xx`;
        chip.onclick = () => togglePrefixo(p);
        chips.appendChild(chip);
    });
}

function togglePrefixo(p) {
    if (prefixosAtivos.has(p)) {
        prefixosAtivos.delete(p);
    } else {
        prefixosAtivos.add(p);
    }
    renderizarChipsPrefixos(prefixosAtivos);
    aplicarFiltroAtual();
}

function toggleTodosPrefixos(ligar) {
    const configKey = abaAtual.startsWith('Projecao_')
        ? abaAtual.replace('Projecao_', '')
        : usuarioAtual.nome.split(' ')[0].toUpperCase();
    const prefixosOriginais = prefixosConfig[configKey] || usuarioAtual.prefixos || [];

    if (ligar) {
        prefixosAtivos = new Set(prefixosOriginais);
    } else {
        prefixosAtivos = new Set();
    }
    renderizarChipsPrefixos(prefixosAtivos);
    aplicarFiltroAtual();
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
    } else if (filtroProjecaoAtual === 'zeradoCobertura') {
        itensParaExibir = dadosProjecao.filter(i => i.cobertura <= 0);
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
    // Aplica o ponto de compra primeiro
    let filtrado = aplicarFiltroPrefixo(dadosProjecao).filter(i => i.cobertura <= pontoCompraAtual);

    if (tipo === 'rp') {
        filtrado = dadosProjecao.filter(i => i.temRP);
    } else if (tipo === 'cd') {
        filtrado = dadosProjecao.filter(i => i.saldoCD > 0);
    } else if (tipo === 'empenho') {
        filtrado = dadosProjecao.filter(i => i.temEmpenho);
    } else if (tipo === 'zerado') {
        filtrado = dadosProjecao.filter(i => i.zeradoSemCobertura);
    } else if (tipo === 'zeradoCobertura') {
        filtrado = dadosProjecao.filter(i => i.cobertura <= 0);
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
            if (item.cobertura === 0) {
                statusBadges += '<span class="badge-status zerado"><i class="ph ph-warning-diamond"></i> CRÍTICO</span>';
            } else {
                statusBadges += '<span class="badge-status comprar"><i class="ph ph-shopping-cart"></i> COMPRAR</span>';
            }
            temAlgumStatus = true;
        }

        if (!temAlgumStatus) {
            statusBadges = '<span class="badge-status comprar"><i class="ph ph-shopping-cart"></i> COMPRAR</span>';
        }

        // Texto limpo do status para o carrinho/CSV
        let statusTexto = 'COMPRAR';
        if (item.zeradoSemCobertura) {
            statusTexto = item.cobertura === 0 ? 'CRÍTICO' : 'COMPRAR';
        } else if (item.temRP)       statusTexto = 'RP';
        else if (item.saldoCD > 0)   statusTexto = 'CD';
        else if (item.temEmpenho)    statusTexto = 'EMPENHO';

        tbody.innerHTML += `
            <tr>
                <td><b>${item.codigo}</b></td>
                <td>${item.descricao}</td>
                <td>${item.cobertura || 0}</td>
                <td>${item.saldoCD || 0}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoDiario > 0 ? item.consumoDiario.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 2}) : '—'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoMensal > 0 ? item.consumoMensal.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 2}) : '—'}</td>
                <td>${statusBadges}</td>
                <td style="text-align:center;">
                    <button onclick="abrirSaldoSub('${item.codigo}', ${item.consumoDiario || 0})"
                        style="background:transparent;border:1.5px solid var(--border);border-radius:8px;padding:5px 7px;cursor:pointer;color:var(--text-muted);" title="Ver saldo nos subestoques">
                        <i class="ph ph-warehouse"></i>
                    </button>
                </td>
                <td style="text-align:center;">
                    ${(() => {
                        const sc = statusCompraMap[item.codigo];
                        const faseIcon = {
                            'EM ORÇAMENTO':                       'ph-magnifying-glass',
                            'REQUISIÇÃO DE COMPRA':               'ph-file-text',
                            'ORDEM DE COMPRA':                    'ph-receipt',
                            'AGUARDANDO ENTREGA DO FORNECEDOR':   'ph-truck'
                        };
                        const faseColor = {
                            'EM ORÇAMENTO':                       '#f59e0b',
                            'REQUISIÇÃO DE COMPRA':               '#3b82f6',
                            'ORDEM DE COMPRA':                    '#8b5cf6',
                            'AGUARDANDO ENTREGA DO FORNECEDOR':   '#10b981'
                        };
                        if (sc) {
                            const cor   = faseColor[sc.status] || 'var(--accent)';
                            const icone = faseIcon[sc.status]  || 'ph-clock';
                            return `<button class="btn-status-compra ativo"
                                data-codigo="${item.codigo}"
                                data-descricao="${item.descricao.replace(/"/g,'&quot;')}"
                                data-cobertura="${item.cobertura || 0}"
                                style="border-color:${cor};color:${cor};"
                                title="${sc.status}"
                                onclick="abrirMenuStatusCompra(this)">
                                <i class="ph ${icone}"></i>
                            </button>`;
                        }
                        return `<button class="btn-status-compra"
                            data-codigo="${item.codigo}"
                            data-descricao="${item.descricao.replace(/"/g,'&quot;')}"
                            data-cobertura="${item.cobertura || 0}"
                            title="Definir status do processo de compra"
                            onclick="abrirMenuStatusCompra(this)">
                            <i class="ph ph-clock-countdown"></i>
                        </button>`;
                    })()}
                </td>
                <td style="text-align:center;">
                    <button class="btn-add-carrinho ${carrinho.some(c => c.codigo === item.codigo) ? 'no-carrinho' : ''}"
                        data-codigo="${item.codigo}"
                        data-descricao="${item.descricao.replace(/"/g, '&quot;')}"
                        data-cobertura="${item.cobertura || 0}"
                        data-status="${statusTexto}"
                        data-consumo-diario="${item.consumoDiario || 0}"
                        data-consumo-mensal="${item.consumoMensal || 0}"
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
    document.querySelectorAll('.fn-btn').forEach(b => b.classList.remove('active'));
    // Atualiza active no drawer mobile
    document.querySelectorAll('.mobile-nav-link').forEach(b => b.classList.remove('active'));
    // Esconde views de protocolos sempre que trocar de aba
    const vp  = document.getElementById('view-protocolos');
    const vps = document.getElementById('view-protocolos-sup');
    const vao = document.getElementById('view-admin-opme');
    const vpro = document.getElementById('view-projecao-opme');
    if (vp)   vp.style.display   = 'none';
    if (vps)  vps.style.display  = 'none';
    if (vao)  vao.style.display  = 'none';
    if (vpro) vpro.style.display = 'none';

    // Ativa botão no nav oculto E no float nav
    const mapaFn = {
        'Dashboard': 'fn-dashboard', 'Digitadas': 'fn-digitadas',
        'Recebimento': 'fn-recebimento', 'Adiantamento': 'fn-adiantamento',
        'Admin': 'fn-admin', 'AdminOpme': 'fn-admin-opme',
        'ProtocolosOpme': 'fn-protocolos-opme',
        'ProtocolosSup': 'fn-protocolos-sup',
        'ProjecaoOPME':  'fn-projecao-opme'
    };
    if (aba.startsWith('Projecao_')) {
        const nome = aba.replace('Projecao_', '').toLowerCase();
        const btnDir = document.getElementById(`btn-proj-${nome}`);
        const fnDir  = document.getElementById(`fn-proj-${nome}`);
        if (btnDir) btnDir.classList.add('active');
        if (fnDir)  fnDir.classList.add('active');
        else { const fp = document.getElementById('fn-projecao'); if (fp) fp.classList.add('active'); }
    } else {
        const btnEl = document.getElementById('btn-' + aba.toLowerCase());
        if (btnEl) btnEl.classList.add('active');
        const fnId = mapaFn[aba];
        if (fnId) { const fn = document.getElementById(fnId); if (fn) fn.classList.add('active'); }
    }
    setTimeout(atualizarNavIndicador, 10);

    if (aba.startsWith('Projecao_')) {
        const nomeUsuario = aba.replace('Projecao_', '').toLowerCase();
        const btnDir = document.getElementById(`btn-proj-${nomeUsuario}`);
        if (btnDir) btnDir.classList.add('active');
    } else {
        const btnEl = document.getElementById('btn-' + aba.toLowerCase());
        if (btnEl) btnEl.classList.add('active');
    }

    // Atualiza o indicador deslizante
    setTimeout(atualizarNavIndicador, 10);

    const monitorSec = document.getElementById('monitorAdiantamentosSection');
    if (monitorSec) monitorSec.style.display = 'none';

    // Em modo OPME só permite ProtocolosOpme
    if (modoAtual === 'opme' && aba !== 'ProtocolosOpme' && aba !== 'AdminOpme' && aba !== 'ProjecaoOPME') {
        switchTab('ProtocolosOpme');
        return;
    }

    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'none';
        const isGestorOuAdmin = temPermissao('gestor') || temPermissao('administrador');
        document.getElementById('dash-gestor').style.display = isGestorOuAdmin ? 'block' : 'none';
        if (!_cache.dashboard) { carregarEstatisticas(); _cache.dashboard = true; }
    } else if (aba === 'ProtocolosOpme') {
        document.getElementById('view-dashboard').style.display  = 'none';
        document.getElementById('view-forms').style.display      = 'none';
        document.getElementById('view-projecao').style.display   = 'none';
        document.getElementById('view-admin').style.display      = 'none';
        const vp = document.getElementById('view-protocolos');
        if (vp) vp.style.display = 'block';
        if (!_cache.protocolosOpme) { carregarMeusProtocolos(); _cache.protocolosOpme = true; }
    } else if (aba === 'ProtocolosSup') {
        document.getElementById('view-dashboard').style.display  = 'none';
        document.getElementById('view-forms').style.display      = 'none';
        document.getElementById('view-projecao').style.display   = 'none';
        document.getElementById('view-admin').style.display      = 'none';
        const vps = document.getElementById('view-protocolos-sup');
        if (vps) vps.style.display = 'block';
        if (!_cache.protocolosSup) { carregarProtocolosSup(); _cache.protocolosSup = true; }
    } else if (aba === 'Projecao' || aba.startsWith('Projecao_')) {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'block';
        document.getElementById('view-admin').style.display = 'none';

        if (aba.startsWith('Projecao_')) {
            const nomeUsuario = aba.replace('Projecao_', '');
            document.getElementById('projTitulo').innerText = `Projeção de Compras — ${nomeUsuario.charAt(0) + nomeUsuario.slice(1).toLowerCase()}`;
            prefixosAtivos = new Set(prefixosConfig[nomeUsuario] || []);
            carregarProjecao(nomeUsuario); // já tem cache interno via dadosProjecao.length
        } else {
            const nomeUsuario = usuarioAtual.nome.split(' ')[0];
            document.getElementById('projTitulo').innerText = `Projeção de Compras — ${nomeUsuario}`;
            prefixosAtivos = new Set(usuarioAtual.prefixos || []);
            carregarProjecao();
        }
    } else if (aba === 'ProjecaoOPME') {
        document.getElementById('view-dashboard').style.display   = 'none';
        document.getElementById('view-forms').style.display       = 'none';
        document.getElementById('view-projecao').style.display    = 'none';
        document.getElementById('view-admin').style.display       = 'none';
        const _vpo = document.getElementById('view-projecao-opme');
        if (_vpo) _vpo.style.display = 'block';
        // Se ainda não rodou a projeção nesta sessão, mostra estado inicial
        // (nada a carregar automaticamente — usuário clica em "Rodar Projeção")
    } else if (aba === 'AdminOpme') {
        document.getElementById('view-dashboard').style.display   = 'none';
        document.getElementById('view-forms').style.display       = 'none';
        document.getElementById('view-projecao').style.display    = 'none';
        document.getElementById('view-admin').style.display       = 'none';
        const vao = document.getElementById('view-admin-opme');
        if (vao) vao.style.display = 'block';
        const vp = document.getElementById('view-protocolos');
        if (vp) vp.style.display = 'none';
        carregarUsuariosOpme();
    } else if (aba === 'Admin') {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'block';
        if (!_cache.usuarios) { carregarUsuarios(); _cache.usuarios = true; }
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
        if (aba === 'Adiantamento') {
            document.getElementById('historicoNotasSection').style.display = 'none';
            if (!_cache.adiantamento) { carregarAdiantamentos(); _cache.adiantamento = true; }
        } else if (aba === 'Digitadas' || aba === 'Recebimento') {
            if (!_cache.historico) { setTimeout(() => carregarHistoricoNotas(), 200); _cache.historico = true; }
        } else {
            document.getElementById('historicoNotasSection').style.display = 'none';
        }
    }
}

// --- NAVEGAÇÃO POR SETAS NO FORMULÁRIO ---
// Ordem: f_nf(0) → f_data(1) → f_fornecedor(2) → f_razao(3) → f_vencimento(4) → f_setor(5) → f_lote(6)
const _NAV_IDS = ['f_nf','f_data','f_fornecedor','f_razao','f_vencimento','f_setor','f_lote'];

function navForm(e, el) {
    const isSelect = el.tagName === 'SELECT';

    // Não interfere no autocomplete de fornecedor se a lista estiver aberta
    const acList = document.getElementById('ac-list');
    if (acList && acList.style.display !== 'none' &&
        (e.key === 'ArrowDown' || e.key === 'ArrowUp')) return;

    // Setas esquerda/direita sempre navegam entre campos
    // Seta cima/baixo também navega (exceto em select, onde muda opção)
    let irProximo = false, irAnterior = false;

    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
        if (isSelect && e.key === 'ArrowDown') return; // deixa mudar opção
        irProximo = true;
    } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
        if (isSelect && e.key === 'ArrowUp') return; // deixa mudar opção
        irAnterior = true;
    } else if (e.key === 'Enter') {
        return; // Enter confirma o formulário normalmente
    } else {
        return;
    }

    e.preventDefault();
    const idx = parseInt(el.getAttribute('data-nav') ?? '-1');
    if (idx < 0) return;

    const nextIdx = irProximo ? idx + 1 : idx - 1;
    const nextId  = _NAV_IDS[nextIdx];
    if (nextId) {
        const next = document.getElementById(nextId);
        if (next) {
            next.focus();
            if (next.select) next.select();
        }
    }
}

function limparFormulario() {
    document.getElementById('f_nf').value = '';
    document.getElementById('f_fornecedor').value = '';
    document.getElementById('f_razao').value = 'HC';
    document.getElementById('f_vencimento').value = '';
    document.getElementById('f_setor').value = '';
    document.getElementById('f_lote').value = 'Sim';
    // Desmarca todos os radios
    document.querySelectorAll('#notaForm input[type="radio"]').forEach(r => r.checked = false);
    document.getElementById('f_nf').focus();
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
                    <label class="radio-item"><input type="radio" name="gSis" value="Hosplog"> HOSPLOG</label>
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
    head.innerHTML = `<th>NF</th><th>FORNECEDOR</th><th>RAZÃO</th><th>VENC.</th><th>SETOR</th><th>AÇÃO</th>`;
}

function adicionarNota() {
    const nf = document.getElementById('f_nf').value;
    if (!nf) return alert("Número da NF é obrigatório!");

    // Data de digitação — digitada pelo usuário
    const _dataInput = document.getElementById('f_data').value;
    if (!_dataInput) return alert("Informe a data de digitação!");
    const _dataHoje = _dataInput; // já está no formato YYYY-MM-DD

    // Vencimento direto do input date (YYYY-MM-DD)
    const _vencFormatado = document.getElementById('f_vencimento').value;

    const nota = {
        destino:         abaAtual,
        responsavel:     usuarioAtual.nome,
        data:            _dataHoje,
        nf:              nf,
        fornecedor:      document.getElementById('f_fornecedor').value,
        razaoSocial:     document.getElementById('f_razao').value,
        vencimento:      _vencFormatado,
        valor:           0,
        setor:           document.getElementById('f_setor').value || "GERAL",
        possuiLote:      document.getElementById('f_lote').value,
        numAdiantamento: document.getElementById('f_num_adi').value,
        statusDigitacao: abaAtual === 'Digitadas'
            ? (document.querySelector('input[name="gSis"]:checked')?.value || '')
            : ""
    };

    listas[abaAtual].push(nota);
    atualizarTabela();
    document.getElementById('f_nf').value = "";
    document.getElementById('f_vencimento').value = "";
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
                <td>${n.setor || '—'}</td>
                <td><button onclick="removerNota(${i})" style="border:none;background:none;cursor:pointer;color:var(--danger)"><i class="ph ph-trash" style="font-size:20px"></i></button></td>
            </tr>`;
    });
    document.getElementById('areaAcoes').style.display = listas[abaAtual].length > 0 ? 'grid' : 'none';
}

function removerNota(i) {
    listas[abaAtual].splice(i, 1);
    atualizarTabela();
}

// --- EXPORTAR LOTE CSV COM CORES ---

function abrirModalCSVLote() {
    fecharModal('modalCSVLote');
    document.getElementById('modalCSVLote').style.display = 'flex';

    const inpFundo = document.getElementById('csvCorFundo');
    const inpFonte = document.getElementById('csvCorFonte');

    function atualizarPrevia() {
        const fundo = inpFundo.value;
        const fonte = inpFonte.value;
        document.getElementById('csvPrevia').style.background = fundo;
        document.getElementById('csvCorFundoLabel').textContent = fundo;
        document.getElementById('csvCorFonteLabel').textContent = fonte;
        ['csvPrevCol1','csvPrevCol2','csvPrevCol3','csvPrevCol4','csvPrevCol5'].forEach(id => {
            document.getElementById(id).style.color = fonte;
        });
    }

    inpFundo.oninput = atualizarPrevia;
    inpFonte.oninput = atualizarPrevia;
    atualizarPrevia();
}

function confirmarExportarCSVLote() {
    const corFundo = document.getElementById('csvCorFundo').value;
    const corFonte = document.getElementById('csvCorFonte').value;
    const notas = listas[abaAtual] || [];
    if (!notas.length) { mostrarToast('Nenhuma nota no lote.', 'error'); return; }

    try {
        const wb = XLSX.utils.book_new();
        const cabecalho = ['NF', 'FORNECEDOR', 'RAZÃO SOCIAL', 'VENCIMENTO', 'SETOR'];
        const linhas = notas.map(n => [
            n.nf,
            n.fornecedor,
            n.razaoSocial,
            formatarDataBR(n.vencimento),
            n.setor || ''
        ]);

        const ws = XLSX.utils.aoa_to_sheet([cabecalho, ...linhas]);

        // Larguras das colunas
        ws['!cols'] = [{ wch: 14 }, { wch: 35 }, { wch: 18 }, { wch: 14 }, { wch: 22 }];

        // Estilo do cabeçalho
        const hexParaARGB = hex => 'FF' + hex.replace('#','').toUpperCase();
        const estiloHdr = {
            font: { bold: true, color: { rgb: hexParaARGB(corFonte) }, sz: 11 },
            fill: { fgColor: { rgb: hexParaARGB(corFundo) }, patternType: 'solid' },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
                bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
                right:  { style: 'thin', color: { rgb: 'CCCCCC' } }
            }
        };
        cabecalho.forEach((_, ci) => {
            const cell = XLSX.utils.encode_cell({ r: 0, c: ci });
            if (ws[cell]) ws[cell].s = estiloHdr;
        });

        // Estilo das linhas de dados (zebra)
        linhas.forEach((_, ri) => {
            const bgZebra = ri % 2 === 0 ? 'FFFFFF' : 'F8FAFC';
            cabecalho.forEach((_, ci) => {
                const cell = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
                if (!ws[cell]) ws[cell] = { v: '', t: 's' };
                ws[cell].s = {
                    fill: { fgColor: { rgb: bgZebra }, patternType: 'solid' },
                    alignment: { vertical: 'center' },
                    border: {
                        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                        right:  { style: 'thin', color: { rgb: 'E2E8F0' } }
                    }
                };
            });
        });

        XLSX.utils.book_append_sheet(wb, ws, 'Lote');
        const hoje = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
        XLSX.writeFile(wb, `lote_${abaAtual.toLowerCase()}_${hoje}.xlsx`);
        fecharModal('modalCSVLote');
        mostrarToast('✅ Arquivo exportado com sucesso!', 'success');
        tocarSomMSN();
    } catch(e) {
        console.error(e);
        mostrarToast('Erro ao gerar arquivo.', 'error');
    }
}



function enviarTudo() {
    if (!listas[abaAtual] || !listas[abaAtual].length) {
        mostrarToast('Nenhuma nota no lote.', 'error'); return;
    }
    // Resetar para etapa 1
    document.getElementById('envioStep1').style.display = 'block';
    document.getElementById('envioStep2').style.display = 'none';
    document.getElementById('modalConfirmaEnvio').style.display = 'flex';
}

function abrirSelecaoNotasErro() {
    const lista = document.getElementById('envioNotasLista');
    lista.innerHTML = '';
    listas[abaAtual].forEach((n, i) => {
        lista.innerHTML += `
            <label style="display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:8px;border:1.5px solid var(--border);cursor:pointer;transition:all .15s;"
                onmouseover="this.style.borderColor='#f59e0b'" onmouseout="this.style.borderColor='var(--border)'">
                <input type="checkbox" data-idx="${i}" value="${i}"
                    style="width:16px;height:16px;accent-color:#f59e0b;cursor:pointer;flex-shrink:0;">
                <div>
                    <span style="font-weight:800;font-size:13px;">NF ${n.nf}</span>
                    <span style="font-size:11px;color:var(--text-muted);margin-left:8px;">${n.fornecedor} — ${n.razaoSocial}</span>
                </div>
            </label>`;
    });
    document.getElementById('envioStep1').style.display = 'none';
    document.getElementById('envioStep2').style.display = 'block';
}

async function confirmarEnvioComErrosSelecionados() {
    const checks = document.querySelectorAll('#envioNotasLista input[type="checkbox"]:checked');
    const idxsComErro = new Set([...checks].map(c => parseInt(c.value)));
    fecharModal('modalConfirmaEnvio');

    listas[abaAtual].forEach((n, i) => {
        if (idxsComErro.has(i)) {
            n.statusDigitacao = (n.statusDigitacao ? n.statusDigitacao + ' ' : '') + '(erro de reprocessamento)';
        }
    });

    await _executarEnvio();
}

async function confirmarEnvioRelatorio(notasComErro) {
    fecharModal('modalConfirmaEnvio');
    await _executarEnvio();
}

async function _executarEnvio() {
    const btn = document.getElementById('btnEnviar');
    btn.disabled = true;
    btn.innerText = "ENVIANDO...";
    const total = listas[abaAtual].length;
    try {
        await fetch(URL_SCRIPT, { method: 'POST', mode: 'no-cors', body: JSON.stringify(listas[abaAtual]) });
        mostrarToast(`✅ ${total} nota(s) enviadas para a planilha!`, 'success');
        registrarLog('SYNC NOTAS', `${total} nota(s) enviadas — aba ${abaAtual}`);
        listas[abaAtual] = [];
        atualizarTabela();
    } catch (e) {
        mostrarToast('Erro ao conectar com o Google Sheets.', 'error');
    } finally {
        btn.disabled = false;
        btn.innerText = "🗒️ ENVIAR RELATÓRIO DE NOTAS DIGITADAS";
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
        const res     = await fetch(addAuth(`${URL_SCRIPT}?search=${encodeURIComponent(q)}&tab=${abaAtual}`));
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
    const codigo          = btn.dataset.codigo;
    const descricao       = btn.dataset.descricao;
    const cobertura       = parseFloat(btn.dataset.cobertura) || 0;
    const statusTexto     = btn.dataset.status;
    const consumoDiario   = parseFloat(btn.dataset.consumoDiario) || 0;
    const consumoMensal   = parseFloat(btn.dataset.consumoMensal) || 0;

    const idx = carrinho.findIndex(i => i.codigo === codigo);
    if (idx >= 0) {
        carrinho.splice(idx, 1);
    } else {
        carrinho.push({ codigo, descricao, cobertura, statusTexto, consumoDiario, consumoMensal });
    }
    atualizarContadorCarrinho();
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
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoDiario > 0 ? item.consumoDiario.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 2}) : '—'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoMensal > 0 ? item.consumoMensal.toLocaleString('pt-BR', {minimumFractionDigits: 0, maximumFractionDigits: 2}) : '—'}</td>
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
        codigo:         item.codigo,
        descricao:      item.descricao,
        cobertura:      item.cobertura,
        status:         item.statusTexto,
        responsavel:    usuarioAtual.nome,
        observacao:     obs,
        data:           new Date().toLocaleDateString('pt-BR'),
        aba:            abaDestino,
        consumoDiario:  item.consumoDiario || 0,
        consumoMensal:  item.consumoMensal || 0
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
        registrarLog('SAÍDA ADIANTAMENTO', `NF ${nf} — ${responsavel}`);
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
        aplicarFiltroAtual();
        return;
    }

    // Parte sempre da base já filtrada por blacklist + prefixos
    let base = aplicarFiltroPrefixo(dadosProjecao).filter(i =>
        i.codigo.toLowerCase().includes(termo) ||
        i.descricao.toLowerCase().includes(termo)
    );

    // Aplica o filtro de status por cima da busca
    if (filtroProjecaoAtual === 'rp')           base = base.filter(i => i.temRP);
    else if (filtroProjecaoAtual === 'cd')       base = base.filter(i => i.saldoCD > 0);
    else if (filtroProjecaoAtual === 'empenho')  base = base.filter(i => i.temEmpenho);
    else if (filtroProjecaoAtual === 'zerado')          base = base.filter(i => i.zeradoSemCobertura);
    else if (filtroProjecaoAtual === 'zeradoCobertura') base = base.filter(i => i.cobertura <= 0);

    atualizarTabelaProjecao(base);
}

function limparBuscaProjecao() {
    document.getElementById('inputBuscaProjecao').value = '';
    document.getElementById('btnLimparBuscaProj').style.display = 'none';
    aplicarFiltroAtual();
}

function atualizarPontoCompra(valor) {
    pontoCompraAtual = parseInt(valor);
    document.getElementById('pontoCompraLabel').textContent = valor + ' dias';
    aplicarFiltroAtual();
}

// =============================================================================
// STATUS DE PROCESSO DE COMPRA
// =============================================================================

async function carregarStatusCompra(comprador) {
    try {
        const param = comprador ? `&comprador=${encodeURIComponent(comprador)}` : '';
        const resp = await fetch(addAuth(`${URL_SCRIPT}?action=getStatusCompra${param}&t=${Date.now()}`));
        const data = await resp.json();
        statusCompraMap = {};
        (data.itens || []).forEach(i => {
            if (!comprador || i.comprador.toUpperCase() === comprador.toUpperCase()) {
                statusCompraMap[i.codigo] = i;
            }
        });
    } catch(e) { console.warn("Erro ao carregar statusCompra:", e); }
}

function abrirMenuStatusCompra(btn) {
    const codigo    = btn.dataset.codigo;
    const descricao = btn.dataset.descricao;
    const cobertura = btn.dataset.cobertura;
    itemStatusCompraAtivo = { codigo, descricao, cobertura: parseFloat(cobertura) || 0 };

    document.getElementById('statusCompraMenuTitulo').textContent = descricao.substring(0, 40) + (descricao.length > 40 ? '…' : '');

    // Destaca fase atual se houver
    const sc = statusCompraMap[codigo];
    document.querySelectorAll('.status-opcao').forEach(b => b.classList.remove('ativo'));
    if (sc) {
        document.querySelectorAll('.status-opcao').forEach(b => {
            if (b.textContent.trim().includes(sc.status)) b.classList.add('ativo');
        });
    }

    document.getElementById('statusCompraOverlay').style.display = 'block';
    document.getElementById('statusCompraMenu').style.display    = 'block';
}

function fecharMenuStatusCompra() {
    document.getElementById('statusCompraOverlay').style.display = 'none';
    document.getElementById('statusCompraMenu').style.display    = 'none';
    itemStatusCompraAtivo = null;
}

async function definirStatusCompra(status) {
    if (!itemStatusCompraAtivo) return;
    const { codigo, descricao, cobertura } = itemStatusCompraAtivo;
    const comprador = compradorProjecaoAtual || usuarioAtual.nome.split(" ")[0].toUpperCase();

    fecharMenuStatusCompra();

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'salvarStatusCompra', codigo, descricao, comprador, status, cobertura })
        });
        statusCompraMap[codigo] = { codigo, descricao, comprador, status, coberturaRegistrada: cobertura };
        atualizarTabelaProjecao(itensFiltradosProjecao);
        renderizarGraficoProcesso();
        mostrarToast(`✅ Status definido: ${status}`, 'success');
    } catch(e) { mostrarToast('Erro ao salvar status', 'error'); }
}

async function removerStatusCompra() {
    if (!itemStatusCompraAtivo) return;
    const { codigo } = itemStatusCompraAtivo;
    const comprador  = compradorProjecaoAtual || usuarioAtual.nome.split(" ")[0].toUpperCase();

    fecharMenuStatusCompra();

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'removerStatusCompra', codigo, comprador })
        });
        delete statusCompraMap[codigo];
        atualizarTabelaProjecao(itensFiltradosProjecao);
        renderizarGraficoProcesso();
        mostrarToast('Status removido', 'info');
    } catch(e) { mostrarToast('Erro ao remover status', 'error'); }
}

async function marcarEntregue() {
    if (!itemStatusCompraAtivo) return;
    const { codigo } = itemStatusCompraAtivo;
    const comprador  = compradorProjecaoAtual || usuarioAtual.nome.split(" ")[0].toUpperCase();

    fecharMenuStatusCompra();

    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'removerStatusCompra', codigo, comprador })
        });
        delete statusCompraMap[codigo];
        // Remove da projeção localmente até a cobertura cair novamente
        dadosProjecao = dadosProjecao.filter(i => i.codigo !== codigo);
        aplicarFiltroAtual();
        renderizarGraficoProcesso();
        mostrarToast('✅ Item marcado como entregue e removido da projeção', 'success');
    } catch(e) { mostrarToast('Erro ao marcar entrega', 'error'); }
}

// =============================================================================
// GRÁFICO DE PROCESSO DE COMPRA
// =============================================================================

function renderizarGraficoProcesso() {
    const canvas  = document.getElementById('projProcessoChart');
    const legenda = document.getElementById('projProcessoLegenda');
    if (!canvas || !legenda) return;

    const fases = [
        { label: 'Sem processo',           cor: 'rgba(100,116,139,0.5)' },
        { label: 'Em Orçamento',           cor: '#f59e0b' },
        { label: 'Req. de Compra',         cor: '#3b82f6' },
        { label: 'Ordem de Compra',        cor: '#8b5cf6' },
        { label: 'Ag. Entrega Forn.',      cor: '#10b981' },
    ];

    const total = dadosProjecao.length || 1;
    const contagem = {
        'EM ORÇAMENTO': 0, 'REQUISIÇÃO DE COMPRA': 0,
        'ORDEM DE COMPRA': 0, 'AGUARDANDO ENTREGA DO FORNECEDOR': 0
    };
    Object.values(statusCompraMap).forEach(sc => {
        if (contagem[sc.status] !== undefined) contagem[sc.status]++;
    });
    const comProcesso = Object.values(contagem).reduce((a, b) => a + b, 0);
    const semProcesso = Math.max(0, (dadosProjecao.length || 0) - comProcesso);

    const valores = [
        semProcesso,
        contagem['EM ORÇAMENTO'],
        contagem['REQUISIÇÃO DE COMPRA'],
        contagem['ORDEM DE COMPRA'],
        contagem['AGUARDANDO ENTREGA DO FORNECEDOR']
    ];

    const ctx = canvas.getContext('2d');
    const cx  = canvas.width  / 2;
    const cy  = canvas.height / 2;
    const r   = Math.min(cx, cy) - 8;
    let angulo = -Math.PI / 2;

    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const fatias = fases.map((f, i) => ({ ...f, valor: valores[i] })).filter(f => f.valor > 0);
    const totalValores = fatias.reduce((a, f) => a + f.valor, 0) || 1;

    fatias.forEach(f => {
        const angFim = angulo + (f.valor / totalValores) * 2 * Math.PI;
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, angulo, angFim);
        ctx.closePath();
        ctx.fillStyle = f.cor;
        ctx.fill();
        angulo = angFim;
    });

    // Furo central (donut)
    ctx.beginPath();
    ctx.arc(cx, cy, r * 0.52, 0, 2 * Math.PI);
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--bg-card').trim() || '#1e1e2e';
    ctx.fill();

    // Centro: total em processo
    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 20px Calibri, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(comProcesso, cx, cy - 7);
    ctx.font = 'bold 8px Calibri, sans-serif';
    ctx.fillStyle = 'rgba(255,255,255,0.6)';
    ctx.fillText('em processo', cx, cy + 9);

    legenda.innerHTML = fatias.map(f => `
        <div style="display:flex;align-items:center;gap:8px;font-size:12px;">
            <span style="width:10px;height:10px;border-radius:50%;background:${f.cor};flex-shrink:0;display:inline-block;"></span>
            <span style="color:var(--text-muted)">${f.label}</span>
            <b style="color:var(--text-main);margin-left:auto;padding-left:8px;">${f.valor}</b>
            <span style="color:var(--text-muted);font-size:11px;">(${Math.round(f.valor/totalValores*100)}%)</span>
        </div>`).join('');
}

// =============================================================================
// NOTIFICAÇÕES INTERNAS
// =============================================================================

let notificacoesPendentes = [];

async function carregarNotificacoes() {
    // Admin vê todas as notificações; comprador só vê as suas
    const isAdmin = temPermissao('administrador');
    const comprador = isAdmin ? '' : (compradorProjecaoAtual || usuarioAtual.nome.split(" ")[0].toUpperCase());
    try {
        const url = comprador
            ? addAuth(`${URL_SCRIPT}?action=getNotificacoes&comprador=${encodeURIComponent(comprador)}&t=${Date.now()}`)
            : addAuth(`${URL_SCRIPT}?action=getNotificacoes&t=${Date.now()}`);
        const resp = await fetch(url);
        const data = await resp.json();
        notificacoesPendentes = data.notificacoes || [];
        // Marca todos como "já vistos" no primeiro load — evita toasts de notificações antigas
        notificacoesPendentes.forEach(n => _ultimasNotifIds.add(n.linha));
        atualizarSinoNotificacoes();
    } catch(e) { console.warn("Erro ao carregar notificações:", e); }
}

function atualizarSinoNotificacoes() {
    const contador = document.getElementById('notifContador');
    const btn      = document.getElementById('notifHeaderBtn');
    if (!contador || !btn) return;
    const qtd = notificacoesPendentes.length;
    if (qtd > 0) {
        contador.style.display = 'flex';
        contador.textContent   = qtd > 9 ? '9+' : qtd;
        btn.classList.add('tem-notif');
    } else {
        contador.style.display = 'none';
        btn.classList.remove('tem-notif');
    }
}

async function abrirPainelNotificacoes() {
    const drawer = document.getElementById('notifDrawer');
    // Toggle — fecha se já estiver aberto
    if (drawer.style.display === 'flex') {
        fecharPainelNotificacoes();
        return;
    }

    const lista = document.getElementById('notifLista');
    document.getElementById('notifOverlay').style.display = 'block';
    drawer.style.display = 'flex';

    if (!notificacoesPendentes.length) {
        lista.innerHTML = '<div class="notif-vazia"><i class="ph ph-bell-slash"></i><span>Nenhuma notificação pendente</span></div>';
        return;
    }

    lista.innerHTML = notificacoesPendentes.map(n => `
        <div class="notif-item" data-linha="${n.linha}" data-codigo="${n.codigo}">
            <div class="notif-item-header">
                <span class="notif-item-codigo">${n.codigo}${temPermissao('administrador') && n.comprador ? ` · <span style="color:var(--accent);font-size:10px;">${n.comprador}</span>` : ''}</span>
                <span class="notif-item-data">${n.data}</span>
            </div>
            <p class="notif-item-msg">${n.mensagem}</p>
            ${n.tipo === 'ENTREGA_PERGUNTA' ? `
            <div class="notif-item-acoes">
                <button class="btn-notif-sim" onclick="responderEntrega(${n.linha}, '${n.codigo}', true, this)">
                    <i class="ph ph-check"></i> SIM, FOI ENTREGUE
                </button>
                <button class="btn-notif-nao" onclick="responderEntrega(${n.linha}, '${n.codigo}', false, this)">
                    <i class="ph ph-x"></i> NÃO, AINDA NÃO
                </button>
            </div>` : ''}
        </div>`).join('');

    // Marca todas como lidas automaticamente ao abrir
    const todasLinhas = [...notificacoesPendentes];
    todasLinhas.forEach(n => {
        fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'marcarNotificacaoLida', linha: n.linha })
        });
        _ultimasNotifIds.add(n.linha);
    });

    // Limpa localmente após um pequeno delay (deixa o usuário ver)
    setTimeout(() => {
        notificacoesPendentes = notificacoesPendentes.filter(n => n.tipo === 'ENTREGA_PERGUNTA');
        atualizarSinoNotificacoes();
    }, 1500);
}

function fecharPainelNotificacoes() {
    document.getElementById('notifOverlay').style.display = 'none';
    document.getElementById('notifDrawer').style.display  = 'none';
}

async function responderEntrega(linha, codigo, foiEntregue, btn) {
    const comprador = compradorProjecaoAtual || usuarioAtual.nome.split(" ")[0].toUpperCase();
    btn.closest('.notif-item').style.opacity = '0.5';

    // Marca notificação como lida
    await fetch(URL_SCRIPT, {
        method: 'POST', mode: 'no-cors',
        body: JSON.stringify({ tipo: 'marcarNotificacaoLida', linha })
    });

    if (foiEntregue) {
        // Remove do StatusCompra e da projeção local
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'removerStatusCompra', codigo, comprador })
        });
        delete statusCompraMap[codigo];
        dadosProjecao = dadosProjecao.filter(i => i.codigo !== codigo);
        aplicarFiltroAtual();
        renderizarGraficoProcesso();
        mostrarToast('✅ Item removido da projeção — voltará quando a cobertura cair novamente', 'success');
    } else {
        // Mantém com status AGUARDANDO ENTREGA DO FORNECEDOR
        if (statusCompraMap[codigo]) {
            statusCompraMap[codigo].status = 'AGUARDANDO ENTREGA DO FORNECEDOR';
        }
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'salvarStatusCompra', codigo, descricao: statusCompraMap[codigo]?.descricao || '', comprador, status: 'AGUARDANDO ENTREGA DO FORNECEDOR', cobertura: statusCompraMap[codigo]?.coberturaRegistrada || 0 })
        });
        atualizarTabelaProjecao(itensFiltradosProjecao);
        renderizarGraficoProcesso();
        mostrarToast('ℹ️ Item mantido na projeção como "Aguardando Entrega"', 'info');
    }

    notificacoesPendentes = notificacoesPendentes.filter(n => n.linha !== linha);
    atualizarSinoNotificacoes();
    btn.closest('.notif-item').remove();
    if (!document.querySelectorAll('.notif-item').length) {
        document.getElementById('notifLista').innerHTML = '<div class="notif-vazia"><i class="ph ph-bell-slash"></i><span>Nenhuma notificação pendente</span></div>';
    }
}

async function marcarNotifLida(linha, btn) {
    await fetch(URL_SCRIPT, {
        method: 'POST', mode: 'no-cors',
        body: JSON.stringify({ tipo: 'marcarNotificacaoLida', linha })
    });
    notificacoesPendentes = notificacoesPendentes.filter(n => n.linha !== linha);
    atualizarSinoNotificacoes();
    btn.closest('.notif-item').remove();
}

function aplicarFiltroAtual() {
    let itens = aplicarFiltroPrefixo(dadosProjecao);
    // Aplica o ponto de compra — mostra apenas itens com cobertura dentro do limite
    itens = itens.filter(i => i.cobertura <= pontoCompraAtual);
    if (filtroProjecaoAtual === 'rp')          itens = itens.filter(i => i.temRP);
    else if (filtroProjecaoAtual === 'cd')      itens = itens.filter(i => i.saldoCD > 0);
    else if (filtroProjecaoAtual === 'empenho') itens = itens.filter(i => i.temEmpenho);
    else if (filtroProjecaoAtual === 'zerado')          itens = itens.filter(i => i.zeradoSemCobertura);
    else if (filtroProjecaoAtual === 'zeradoCobertura') itens = itens.filter(i => i.cobertura <= 0);

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

let _usuarioEditando = null;
let _loginParaDeletar = null;

const TODAS_PERMS = ['digitador', 'gestor', 'comprador', 'diretor', 'administrador'];

const PERM_BADGE = {
    digitador:     { txt: 'Digitador',    cor: 'rgba(100,116,139,0.15)', color: 'var(--text-muted)' },
    gestor:        { txt: 'Gestor',       cor: 'rgba(2,132,199,0.15)',   color: 'var(--accent)'     },
    comprador:     { txt: 'Comprador',    cor: 'rgba(22,163,74,0.15)',   color: 'var(--success)'    },
    diretor:       { txt: 'Diretor',      cor: 'rgba(139,92,246,0.15)',  color: '#8b5cf6'           },
    administrador: { txt: 'Admin',        cor: 'rgba(245,158,11,0.15)',  color: '#f59e0b'           }
};

async function carregarUsuarios() {
    document.getElementById('admin-loading').style.display = 'flex';
    document.querySelector('#tabelaUsuarios tbody').innerHTML = '';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getUsuarios&solicitante=${encodeURIComponent(loginAtual)}`));
        const data = await res.json();
        if (data.erro) throw new Error(data.erro);
        renderizarTabelaUsuarios(data);
    } catch (e) {
        document.querySelector('#tabelaUsuarios tbody').innerHTML =
            `<tr><td colspan="4" style="text-align:center;color:var(--danger)">Erro ao carregar usuários.</td></tr>`;
    } finally {
        document.getElementById('admin-loading').style.display = 'none';
    }
}

function renderizarTabelaUsuarios(lista) {
    const tbody = document.querySelector('#tabelaUsuarios tbody');
    tbody.innerHTML = '';
    if (!lista.length) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center">Nenhum usuário encontrado.</td></tr>';
        return;
    }
    lista.forEach(u => {
        const perms = (u.permissoes || u.role || 'digitador').toString().split(',').map(p => p.trim());
        const badges = perms.map(p => {
            const b = PERM_BADGE[p]; if (!b) return '';
            return `<span style="background:${b.cor};color:${b.color};padding:3px 9px;border-radius:6px;font-size:10px;font-weight:800;text-transform:uppercase;margin-right:4px;display:inline-block;margin-bottom:3px;">${b.txt}</span>`;
        }).join('');
        const status = u.primeiroAcesso === 'true' || u.primeiroAcesso === true
            ? '<span class="badge-primeiro-acesso">⏳ Aguardando 1º acesso</span>'
            : '<span class="badge-ativo">✓ Ativo</span>';
        const isSelf = u.login === loginAtual;
        const permsEsc    = perms.join(',').replace(/'/g, "\\'");
        const prefixosEsc = (u.prefixos || '').replace(/'/g, "\\'");
        const emailEsc    = (u.email || '').replace(/'/g, "\\'");
        tbody.innerHTML += `
            <tr>
                <td><b>${u.nome}</b><br><span style="font-size:11px;color:var(--text-muted)">${u.login}</span></td>
                <td>${badges}</td>
                <td>${status}</td>
                <td style="text-align:center;">
                    <div class="admin-acoes">
                        <button class="btn-admin-acao ver-como" onclick="verComoUsuario('${u.login}','${u.nome.replace(/'/g,"\\'")}');" title="Ver como este usuário"><i class="ph ph-eye"></i></button>
                        <button class="btn-admin-acao editar"  onclick="abrirModalUsuario('${u.login}','${u.nome.replace(/'/g,"\\'")}','${permsEsc}','${prefixosEsc}','${emailEsc}')" title="Editar"><i class="ph ph-pencil-simple"></i></button>
                        <button class="btn-admin-acao resetar" onclick="abrirModalResetarSenha('${u.login}','${u.nome.replace(/'/g,"\\'")}');" title="Resetar senha"><i class="ph ph-key"></i></button>
                        <button class="btn-admin-acao deletar" onclick="abrirModalDeletar('${u.login}','${u.nome.replace(/'/g,"\\'")}');" title="Remover" ${isSelf ? 'disabled style="opacity:0.3;cursor:not-allowed"' : ''}><i class="ph ph-trash"></i></button>
                    </div>
                </td>
            </tr>`;
    });
}

function abrirModalUsuario(login, nome, permsStr, prefixos, email) {
    const editando = !!login;
    _usuarioEditando = editando ? login : null;
    const perms = permsStr ? permsStr.split(',').map(p => p.trim()) : ['digitador'];

    document.getElementById('modalUsuarioTitulo').innerHTML = editando
        ? '<i class="ph ph-pencil-simple"></i> EDITAR USUÁRIO'
        : '<i class="ph ph-user-plus"></i> NOVO USUÁRIO';

    document.getElementById('u_nome').value  = nome  || '';
    document.getElementById('u_login').value = login || '';
    document.getElementById('u_email').value = email || '';
    document.getElementById('u_login').disabled = editando;
    document.getElementById('u_senha_group').style.display   = editando ? 'none'  : 'flex';
    document.getElementById('u_editando_info').style.display = editando ? 'block' : 'none';
    document.getElementById('u_senha').value = editando ? '' : 'Core@26';

    // Prefixos — visível só para comprador
    const isComprador = perms.includes('comprador');
    document.getElementById('u_prefixos_group').style.display = isComprador ? 'block' : 'none';
    document.getElementById('u_prefixos').value = prefixos || '';

    // Marca os checkboxes das permissões
    TODAS_PERMS.forEach(p => {
        const el = document.getElementById(`p_${p}`);
        if (el) el.checked = perms.includes(p);
    });

    document.getElementById('modalUsuario').style.display = 'flex';
}

async function salvarUsuario() {
    const nome     = document.getElementById('u_nome').value.trim();
    const login    = document.getElementById('u_login').value.trim().toLowerCase();
    const senha    = document.getElementById('u_senha').value.trim();
    const editando = !!_usuarioEditando;

    if (!nome || !login) return alert('⚠️ Preencha nome e login.');
    if (!editando && senha.length < 6) return alert('⚠️ A senha inicial deve ter pelo menos 6 caracteres.');

    const perms = TODAS_PERMS.filter(p => document.getElementById(`p_${p}`)?.checked);
    if (perms.length === 0) return alert('⚠️ Selecione ao menos uma permissão.');

    const prefixos  = document.getElementById('u_prefixos').value.trim();
    const email     = document.getElementById('u_email').value.trim();
    const senhaHash = !editando ? senha : null;

    const btn = document.getElementById('btnSalvarUsuario');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> SALVANDO...';

    try {
        const action = editando ? 'editarUsuario' : 'criarUsuario';
        let url = addAuth(`${URL_SCRIPT}?action=${action}&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(login)}&nome=${encodeURIComponent(nome)}&permissoes=${encodeURIComponent(perms.join(','))}&prefixos=${encodeURIComponent(prefixos)}&email=${encodeURIComponent(email)}`);
        if (!editando) url += `&senha=${encodeURIComponent(senhaHash)}`;
        const res  = await fetch(url);
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro || 'Erro desconhecido');
        fecharModal('modalUsuario');
        tocarSomMSN();
        registrarLog(editando ? 'EDITAR USUÁRIO' : 'CRIAR USUÁRIO', `Login: ${login} — Permissões: ${perms.join(',')}`);
        await carregarUsuarios();
    } catch (e) {
        alert('Erro ao salvar: ' + e.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-check-circle"></i> SALVAR';
    }
}

let _loginParaResetar = null;
let _nomeParaResetar  = null;

function abrirModalResetarSenha(login, nome) {
    _loginParaResetar = login;
    _nomeParaResetar  = nome;
    document.getElementById('resetarUsuarioNome').innerText  = nome;
    document.getElementById('resetarUsuarioLogin').innerText = login;
    document.getElementById('modalResetarSenha').style.display = 'flex';
}

async function confirmarResetarSenha() {
    const login = _loginParaResetar;
    const nome  = _nomeParaResetar;
    if (!login) return;
    fecharModal('modalResetarSenha');
    _loginParaResetar = null;
    _nomeParaResetar  = null;

    const novaHash = 'Core@26'; // GS faz o hash server-side
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=resetarSenha&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(login)}&novaSenha=${encodeURIComponent(novaHash)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        tocarSomMSN();
        mostrarToast(`✅ Senha de ${nome} redefinida para Core@26.`, 'success');
        registrarLog('RESETAR SENHA', `Login: ${login}`);
        await carregarUsuarios();
    } catch (e) {
        mostrarToast('Erro ao resetar senha: ' + e.message, 'error');
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
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=deletarUsuario&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(_loginParaDeletar)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        tocarSomMSN();
        registrarLog('DELETAR USUÁRIO', `Login removido: ${_loginParaDeletar}`);
        await carregarUsuarios();
    } catch (e) {
        alert('Erro ao remover usuário: ' + e.message);
    } finally {
        _loginParaDeletar = null;
    }
}

// =========================================================
// MIGRAÇÃO DE SENHAS PARA SHA-256
// =========================================================
async function migrarSenhas() {
    if (!confirm('Isso vai converter todas as senhas em texto puro para hash SHA-256.\n\nOs usuários continuarão usando a mesma senha — só o armazenamento muda.\n\nContinuar?')) return;

    const btn = event.currentTarget;
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> MIGRANDO...';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=migrarSenhas&solicitante=${encodeURIComponent(loginAtual)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        mostrarToast(`✅ ${data.migradas} senha(s) convertidas para hash SHA-256.`, 'success', 5000);
        registrarLog('MIGRAÇÃO SENHAS', `${data.migradas} senha(s) hasheadas`);
    } catch (e) {
        mostrarToast('Erro na migração: ' + e.message, 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-shield-check"></i> MIGRAR SENHAS PARA HASH';
    }
}
function switchAdminTab(tab) {
    document.getElementById('adminTabUsuarios').style.display  = tab === 'usuarios'  ? 'block' : 'none';
    document.getElementById('adminTabBlacklist').style.display = tab === 'blacklist' ? 'block' : 'none';
    document.getElementById('adminTabLog').style.display       = tab === 'log'       ? 'block' : 'none';
    document.getElementById('adminTabSistema').style.display   = tab === 'sistema'   ? 'block' : 'none';
    document.getElementById('tabBtnUsuarios').classList.toggle('ativo',  tab === 'usuarios');
    document.getElementById('tabBtnBlacklist').classList.toggle('ativo', tab === 'blacklist');
    document.getElementById('tabBtnLog').classList.toggle('ativo',       tab === 'log');
    document.getElementById('tabBtnSistema').classList.toggle('ativo',   tab === 'sistema');
    if (tab === 'blacklist') { carregarBlacklist(); preencherCompradoresAddProj(); }
    if (tab === 'log')       carregarLog();
}

async function ativarTriggerNotificacoes() {
    try {
        await fetch(addAuth(`${URL_SCRIPT}?action=ativarTrigger`));
        mostrarToast('✅ Trigger ativado — verificação roda a cada 1 hora', 'success');
    } catch(e) {
        mostrarToast('Erro ao ativar trigger', 'error');
    }
}

async function rodarVerificacaoAgora() {
    mostrarToast('⏳ Rodando verificação...', 'info');
    try {
        await fetch(addAuth(`${URL_SCRIPT}?action=rodarVerificacaoAgora`));
        mostrarToast('✅ Verificação concluída — cheque as notificações', 'success');
        await carregarNotificacoes();
    } catch(e) {
        mostrarToast('Erro ao rodar verificação', 'error');
    }
}

// =========================================================
// LOG DE ATIVIDADES
// =========================================================

// Grava uma entrada no log (fire-and-forget via POST no-cors)
function registrarLog(acao, detalhe = '') {
    fetch(URL_SCRIPT, {
        method: 'POST',
        mode: 'no-cors',
        body: JSON.stringify({
            tipo: 'log',
            login:   loginAtual || '',
            nome:    usuarioAtual?.nome || loginAtual || '',
            acao:    acao,
            detalhe: detalhe
        })
    }).catch(() => {});
}

function confirmarLimparLog() {
    document.getElementById('modalLimparLog').style.display = 'flex';
}

async function executarLimparLog() {
    fecharModal('modalLimparLog');
    const btn = document.querySelector('.btn-limpar-log');
    btn.disabled = true;
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> LIMPANDO...';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=limparLog&solicitante=${encodeURIComponent(loginAtual)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        _logCompleto = [];
        renderizarTabelaLog([]);
        mostrarToast('Log limpo com sucesso.', 'success');
        registrarLog('LIMPAR LOG', 'Log de atividades apagado pelo administrador');
    } catch (e) {
        mostrarToast('Erro ao limpar log: ' + e.message, 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="ph ph-trash"></i> LIMPAR LOG';
    }
}

async function carregarLog() {
    document.getElementById('log-loading').style.display = 'flex';
    document.querySelector('#tabelaLog tbody').innerHTML = '';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getLog&solicitante=${encodeURIComponent(loginAtual)}`));
        const data = await res.json();
        if (data.erro) throw new Error(data.erro);
        _logCompleto = [...data].reverse(); // mais recentes primeiro
        renderizarTabelaLog(_logCompleto);
    } catch (e) {
        document.querySelector('#tabelaLog tbody').innerHTML =
            `<tr><td colspan="4" style="text-align:center;color:var(--danger)">Erro ao carregar log.</td></tr>`;
    } finally {
        document.getElementById('log-loading').style.display = 'none';
    }
}

function filtrarLog() {
    const inicio = document.getElementById('logInicio').value;
    const fim    = document.getElementById('logFim').value;

    if (!inicio || !fim) { mostrarToast('⚠️ Selecione início e fim.', 'warning'); return; }

    const dtInicio = new Date(inicio + 'T00:00:00');
    const dtFim    = new Date(fim    + 'T23:59:59');

    const filtrado = _logCompleto.filter(entry => {
        // dataHora formato: "21/04/2026 14:52:53" — converte para Date
        const partes = entry.dataHora?.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (!partes) return false;
        const dt = new Date(`${partes[3]}-${partes[2]}-${partes[1]}T${partes[4]}:${partes[5]}:${partes[6]}`);
        return dt >= dtInicio && dt <= dtFim;
    });

    renderizarTabelaLog(filtrado);
    document.getElementById('btnLimparLog').style.display = 'inline-flex';
    mostrarToast(`${filtrado.length} registro(s) no período.`, filtrado.length ? 'success' : 'info');
}

function limparFiltroLog() {
    document.getElementById('logInicio').value = '';
    document.getElementById('logFim').value    = '';
    document.getElementById('btnLimparLog').style.display = 'none';
    renderizarTabelaLog(_logCompleto);
}

function renderizarTabelaLog(lista) {
    const tbody = document.querySelector('#tabelaLog tbody');
    tbody.innerHTML = '';
    const resumo = document.getElementById('logResumo');

    if (!lista.length) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:24px;color:var(--text-muted);">
            <i class="ph ph-scroll" style="font-size:24px;display:block;margin-bottom:8px;"></i>
            Nenhuma atividade registrada ainda.</td></tr>`;
        if (resumo) resumo.textContent = '';
        return;
    }

    lista.forEach(entry => {
        const cor = entry.acao?.includes('LOGIN') ? 'var(--accent)'
                  : entry.acao?.includes('DELETAR') || entry.acao?.includes('REMOVER') ? 'var(--danger)'
                  : entry.acao?.includes('SYNC') || entry.acao?.includes('ENVIO') ? 'var(--success)'
                  : entry.acao?.includes('BLACKLIST') ? '#f59e0b'
                  : 'var(--text-muted)';
        tbody.innerHTML += `
            <tr>
                <td style="font-size:11px;color:var(--text-muted);white-space:nowrap;">${entry.dataHora || '—'}</td>
                <td><b>${entry.nome || entry.login || '—'}</b></td>
                <td><span style="background:${cor}22;color:${cor};padding:3px 10px;border-radius:6px;font-size:10px;font-weight:900;text-transform:uppercase;white-space:nowrap;">${entry.acao || '—'}</span></td>
                <td style="font-size:12px;color:var(--text-muted);">${entry.detalhe || '—'}</td>
            </tr>`;
    });

    if (resumo) resumo.textContent = `${lista.length} registro(s) exibido(s)`;
}

// =========================================================
// BLACKLIST
// =========================================================

// Carrega só os códigos (cache leve) — para filtro na projeção
// =========================================================
// MODAL SALDO SUBESTOQUES
// =========================================================

async function abrirSaldoSub(codigo, consumoDiario) {
    document.getElementById('modalSaldoSub').style.display = 'flex';
    document.getElementById('modalSaldoSubCodigo').textContent = codigo;
    document.getElementById('modalSaldoSubLoading').style.display = 'block';
    document.getElementById('modalSaldoSubConteudo').innerHTML = '';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getSaldoSub&codigo=${encodeURIComponent(codigo)}`));
        const data = await res.json();

        if (data.erro || !data.length) {
            document.getElementById('modalSaldoSubConteudo').innerHTML =
                `<p style="text-align:center;color:var(--text-muted);padding:24px;">Nenhum saldo encontrado para este item.</p>`;
            return;
        }

        const totalQtd = data.reduce((s, r) => s + (r.quantidade || 0), 0);
        const coberturaTotal = consumoDiario > 0 ? Math.round(totalQtd / consumoDiario) : '—';

        let html = `
            <div style="background:var(--bg-system);border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;gap:24px;flex-wrap:wrap;">
                <div><small style="color:var(--text-muted);font-size:10px;">TOTAL EM ESTOQUE</small><p style="font-weight:900;font-size:18px;color:var(--accent);">${totalQtd}</p></div>
                <div><small style="color:var(--text-muted);font-size:10px;">COBERTURA TOTAL</small><p style="font-weight:900;font-size:18px;color:var(--accent);">${coberturaTotal}${typeof coberturaTotal === 'number' ? ' dias' : ''}</p></div>
            </div>
            <div style="margin-top:12px;">
            <table style="font-size:13px;width:100%;">
                <thead><tr>
                    <th style="text-align:left;">SUBESTOQUE</th>
                    <th style="text-align:right;">QUANTIDADE</th>
                    <th style="text-align:right;">COBERTURA (DIAS)</th>
                </tr></thead>
                <tbody>`;

        data.forEach(r => {
            const cob = consumoDiario > 0 && r.quantidade > 0
                ? Math.round(r.quantidade / consumoDiario)
                : '—';
            const corCob = typeof cob === 'number'
                ? cob < 25 ? 'color:#ef4444;' : cob < 90 ? 'color:#f59e0b;' : 'color:#10b981;'
                : '';
            html += `<tr>
                <td>${r.subestoque}</td>
                <td style="text-align:right;font-weight:700;">${r.quantidade}</td>
                <td style="text-align:right;font-weight:900;${corCob}">${cob}</td>
            </tr>`;
        });

        html += '</tbody></table></div>';
        document.getElementById('modalSaldoSubConteudo').innerHTML = html;
    } catch(e) {
        document.getElementById('modalSaldoSubConteudo').innerHTML =
            `<p style="color:var(--danger);text-align:center;padding:20px;">Erro ao carregar: ${e.message}</p>`;
    } finally {
        document.getElementById('modalSaldoSubLoading').style.display = 'none';
    }
}

let _addProjItemCache = null;

function limparPreviewAddProj() {
    document.getElementById('addProjPreview').style.display = 'none';
    document.getElementById('addProjErro').style.display = 'none';
    _addProjItemCache = null;
}

async function preencherCompradoresAddProj() {
    const sel = document.getElementById('addProjComprador');
    if (!sel || sel.options.length > 1) return;
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getCompradores`));
        const data = await res.json();
        sel.innerHTML = data.map(c => `<option value="${c.nomeUsuario}">${c.nome}</option>`).join('');
    } catch(e) { sel.innerHTML = '<option value="">Erro ao carregar</option>'; }
}

async function buscarItemParaProj() {
    const codigo  = document.getElementById('addProjCodigo').value.trim().toUpperCase();
    const consumo = document.getElementById('addProjConsumo').value.trim();
    const erroEl  = document.getElementById('addProjErro');
    const prevEl  = document.getElementById('addProjPreview');

    limparPreviewAddProj();

    if (!codigo) { erroEl.textContent = 'Informe o código do item.'; erroEl.style.display = 'block'; return; }
    if (!consumo || isNaN(parseFloat(consumo))) { erroEl.textContent = 'Informe o consumo diário.'; erroEl.style.display = 'block'; return; }

    const btn = document.querySelector('[onclick="buscarItemParaProj()"]');
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> BUSCANDO...';
    btn.disabled = true;

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=buscarItemBaseDados&codigo=${encodeURIComponent(codigo)}`));
        const data = await res.json();

        if (data.erro || !data.codigo) {
            erroEl.textContent = data.erro || 'Item não encontrado na BASE_DADOS.';
            erroEl.style.display = 'block';
            return;
        }

        _addProjItemCache = { ...data, consumoDiario: parseFloat(consumo) };

        const consumoMensal = (parseFloat(consumo) * 30).toFixed(3);
        document.getElementById('addProjDados').innerHTML = `
            <div><small style="color:var(--text-muted);font-size:10px;">CÓDIGO</small><p style="font-weight:900;font-family:monospace;">${data.codigo}</p></div>
            <div style="grid-column:span 2"><small style="color:var(--text-muted);font-size:10px;">DESCRIÇÃO</small><p style="font-weight:700;font-size:13px;">${data.descricao}</p></div>
            <div><small style="color:var(--text-muted);font-size:10px;">COBERTURA ATUAL</small><p style="font-weight:700;">${data.cobertura ?? '—'}</p></div>
            <div><small style="color:var(--text-muted);font-size:10px;">STATUS</small><p style="font-weight:700;">${data.status || '—'}</p></div>
            <div><small style="color:var(--text-muted);font-size:10px;">SALDO CD (150)</small><p style="font-weight:700;">${data.saldo150 ?? '—'}</p></div>
            <div><small style="color:var(--text-muted);font-size:10px;">CONSUMO DIÁRIO (definido)</small><p style="font-weight:700;color:var(--accent);">${consumo}</p></div>
            <div><small style="color:var(--text-muted);font-size:10px;">CONSUMO MENSAL EST.</small><p style="font-weight:700;color:var(--accent);">${consumoMensal}</p></div>`;

        prevEl.style.display = 'block';
    } catch(e) {
        erroEl.textContent = 'Erro ao buscar item: ' + e.message;
        erroEl.style.display = 'block';
    } finally {
        btn.innerHTML = '<i class="ph ph-magnifying-glass"></i> BUSCAR ITEM';
        btn.disabled = false;
    }
}

async function confirmarAddProj() {
    if (!_addProjItemCache) return;
    const comprador = document.getElementById('addProjComprador').value;
    if (!comprador) { mostrarToast('Selecione o comprador.', 'warning'); return; }

    const btn = document.querySelector('[onclick="confirmarAddProj()"]');
    btn.innerHTML = '<i class="ph ph-circle-notch rotating"></i> ADICIONANDO...';
    btn.disabled = true;

    try {
        const res = await fetch(addAuth(`${URL_SCRIPT}?action=adicionarItemProjecao`), {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify({
                tipo: 'adicionarItemProjecao',
                codigo:       _addProjItemCache.codigo,
                descricao:    _addProjItemCache.descricao,
                cobertura:    _addProjItemCache.cobertura,
                status:       _addProjItemCache.status,
                saldo150:     _addProjItemCache.saldo150,
                consumoDiario: _addProjItemCache.consumoDiario,
                comprador,
                solicitante:  loginAtual
            })
        });

        mostrarToast(`Item ${_addProjItemCache.codigo} adicionado à projeção de ${comprador}!`, 'success');
        registrarLog('ADD PROJEÇÃO', `Código ${_addProjItemCache.codigo} adicionado à projeção de ${comprador}`);
        limparPreviewAddProj();
        document.getElementById('addProjCodigo').value = '';
        document.getElementById('addProjConsumo').value = '';
    } catch(e) {
        mostrarToast('Erro ao adicionar item.', 'error');
    } finally {
        btn.innerHTML = '<i class="ph ph-check-circle"></i> CONFIRMAR E ADICIONAR';
        btn.disabled = false;
    }
}

async function carregarBlacklistCache() {
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getBlacklist`));
        const data = await res.json();
        blacklistCodigos = new Set((data || []).map(item => item.codigo.trim().toUpperCase()));
    } catch (e) {
        console.warn('Erro ao carregar blacklist cache:', e);
    }
}

// Carrega a lista completa para exibição no painel admin
async function carregarBlacklist() {
    document.getElementById('blacklist-loading').style.display = 'flex';
    document.querySelector('#tabelaBlacklist tbody').innerHTML = '';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getBlacklist`));
        const data = await res.json();
        renderizarTabelaBlacklist(data || []);
    } catch (e) {
        document.querySelector('#tabelaBlacklist tbody').innerHTML =
            `<tr><td colspan="5" style="text-align:center;color:var(--danger)">Erro ao carregar blacklist.</td></tr>`;
    } finally {
        document.getElementById('blacklist-loading').style.display = 'none';
    }
}

function renderizarTabelaBlacklist(lista) {
    const tbody = document.querySelector('#tabelaBlacklist tbody');
    tbody.innerHTML = '';
    if (!lista.length) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--text-muted);padding:24px;">
            <i class="ph ph-check-circle" style="font-size:24px;display:block;margin-bottom:8px;"></i>
            Nenhum código bloqueado.</td></tr>`;
        return;
    }
    lista.forEach((item, idx) => {
        tbody.innerHTML += `
            <tr>
                <td><b style="font-family:monospace;">${item.codigo}</b></td>
                <td style="color:var(--text-muted);font-size:12px;">${item.motivo || '—'}</td>
                <td style="font-size:12px;">${item.adicionadoPor || '—'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.data || '—'}</td>
                <td style="text-align:center;">
                    <button class="btn-admin-acao deletar" onclick="removerBlacklist('${item.codigo}')" title="Desbloquear">
                        <i class="ph ph-x-circle"></i>
                    </button>
                </td>
            </tr>`;
    });
}

function abrirModalBlacklist() {
    document.getElementById('bl_codigo').value = '';
    document.getElementById('bl_motivo').value = '';
    document.getElementById('modalBlacklist').style.display = 'flex';
    setTimeout(() => document.getElementById('bl_codigo').focus(), 100);
}

async function salvarBlacklist() {
    const codigo = document.getElementById('bl_codigo').value.trim().toUpperCase();
    const motivo = document.getElementById('bl_motivo').value.trim();
    if (!codigo) return alert('⚠️ Informe o código do item.');

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=addBlacklist&solicitante=${encodeURIComponent(loginAtual)}&codigo=${encodeURIComponent(codigo)}&motivo=${encodeURIComponent(motivo)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        fecharModal('modalBlacklist');
        blacklistCodigos.add(codigo);
        tocarSomMSN();
        registrarLog('BLACKLIST ADD', `Código bloqueado: ${codigo}${motivo ? ' — ' + motivo : ''}`);
        await carregarBlacklist();
    } catch (e) {
        alert('Erro ao bloquear código: ' + e.message);
    }
}

async function adicionarBlacklistMulti() {
    const raw = document.getElementById('inputBlacklistMulti').value.trim();
    if (!raw) return;
    const codigos = raw.split(/[\s,;]+/).map(c => c.trim().toUpperCase()).filter(Boolean);
    if (!codigos.length) return;

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=addBlacklistLote&solicitante=${encodeURIComponent(loginAtual)}&codigos=${encodeURIComponent(codigos.join(','))}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        codigos.forEach(c => blacklistCodigos.add(c));
        document.getElementById('inputBlacklistMulti').value = '';
        tocarSomMSN();
        await carregarBlacklist();
    } catch (e) {
        alert('Erro ao adicionar lote: ' + e.message);
    }
}

async function removerBlacklist(codigo) {
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=removerBlacklist&solicitante=${encodeURIComponent(loginAtual)}&codigo=${encodeURIComponent(codigo)}`));
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro);
        blacklistCodigos.delete(codigo.trim().toUpperCase());
        tocarSomMSN();
        registrarLog('BLACKLIST REMOVE', `Código desbloqueado: ${codigo}`);
        await carregarBlacklist();
    } catch (e) {
        alert('Erro ao remover da blacklist: ' + e.message);
    }
}
// =============================================================================
// NAV SUPERTECH — INDICADOR DESLIZANTE + TEXT SCRAMBLE
// =============================================================================

function atualizarNavIndicador() {
    const indicator = document.getElementById('navIndicador') || document.getElementById('navIndicator');
    const activeBtn = document.querySelector('.nav-btn.active');
    if (!indicator || !activeBtn) return;
    const container = document.getElementById('navContainer');
    const containerRect = container.getBoundingClientRect();
    const btnRect = activeBtn.getBoundingClientRect();
    indicator.style.left  = (btnRect.left - containerRect.left + btnRect.width * 0.1) + 'px';
    indicator.style.width = (btnRect.width * 0.8) + 'px';
}

// Text scramble ao passar o mouse
const SCRAMBLE_CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@#$%&';
function scrambleText(el) {
    const original = el.dataset.original || el.textContent.trim();
    el.dataset.original = original;
    let frame = 0;
    const totalFrames = original.length * 2;
    if (el._scrambleInterval) clearInterval(el._scrambleInterval);
    el._scrambleInterval = setInterval(() => {
        el.textContent = original.split('').map((char, i) => {
            if (char === ' ') return ' ';
            if (i < Math.floor(frame / 2)) return original[i];
            return SCRAMBLE_CHARS[Math.floor(Math.random() * SCRAMBLE_CHARS.length)];
        }).join('');
        frame++;
        if (frame >= totalFrames) {
            el.textContent = original;
            clearInterval(el._scrambleInterval);
        }
    }, 28);
}

// Inicializa os efeitos do nav após o login
function iniciarNavAnimacoes() {
    // Scramble em todos os botões
    document.querySelectorAll('.nav-btn').forEach(btn => {
        const textSpan = btn.querySelector('.nav-btn-text');
        if (!textSpan) return;
        btn.addEventListener('mouseenter', () => scrambleText(textSpan));
    });

    // Observer para novos botões de projeção adicionados dinamicamente
    const navContainer = document.getElementById('navContainer');
    if (navContainer) {
        const obs = new MutationObserver(() => {
            document.querySelectorAll('.nav-btn:not([data-scramble])').forEach(btn => {
                btn.setAttribute('data-scramble', '1');
                const textSpan = btn.querySelector('.nav-btn-text');
                if (textSpan) btn.addEventListener('mouseenter', () => scrambleText(textSpan));
            });
            atualizarNavIndicador();
        });
        obs.observe(navContainer, { childList: true, subtree: true });
    }

    // Posição inicial do indicador
    setTimeout(atualizarNavIndicador, 100);
    window.addEventListener('resize', atualizarNavIndicador);
}

// =============================================================================
// MODAL DE PERFIL + SOLICITAÇÃO DE TEMA
// =============================================================================

// Mapa de temas personalizados: login → nome do tema
const TEMAS_PERSONALIZADOS = {
    'supvitoria': 'Vitória'
};

function abrirModalPerfil() {
    const nome  = usuarioAtual?.nome || '';
    const login = loginAtual || '';
    document.getElementById('perfilNome').textContent  = nome;
    document.getElementById('perfilLogin').textContent = login;
    document.getElementById('perfilAvatar').textContent = nome.charAt(0).toUpperCase();
    document.getElementById('perfilPainelPrincipal').style.display = 'block';
    document.getElementById('perfilPainelTema').style.display      = 'none';
    document.getElementById('perfilTemaErro').textContent          = '';

    // Mostra seletor de tema personalizado se aplicável
    const temaNome = TEMAS_PERSONALIZADOS[login.toLowerCase()];
    const seletor  = document.getElementById('perfilSeletorTema');
    if (temaNome) {
        seletor.style.display = 'block';
        document.getElementById('btnTemaPersonalizadoLabel').textContent = temaNome;
        const usandoPersonalizado = document.body.getAttribute('data-user') === login.toLowerCase();
        document.getElementById('btnTemaPadrao').classList.toggle('ativo', !usandoPersonalizado);
        document.getElementById('btnTemaPersonalizado').classList.toggle('ativo', usandoPersonalizado);
    } else {
        seletor.style.display = 'none';
    }

    document.getElementById('modalPerfil').style.display = 'flex';
}

function selecionarTemaPerfil(opcao) {
    const login = loginAtual || '';
    if (opcao === 'personalizado') {
        document.body.setAttribute('data-user', login.toLowerCase());
        // Aplica tema padrão do usuário personalizado (claro)
        const temasFixos = { 'supvitoria': 'light' };
        const tema = localStorage.getItem('tema_' + login) || temasFixos[login.toLowerCase()] || 'dark';
        document.body.setAttribute('data-theme', tema);
        const icon = tema === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
        const h = document.getElementById('themeIcon');
        if (h) h.className = icon;
        localStorage.setItem('tema_modo_' + login, 'personalizado');
    } else {
        // Remove tema personalizado, volta ao padrão do sistema
        document.body.setAttribute('data-user', 'default');
        const temaSalvo = localStorage.getItem('tema_' + login) || 'dark';
        document.body.setAttribute('data-theme', temaSalvo);
        const icon = temaSalvo === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
        const h = document.getElementById('themeIcon');
        if (h) h.className = icon;
        localStorage.setItem('tema_modo_' + login, 'padrao');
    }
    // Atualiza botões
    document.getElementById('btnTemaPadrao').classList.toggle('ativo', opcao === 'padrao');
    document.getElementById('btnTemaPersonalizado').classList.toggle('ativo', opcao === 'personalizado');
}

function fecharModalPerfil(e) {
    if (e && e.target !== document.getElementById('modalPerfil')) return;
    const modal = document.getElementById('modalPerfil');
    modal.style.display = 'none';
}

function abrirSolicitarTema() {
    document.getElementById('perfilPainelPrincipal').style.display = 'none';
    document.getElementById('perfilPainelTema').style.display      = 'block';
}

function voltarPainelPerfil() {
    document.getElementById('perfilPainelPrincipal').style.display = 'block';
    document.getElementById('perfilPainelTema').style.display      = 'none';
}

async function enviarSolicitacaoTema() {
    const cor1  = document.getElementById('corPrimaria').value;
    const cor2  = document.getElementById('corSecundaria').value;
    const nome  = usuarioAtual?.nome || loginAtual;
    const login = loginAtual || '';
    const erro  = document.getElementById('perfilTemaErro');

    erro.textContent = '';
    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({
                tipo: 'solicitarTema',
                login, nome,
                corPrimaria: cor1,
                corSecundaria: cor2
            })
        });
        document.getElementById('modalPerfil').style.display = 'none';
        mostrarToast('✅ Solicitação enviada! O administrador receberá um email com seu pedido.', 'success', 5000);
    } catch(e) {
        erro.textContent = 'Erro ao enviar. Tente novamente.';
    }
}

// Aplica data-user no body para temas personalizados
function aplicarTemaUsuario(login) {
    document.body.setAttribute('data-user', login.toLowerCase());
    // Temas fixos por login
    const temaFixo = { 'supvitoria': 'light' };
    if (temaFixo[login.toLowerCase()]) {
        const tema = localStorage.getItem('tema_' + login) || temaFixo[login.toLowerCase()];
        document.body.setAttribute('data-theme', tema);
        const icon = tema === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
        const h = document.getElementById('themeIcon');
        if (h) h.className = icon;
    }
}

// =============================================================================
// VER COMO USUÁRIO (ADMIN) — IMPERSONAÇÃO TOTAL
// =============================================================================

let _adminSnapshot = null; // guarda estado completo do admin

async function verComoUsuario(login, nome) {
    // Salva estado completo do admin
    _adminSnapshot = {
        usuarioAtual: { ...usuarioAtual },
        loginAtual,
        tema: document.body.getAttribute('data-theme'),
        dataUser: document.body.getAttribute('data-user') || loginAtual,
    };

    // Busca dados completos do usuário alvo
    mostrarToast('⏳ Carregando perfil...', 'info', 2000);
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getUsuarios`));
        const lista = await res.json();
        const alvo = lista.find(u => u.login === login);
        if (!alvo) { mostrarToast('Usuário não encontrado', 'error'); return; }

        // Impersona
        loginAtual   = alvo.login;
        usuarioAtual = {
            nome:       alvo.nome,
            permissoes: alvo.permissoes || alvo.role || 'digitador',
            prefixos:   alvo.prefixos ? alvo.prefixos.split('|').map(p => p.trim()).filter(Boolean) : []
        };

        // Aplica tema do usuário
        aplicarTemaUsuario(login);
        const temaSalvo = localStorage.getItem('tema_' + login);
        if (temaSalvo) {
            document.body.setAttribute('data-theme', temaSalvo);
            const icon = temaSalvo === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
            const h = document.getElementById('themeIcon');
            if (h) h.className = icon;
        }

        // Reconstrói o nav com as permissões do usuário
        Object.keys(_cache).forEach(k => _cache[k] = false);
        dadosProjecao = [];
        compradorCarregado = '';
        configurarNavPorUsuario();
        switchTab('Dashboard');

        // Banner
        document.getElementById('verComoNome').textContent = nome;
        document.getElementById('verComoBanner').style.display = 'flex';
        mostrarToast(`👁 Visualizando como ${nome}`, 'info');

    } catch(e) {
        mostrarToast('Erro ao carregar perfil', 'error');
    }
}

function sairPreviewUsuario() {
    if (!_adminSnapshot) return;

    // Restaura estado do admin
    usuarioAtual = _adminSnapshot.usuarioAtual;
    loginAtual   = _adminSnapshot.loginAtual;

    document.body.setAttribute('data-user',  _adminSnapshot.dataUser);
    document.body.setAttribute('data-theme', _adminSnapshot.tema);
    const icon = _adminSnapshot.tema === 'dark' ? 'ph ph-sun' : 'ph ph-moon';
    const h = document.getElementById('themeIcon');
    if (h) h.className = icon;

    // Reconstrói nav com permissões do admin
    Object.keys(_cache).forEach(k => _cache[k] = false);
    dadosProjecao = [];
    compradorCarregado = '';
    configurarNavPorUsuario();
    switchTab('Admin');

    document.getElementById('verComoBanner').style.display = 'none';
    _adminSnapshot = null;
}

// =============================================================================
// PROTOCOLOS OPME ↔ SUPRIMENTOS
// =============================================================================

let protocoloLinhasData = []; // notas do formulário em edição
let protocoloAtualId = null;  // protocolo aberto no detalhe

// --- Detecção de login e configuração do nav ---

function isLoginOpme(login) {
    return (login || '').toLowerCase().startsWith('esp');
}

function configurarNavProtocolos() {
    const login = loginAtual || '';
    const isOpmeMode = modoAtual === 'opme';
    const isOpmeLogin = isLoginOpme(login);

    // Aba de ENVIO: só aparece para login esp... ou supmateus em modo OPME
    const mostrarEnvio = isOpmeLogin || (login === 'supmateus' && isOpmeMode);
    // Aba de RECEBIMENTO: qualquer usuário sup (digitador ou acima) em modo suprimentos
    const temAcessoSup = temPermissao('digitador') || temPermissao('gestor') || temPermissao('administrador');
    const mostrarRecebimento = !isOpmeLogin && !isOpmeMode && temAcessoSup;

    document.getElementById('btn-protocolos-opme').style.display = mostrarEnvio       ? 'inline-flex' : 'none';
    document.getElementById('btn-protocolos-sup').style.display  = mostrarRecebimento ? 'inline-flex' : 'none';
    // Float nav sync
    const fnOpme = document.getElementById('fn-protocolos-opme');
    const fnSup  = document.getElementById('fn-protocolos-sup');
    const fnDivExtra = document.getElementById('fn-div-extra');
    if (fnOpme) fnOpme.style.display = mostrarEnvio ? 'flex' : 'none';
    if (fnSup)  fnSup.style.display  = mostrarRecebimento ? 'flex' : 'none';
    if (fnDivExtra) fnDivExtra.style.display = (mostrarEnvio || mostrarRecebimento) ? 'block' : 'none';
}

// --- Integração com switchTab ---

// Adiciona flags de cache para protocolos
_cache.protocolosOpme = false;
_cache.protocolosSup  = false;

// --- Helpers de status ---

function badgeStatusProtocolo(status) {
    const mapa = {
        'ENVIADO':              { cor: '#f59e0b', bg: 'rgba(245,158,11,0.12)', icon: 'ph-paper-plane-tilt' },
        'RECEBIDO':             { cor: '#10b981', bg: 'rgba(16,185,129,0.12)', icon: 'ph-check-circle' },
        'DEVOLUÇÃO PARCIAL':    { cor: '#ef4444', bg: 'rgba(239,68,68,0.12)',  icon: 'ph-arrow-u-up-left' },
        'DEVOLVIDO CONFIRMADO': { cor: '#6366f1', bg: 'rgba(99,102,241,0.12)', icon: 'ph-check-fat' },
    };
    const s = mapa[status] || { cor: '#94a3b8', bg: 'rgba(148,163,184,0.1)', icon: 'ph-clock' };
    return `<span style="display:inline-flex;align-items:center;gap:5px;background:${s.bg};color:${s.cor};border:1px solid ${s.cor}33;border-radius:999px;padding:3px 10px;font-size:10px;font-weight:800;white-space:nowrap;">
        <i class="ph ${s.icon}"></i>${status}</span>`;
}

// --- OPME: carregar meus protocolos ---

async function carregarMeusProtocolos() {
    const pendentes = document.getElementById('protocolosPendentesBody');
    const historico  = document.getElementById('protocolosOpmeBody');
    if (pendentes) pendentes.innerHTML = '<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';
    if (historico)  historico.innerHTML  = '<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocolos&filtro=meus&t=${Date.now()}`));
        const data = await res.json();
        const lista = data.protocolos || [];

        const emAberto = lista.filter(p => p.status === 'ENVIADO' || p.status === 'DEVOLUÇÃO PARCIAL');
        const todos    = lista;

        const renderLinha = (p, isPendente) => {
            const acoes = `
                <div style="display:flex;gap:6px;flex-wrap:wrap;">
                    <button onclick="verDetalheProtocolo('${p.id}','${p.status}','${p.responsavel}','${p.dataCriacao}',false)"
                        style="background:rgba(6,182,212,0.1);border:1px solid rgba(6,182,212,0.3);border-radius:6px;padding:5px 10px;font-size:11px;font-weight:800;cursor:pointer;color:#06b6d4;display:flex;align-items:center;gap:5px;">
                        <i class="ph ph-eye"></i> Ver
                    </button>
                    <button onclick="exportarProtocoloCSVById('${p.id}','${p.responsavel}','${p.dataCriacao}')"
                        style="background:rgba(99,102,241,0.1);border:1px solid rgba(99,102,241,0.3);border-radius:6px;padding:5px 10px;font-size:11px;font-weight:800;cursor:pointer;color:#818cf8;display:flex;align-items:center;gap:5px;">
                        <i class="ph ph-file-csv"></i> CSV
                    </button>
                    ${p.status === 'DEVOLUÇÃO PARCIAL' ? `
                    <button onclick="confirmarDevolucao('${p.id}')"
                        style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);border-radius:6px;padding:5px 10px;font-size:11px;font-weight:800;cursor:pointer;color:#f87171;display:flex;align-items:center;gap:5px;">
                        <i class="ph ph-check"></i> Confirmar Devolução
                    </button>` : ''}
                </div>`;
            return `<tr style="border-top:1px solid var(--border);">
                <td style="padding:12px 20px;font-weight:900;font-family:monospace;font-size:12px;color:var(--accent);">${p.id}</td>
                <td style="padding:12px 20px;font-size:12px;color:var(--text-muted);">${p.dataCriacao}</td>
                <td style="padding:12px 20px;font-size:13px;font-weight:700;">${p.totalNotas} nota(s)</td>
                <td style="padding:12px 20px;">${badgeStatusProtocolo(p.status)}</td>
                <td style="padding:12px 20px;">${acoes}</td>
            </tr>`;
        };

        if (pendentes) {
            pendentes.innerHTML = emAberto.length
                ? emAberto.map(p => renderLinha(p, true)).join('')
                : '<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;">✅ Nenhum protocolo aguardando retorno.</td></tr>';
        }
        if (historico) {
            historico.innerHTML = todos.length
                ? todos.map(p => renderLinha(p, false)).join('')
                : '<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;">Nenhum protocolo enviado ainda.</td></tr>';
        }
    } catch(e) {
        if (pendentes) pendentes.innerHTML = '<tr><td colspan="5" style="padding:20px;text-align:center;color:var(--danger);">Erro ao carregar.</td></tr>';
    }
}

// --- Suprimentos: carregar todos os protocolos ---

async function carregarProtocolosSup() {
    const tbody = document.getElementById('protocolosSupBody');
    const tbodyHist = document.getElementById('protocolosSupHistBody');
    if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';
    if (tbodyHist) tbodyHist.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';
    const filtro = document.getElementById('filtroStatusProt')?.value || '';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocolos&t=${Date.now()}`));
        const data = await res.json();
        let lista = data.protocolos || [];
        if (filtro) lista = lista.filter(p => p.status === filtro);

        const pendentes = lista.filter(p => p.status === 'ENVIADO' || p.status === 'DEVOLUÇÃO PARCIAL');
        const historico = lista;

        const renderLinha = (p) => {
            const aguardando = p.status === 'ENVIADO';
            const podeDev    = p.status === 'ENVIADO' || p.status === 'RECEBIDO' || p.status === 'DEVOLUÇÃO PARCIAL';
            const btns = `
                <div style="display:flex;gap:5px;flex-wrap:nowrap;align-items:center;">
                    <button onclick="verDetalheProtocolo('${p.id}','${p.status}','${p.responsavel}','${p.dataCriacao}',true)" title="Ver notas"
                        style="background:rgba(6,182,212,0.1);border:1px solid rgba(6,182,212,0.3);border-radius:6px;padding:4px 9px;font-size:10px;font-weight:800;cursor:pointer;color:#06b6d4;white-space:nowrap;">
                        <i class="ph ph-eye"></i> Ver
                    </button>
                    ${aguardando ? `<button onclick="confirmarRecebimentoProtocolo('${p.id}')" title="Confirmar recebimento"
                        style="background:rgba(16,185,129,0.1);border:1px solid rgba(16,185,129,0.3);border-radius:6px;padding:4px 9px;font-size:10px;font-weight:800;cursor:pointer;color:#10b981;white-space:nowrap;">
                        <i class="ph ph-check-circle"></i> Confirmar
                    </button>` : ''}
                    ${podeDev ? `<button onclick="verDetalheProtocolo('${p.id}','${p.status}','${p.responsavel}','${p.dataCriacao}',true)" title="Devolver nota com problema"
                        style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);border-radius:6px;padding:4px 9px;font-size:10px;font-weight:800;cursor:pointer;color:#f87171;white-space:nowrap;">
                        <i class="ph ph-arrow-u-up-left"></i> Devolver
                    </button>` : ''}
                    <button onclick="exportarProtocoloCSVById('${p.id}','${p.responsavel}','${p.dataCriacao}')" title="Exportar CSV"
                        style="background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.2);border-radius:6px;padding:4px 9px;font-size:10px;font-weight:800;cursor:pointer;color:#818cf8;white-space:nowrap;">
                        <i class="ph ph-file-csv"></i>
                    </button>
                </div>`;
            return `<tr style="border-top:1px solid var(--border);">
                <td style="padding:11px 16px;font-weight:900;font-family:monospace;font-size:11px;color:var(--accent);">${p.id}</td>
                <td style="padding:11px 16px;font-size:11px;color:var(--text-muted);">${p.dataCriacao}</td>
                <td style="padding:11px 16px;font-size:12px;font-weight:700;">${p.responsavel}</td>
                <td style="padding:11px 16px;font-size:12px;font-weight:800;">${p.totalNotas}</td>
                <td style="padding:11px 16px;">${badgeStatusProtocolo(p.status)}</td>
                <td style="padding:11px 16px;">${btns}</td>
            </tr>`;
        };

        const vazio = (cols) => `<tr><td colspan="${cols}" style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;">Nenhum protocolo encontrado.</td></tr>`;

        if (tbody)     tbody.innerHTML     = pendentes.length ? pendentes.map(renderLinha).join('') : '<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--text-muted);font-size:13px;">✅ Nenhum protocolo aguardando conferência.</td></tr>';
        if (tbodyHist) tbodyHist.innerHTML = historico.length ? historico.map(renderLinha).join('')  : vazio(6);

    } catch(e) {
        if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--danger);">Erro ao carregar.</td></tr>';
    }
}

// --- Modal novo protocolo ---

function abrirNovoProtocolo() {
    protocoloLinhasData = [];
    document.getElementById('protResponsavel').value = usuarioAtual?.nome || '';
    document.getElementById('protObs').value = '';
    document.getElementById('protocoloErro').textContent = '';
    document.getElementById('protocoloLinhas').innerHTML = '';
    adicionarLinhaProtocolo();
    document.getElementById('modalNovoProtocolo').style.display = 'flex';
}

function fecharModalNovoProtocolo() {
    document.getElementById('modalNovoProtocolo').style.display = 'none';
}

function adicionarLinhaProtocolo() {
    const idx = Date.now();
    const tbody = document.getElementById('protocoloLinhas');
    const tr = document.createElement('tr');
    tr.setAttribute('data-idx', idx);
    tr.style.borderTop = '1px solid var(--border)';
    tr.innerHTML = `
        <td style="padding:6px 8px;"><input type="date" class="prot-input prot-data" value="${new Date().toISOString().split('T')[0]}" style="width:130px;"></td>
        <td style="padding:6px 8px;"><input type="text" class="prot-input prot-empresa" placeholder="Ex: BOSTON" style="width:130px;text-transform:uppercase;"></td>
        <td style="padding:6px 8px;"><input type="text" class="prot-input prot-numero" placeholder="Ex: 3427488" style="width:90px;"></td>
        <td style="padding:6px 8px;"><input type="text" class="prot-input prot-chave" placeholder="44 dígitos" maxlength="44" style="width:200px;font-family:monospace;font-size:11px;"></td>
        <td style="padding:6px 8px;">
            <select class="prot-input prot-nat" style="width:110px;">
                <option value="HC">HC</option>
                <option value="REMESSA">REMESSA</option>
                <option value="FZ">FZ</option>
                <option value="OUTRO">OUTRO</option>
            </select>
        </td>
        <td style="padding:6px 8px;text-align:center;">
            <input type="checkbox" class="prot-lote" style="width:18px;height:18px;cursor:pointer;">
        </td>
        <td style="padding:6px 8px;text-align:center;">
            <button onclick="removerLinhaProtocolo(this)" style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.2);border-radius:6px;padding:4px 8px;cursor:pointer;color:#f87171;">
                <i class="ph ph-trash"></i>
            </button>
        </td>`;
    tbody.appendChild(tr);
}

function removerLinhaProtocolo(btn) {
    btn.closest('tr').remove();
}

function coletarLinhasProtocolo() {
    const linhas = document.querySelectorAll('#protocoloLinhas tr');
    return Array.from(linhas).map(tr => ({
        data:          tr.querySelector('.prot-data')?.value || '',
        empresa:       (tr.querySelector('.prot-empresa')?.value || '').toUpperCase().trim(),
        numero_nota:   tr.querySelector('.prot-numero')?.value?.trim() || '',
        chave_acesso:  tr.querySelector('.prot-chave')?.value?.trim() || '',
        nat_operacao:  tr.querySelector('.prot-nat')?.value || 'HC',
        tem_lote:      tr.querySelector('.prot-lote')?.checked || false
    })).filter(n => n.empresa || n.numero_nota);
}

async function enviarProtocolo() {
    const notas = coletarLinhasProtocolo();
    const erro  = document.getElementById('protocoloErro');
    if (!notas.length) { erro.textContent = 'Adicione pelo menos uma nota.'; return; }
    const semNota = notas.find(n => !n.numero_nota);
    if (semNota) { erro.textContent = 'Preencha o número da nota em todas as linhas.'; return; }
    erro.textContent = '';

    try {
        const res  = await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({
                tipo: 'criarProtocolo',
                responsavel: usuarioAtual?.nome || loginAtual,
                login: loginAtual,
                obs: document.getElementById('protObs').value,
                notas
            })
        });
        document.getElementById('modalNovoProtocolo').style.display = 'none';
        _cache.protocolosOpme = false;
        await carregarMeusProtocolos();
        mostrarToast('✅ Protocolo enviado ao Suprimentos!', 'success', 4000);
    } catch(e) {
        erro.textContent = 'Erro ao enviar. Tente novamente.';
    }
}

// --- CSV Export ---

function exportarProtocoloCSV() {
    const notas = coletarLinhasProtocolo();
    if (!notas.length) { mostrarToast('Adicione notas antes de exportar.', 'error'); return; }
    const responsavel = document.getElementById('protResponsavel').value;
    gerarCSV(notas, responsavel, 'rascunho');
}

async function exportarProtocoloCSVById(id, responsavel, data) {
    mostrarToast('⏳ Preparando CSV...', 'info', 2000);
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocoloNotas&protocolo_id=${id}&t=${Date.now()}`));
        const data2 = await res.json();
        gerarCSV(data2.notas || [], responsavel, id);
    } catch(e) { mostrarToast('Erro ao gerar CSV', 'error'); }
}

function gerarCSV(notas, responsavel, remessa) {
    const XLSX = window.XLSX;
    if (!XLSX) { mostrarToast('Biblioteca não carregada, recarregue a página.', 'error'); return; }

    const borda = {
        top:    { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'thin', color: { rgb: '000000' } },
        left:   { style: 'thin', color: { rgb: '000000' } },
        right:  { style: 'thin', color: { rgb: '000000' } }
    };

    const wsData = [['DATA','EMPRESA','Nº DA NOTA','CHAVE DE ACESSO','NAT. OPERAÇÃO','REMESSA']];
    notas.forEach(n => wsData.push([
        n.data || new Date().toLocaleDateString('pt-BR'),
        (n.empresa||'').toUpperCase(),
        n.numero_nota || '',
        n.chave_acesso || '',
        n.nat_operacao || '',
        remessa
    ]));

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{wch:12},{wch:20},{wch:12},{wch:48},{wch:14},{wch:10}];

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const addr = XLSX.utils.encode_cell({r:R,c:C});
            if (!ws[addr]) ws[addr] = {v:'',t:'s'};
            ws[addr].s = {
                font:      R === 0 ? {bold:true, sz:10, name:'Calibri'} : {sz:10, name:'Calibri'},
                alignment: {horizontal: C === 3 ? 'left' : 'center', vertical:'center'},
                border:    borda
            };
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Remessa ${remessa}`);
    XLSX.writeFile(wb, `remessa_${remessa}.xlsx`);
    mostrarToast('✅ Planilha exportada!', 'success');
}

// --- Modal detalhe do protocolo ---

async function verDetalheProtocolo(id, status, responsavel, dataCriacao, isSup) {
    protocoloAtualId = id;
    document.getElementById('detalheProtId').textContent  = id;
    document.getElementById('detalheProtMeta').textContent = `${responsavel} · ${dataCriacao} · ${status}`;
    document.getElementById('detalheNotasBody').innerHTML = '<tr><td colspan="7" style="padding:16px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';
    document.getElementById('detalheAcoes').innerHTML = '';
    document.getElementById('colDevolver').style.display = 'none';
    document.getElementById('modalDetalheProtocolo').style.display = 'flex';

    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocoloNotas&protocolo_id=${id}&t=${Date.now()}`));
        const data = await res.json();
        const notas = data.notas || [];

        const podeDevolverLinha = isSup && (status === 'ENVIADO' || status === 'RECEBIDO' || status === 'DEVOLUÇÃO PARCIAL');
        if (podeDevolverLinha) document.getElementById('colDevolver').style.display = 'table-cell';

        document.getElementById('detalheNotasBody').innerHTML = notas.map(n => `
            <tr style="border-top:1px solid var(--border);" data-linha="${n.linha}">
                <td style="padding:10px 12px;font-size:12px;">${n.empresa}</td>
                <td style="padding:10px 12px;font-family:monospace;font-size:12px;">${n.numero_nota}</td>
                <td style="padding:10px 12px;font-family:monospace;font-size:10px;max-width:160px;overflow:hidden;text-overflow:ellipsis;" title="${n.chave_acesso}">${n.chave_acesso}</td>
                <td style="padding:10px 12px;font-size:12px;">${n.nat_operacao}</td>
                <td style="padding:10px 12px;text-align:center;font-size:12px;">${n.tem_lote}</td>
                <td style="padding:10px 12px;text-align:center;">
                    ${n.status === 'DEVOLVIDA'
                        ? `<span style="color:#ef4444;font-size:10px;font-weight:800;">DEVOLVIDA</span><br><small style="color:var(--text-muted);font-size:10px;">${n.obs_devolucao}</small>`
                        : '<span style="color:#10b981;font-size:10px;font-weight:800;">NORMAL</span>'}
                </td>
                ${podeDevolverLinha ? `<td style="padding:10px 12px;text-align:center;">
                    ${n.status !== 'DEVOLVIDA' ? `<input type="checkbox" class="chk-devolver" data-linha="${n.linha}" style="width:16px;height:16px;cursor:pointer;">` : '—'}
                </td>` : ''}
            </tr>`).join('');

        // Botões de ação
        const acoes = document.getElementById('detalheAcoes');
        if (isSup && status === 'ENVIADO') {
            acoes.innerHTML = `
                <button onclick="confirmarRecebimentoProtocolo('${id}')" class="btn-novo-usuario" style="flex:1;justify-content:center;">
                    <i class="ph ph-check-circle"></i> CONFIRMAR RECEBIMENTO
                </button>
                <button onclick="devolverNotasSelecionadas('${id}')"
                    style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);border-radius:8px;padding:12px 18px;font-size:11px;font-weight:800;cursor:pointer;color:#f87171;display:flex;align-items:center;gap:8px;">
                    <i class="ph ph-arrow-u-up-left"></i> DEVOLVER NOTAS SELECIONADAS
                </button>`;
        } else if (isSup && (status === 'RECEBIDO' || status === 'DEVOLUÇÃO PARCIAL')) {
            acoes.innerHTML = `
                <button onclick="devolverNotasSelecionadas('${id}')"
                    style="background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.3);border-radius:8px;padding:12px 18px;font-size:11px;font-weight:800;cursor:pointer;color:#f87171;display:flex;align-items:center;gap:8px;">
                    <i class="ph ph-arrow-u-up-left"></i> DEVOLVER NOTAS SELECIONADAS
                </button>`;
        }

        // Botão CSV sempre disponível
        acoes.innerHTML += `
            <button onclick="exportarProtocoloCSVById('${id}','${responsavel}','${dataCriacao}')"
                style="background:rgba(99,102,241,0.1);border:1px solid rgba(99,102,241,0.3);border-radius:8px;padding:12px 18px;font-size:11px;font-weight:800;cursor:pointer;color:#818cf8;display:flex;align-items:center;gap:8px;">
                <i class="ph ph-file-csv"></i> EXPORTAR CSV
            </button>`;

    } catch(e) {
        document.getElementById('detalheNotasBody').innerHTML = '<tr><td colspan="7" style="padding:16px;text-align:center;color:var(--danger);">Erro ao carregar notas.</td></tr>';
    }
}

function fecharModalDetalhe() {
    document.getElementById('modalDetalheProtocolo').style.display = 'none';
}

// --- Ações do Suprimentos ---

async function confirmarRecebimentoProtocolo(id) {
    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'confirmarRecebimentoProtocolo', protocolo_id: id, login: loginAtual, responsavel: usuarioAtual?.nome })
        });
        document.getElementById('modalDetalheProtocolo').style.display = 'none';
        _cache.protocolosSup = false;
        await carregarProtocolosSup();
        mostrarToast('✅ Recebimento confirmado!', 'success');
    } catch(e) { mostrarToast('Erro ao confirmar', 'error'); }
}

async function devolverNotasSelecionadas(id) {
    const checks = document.querySelectorAll('.chk-devolver:checked');
    if (!checks.length) { mostrarToast('Selecione pelo menos uma nota para devolver.', 'error'); return; }
    const obs = prompt('Motivo da devolução (obrigatório):');
    if (!obs) return;
    const notas = Array.from(checks).map(c => ({ linha: parseInt(c.dataset.linha), obs_devolucao: obs }));
    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'devolverNotas', protocolo_id: id, login: loginAtual, responsavel: usuarioAtual?.nome, notas })
        });
        document.getElementById('modalDetalheProtocolo').style.display = 'none';
        _cache.protocolosSup = false;
        await carregarProtocolosSup();
        mostrarToast('✅ Devolução registrada — OPME será notificado.', 'success');
    } catch(e) { mostrarToast('Erro ao registrar devolução', 'error'); }
}

async function confirmarDevolucao(id) {
    try {
        await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'confirmarDevolucao', protocolo_id: id, login: loginAtual, responsavel: usuarioAtual?.nome })
        });
        _cache.protocolosOpme = false;
        await carregarMeusProtocolos();
        mostrarToast('✅ Devolução confirmada.', 'success');
    } catch(e) { mostrarToast('Erro ao confirmar devolução', 'error'); }
}

// Integra inputs CSS do protocolo
const _initNavOrig = iniciarNavAnimacoes;

// =============================================================================
// POLLING EM TEMPO REAL — verifica notificações e protocolos a cada 30s
// =============================================================================

let _pollingInterval   = null;
let _ultimasNotifIds   = new Set(); // IDs já vistos para detectar novas
let _ultimoStatusProts = {};        // { id: status } para detectar mudanças

async function iniciarPolling() {
    if (_pollingInterval) return; // já está rodando
    // Primeira rodada imediata
    await _pollingTick();
    // Repete a cada 30 segundos
    _pollingInterval = setInterval(_pollingTick, 30000);
}

function pararPolling() {
    if (_pollingInterval) { clearInterval(_pollingInterval); _pollingInterval = null; }
}

async function _pollingTick() {
    if (!loginAtual || !sessaoAtual) return; // não logado
    try {
        await _verificarNotificacoes();
        await _verificarStatusProtocolos();
    } catch(e) { /* silencioso — não interrompe o usuário */ }
}

async function _verificarNotificacoes() {
    const isAdmin   = temPermissao('administrador');
    const comprador = isAdmin ? '' : (compradorProjecaoAtual || loginAtual.split(' ')[0].toUpperCase());
    const url = comprador
        ? addAuth(`${URL_SCRIPT}?action=getNotificacoes&comprador=${encodeURIComponent(comprador)}&t=${Date.now()}`)
        : addAuth(`${URL_SCRIPT}?action=getNotificacoes&t=${Date.now()}`);

    const resp = await fetch(url);
    const data = await resp.json();
    const novas = (data.notificacoes || []).filter(n => !_ultimasNotifIds.has(n.linha));

    if (novas.length > 0) {
        // Atualiza lista global
        notificacoesPendentes = data.notificacoes || [];
        atualizarSinoNotificacoes();

        // Marca IDs já vistos
        notificacoesPendentes.forEach(n => _ultimasNotifIds.add(n.linha));

        // Toast para cada notificação nova
        novas.forEach(n => {
            const icone = n.tipo === 'PROTOCOLO_NOVO'     ? '📋' :
                          n.tipo === 'PROTOCOLO_RECEBIDO' ? '✅' :
                          n.tipo === 'PROTOCOLO_DEVOLUCAO'? '↩️' : '🔔';
            mostrarToast(`${icone} ${n.mensagem.substring(0, 80)}${n.mensagem.length > 80 ? '…' : ''}`, 'info', 6000);
        });
    } else {
        // Sem novas, mas atualiza lista silenciosamente
        notificacoesPendentes = data.notificacoes || [];
        atualizarSinoNotificacoes();
    }
}

async function _verificarStatusProtocolos() {
    // Atualiza as tabelas de protocolo se a aba estiver aberta
    if (abaAtual === 'ProtocolosOpme') {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocolos&filtro=meus&t=${Date.now()}`));
        const data = await res.json();
        const lista = data.protocolos || [];

        // Detecta mudanças de status
        let houveAlteracao = false;
        lista.forEach(p => {
            if (_ultimoStatusProts[p.id] && _ultimoStatusProts[p.id] !== p.status) {
                houveAlteracao = true;
                const msg = p.status === 'RECEBIDO'
                    ? `✅ Remessa ${p.id} confirmada pelo Suprimentos`
                    : p.status === 'DEVOLUÇÃO PARCIAL'
                    ? `↩️ Remessa ${p.id} tem notas devolvidas — verifique!`
                    : `🔔 Remessa ${p.id}: status atualizado para ${p.status}`;
                mostrarToast(msg, p.status === 'RECEBIDO' ? 'success' : 'warning', 7000);
            }
            _ultimoStatusProts[p.id] = p.status;
        });

        if (houveAlteracao) {
            // Recarrega a tabela silenciosamente
            _cache.protocolosOpme = false;
            await carregarMeusProtocolos();
        }
    }

    if (abaAtual === 'ProtocolosSup') {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getProtocolos&t=${Date.now()}`));
        const data = await res.json();
        const lista = data.protocolos || [];

        let houveAlteracao = false;
        lista.forEach(p => {
            if (_ultimoStatusProts['sup_' + p.id] === undefined) {
                // Primeiro load — só registra
                _ultimoStatusProts['sup_' + p.id] = p.status;
            } else if (_ultimoStatusProts['sup_' + p.id] !== p.status) {
                houveAlteracao = true;
                _ultimoStatusProts['sup_' + p.id] = p.status;
            }
            // Novo protocolo ENVIADO que não estava antes
            if (!_ultimoStatusProts['sup_' + p.id + '_visto']) {
                _ultimoStatusProts['sup_' + p.id + '_visto'] = true;
                if (p.status === 'ENVIADO') {
                    mostrarToast(`📋 Nova remessa ${p.id} aguardando conferência`, 'info', 6000);
                    houveAlteracao = true;
                }
            }
        });

        if (houveAlteracao) {
            _cache.protocolosSup = false;
            await carregarProtocolosSup();
        }
    }
}
// =============================================================================
// ADMIN OPME — Gestão de usuários do módulo Especiais
// =============================================================================

function switchOpmeAdminTab(tab) {
    document.getElementById('opmeAdminTabUsuarios').style.display = tab === 'usuarios' ? 'block' : 'none';
    document.getElementById('opmeAdminTabLog').style.display      = tab === 'log'      ? 'block' : 'none';
    document.getElementById('opmeAdminTabSistema').style.display  = tab === 'sistema'  ? 'block' : 'none';
    document.getElementById('opmeTabBtnUsuarios').classList.toggle('ativo', tab === 'usuarios');
    document.getElementById('opmeTabBtnLog').classList.toggle('ativo',      tab === 'log');
    document.getElementById('opmeTabBtnSistema').classList.toggle('ativo',  tab === 'sistema');
    if (tab === 'log') carregarLogOpme();
}

async function carregarLogOpme() {
    // Reutiliza o mesmo endpoint de log mas filtra logins esp
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getLog&t=${Date.now()}`));
        const data = await res.json();
        const tbody = document.querySelector('#tabelaLogOpme tbody');
        const logs  = (data.logs || []).filter(l => l.login && l.login.toLowerCase().startsWith('esp'));
        tbody.innerHTML = logs.length
            ? logs.map(l => `<tr>
                <td>${l.data}</td><td>${l.login}</td><td>${l.acao}</td><td>${l.detalhe||''}</td>
              </tr>`).join('')
            : '<tr><td colspan="4" style="padding:16px;text-align:center;color:var(--text-muted);">Nenhum log OPME encontrado.</td></tr>';
    } catch(e) {}
}
function abrirModalNovoUsuarioOpme() {
    document.getElementById('opme_nome').value = '';
    document.getElementById('opme_login').value = 'esp';
    document.getElementById('opme_email').value = '';
    document.getElementById('opmeUsuarioErro').textContent = '';
    document.getElementById('modalNovoUsuarioOpme').style.display = 'flex';
}

async function salvarNovoUsuarioOpme() {
    const nome  = document.getElementById('opme_nome').value.trim();
    const login = document.getElementById('opme_login').value.trim().toLowerCase();
    const email = document.getElementById('opme_email').value.trim();
    const erro  = document.getElementById('opmeUsuarioErro');

    if (!nome)  { erro.textContent = 'Preencha o nome.'; return; }
    if (!login.startsWith('esp')) { erro.textContent = 'Login deve começar com "esp".'; return; }
    if (login.length < 4) { erro.textContent = 'Login muito curto.'; return; }

    erro.textContent = '';
    try {
        const url = addAuth(`${URL_SCRIPT}?action=criarUsuario&solicitante=${encodeURIComponent(loginAtual)}&login=${encodeURIComponent(login)}&nome=${encodeURIComponent(nome)}&permissoes=gestor-opme&prefixos=&email=${encodeURIComponent(email)}&senha=Core%4026`);
        const res  = await fetch(url);
        const data = await res.json();
        if (!data.ok) throw new Error(data.erro || 'Erro');
        fecharModal('modalNovoUsuarioOpme');
        await carregarUsuariosOpme();
        mostrarToast('✅ Usuário OPME criado! Senha padrão: Core@26', 'success', 5000);
    } catch(e) {
        erro.textContent = 'Erro: ' + e.message;
    }
}

async function carregarUsuariosOpme() {
    const tbody = document.getElementById('tabelaUsuariosOpme');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="5" style="padding:24px;text-align:center;color:var(--text-muted);">Carregando...</td></tr>';
    try {
        const res  = await fetch(addAuth(`${URL_SCRIPT}?action=getUsuarios`));
        const lista = await res.json();
        const opme  = lista.filter(u => u.login && u.login.toLowerCase().startsWith('esp'));
        if (!opme.length) {
            tbody.innerHTML = '<tr><td colspan="5" style="padding:24px;text-align:center;color:var(--text-muted);">Nenhum usuário OPME cadastrado.</td></tr>';
            return;
        }
        tbody.innerHTML = opme.map(u => {
            const ativo    = u.status !== 'primeiroAcesso';
            const permsEsc = (u.permissoes||'').replace(/'/g,"\\'");
            const nomeEsc  = (u.nome||'').replace(/'/g,"\\'");
            const emailEsc = (u.email||'').replace(/'/g,"\\'");
            const status   = ativo
                ? '<span class="badge-ativo">✓ Ativo</span>'
                : '<span class="badge-primeiro-acesso">⏳ Aguardando 1º acesso</span>';
            return `<tr>
                <td><b>${u.nome}</b><br><span style="font-size:11px;color:var(--text-muted)">${u.login}</span></td>
                <td><span style="background:rgba(245,158,11,0.12);color:#f59e0b;font-size:10px;font-weight:800;padding:3px 9px;border-radius:6px;display:inline-block;">GESTOR-OPME</span></td>
                <td>${status}</td>
                <td style="text-align:center;">
                    <div class="admin-acoes">
                        <button class="btn-admin-acao ver-como" onclick="verComoUsuario('${u.login}','${nomeEsc}');" title="Ver como"><i class="ph ph-eye"></i></button>
                        <button class="btn-admin-acao editar"  onclick="abrirModalUsuario('${u.login}','${nomeEsc}','${permsEsc}','','${emailEsc}')" title="Editar"><i class="ph ph-pencil-simple"></i></button>
                        <button class="btn-admin-acao resetar" onclick="abrirModalResetarSenha('${u.login}','${nomeEsc}');" title="Resetar senha"><i class="ph ph-key"></i></button>
                        <button class="btn-admin-acao deletar" onclick="abrirModalDeletar('${u.login}','${nomeEsc}');" title="Remover"><i class="ph ph-trash"></i></button>
                    </div>
                </td>
            </tr>`;
        }).join('');
    } catch(e) {
        tbody.innerHTML = '<tr><td colspan="5" style="padding:24px;text-align:center;color:var(--danger);">Erro ao carregar.</td></tr>';
    }
}

async function migrarModulos() {
    mostrarToast('⏳ Preenchendo coluna módulo...', 'info', 3000);
    try {
        await fetch(addAuth(`${URL_SCRIPT}?action=migrarModulos`));
        mostrarToast('✅ Coluna módulo preenchida com sucesso!', 'success');
    } catch(e) {
        mostrarToast('Erro ao executar migração', 'error');
    }
}

// =============================================================================
// PROJEÇÃO OPME — Funções Frontend
// Visível apenas para usuários esp* (modo OPME)
// =============================================================================

var opmeData     = null;
var opmeTabAtiva = 'c2';
var opmeDescs    = {
    c2: 'SEM RP · Cobertura &lt; 45 dias · SEM empenho ativo',
    c4: 'COM RP · Cobertura &lt; 20 dias · SEM empenho — verificar o que a última COP trouxe',
    c5: 'COM RP · Cobertura &lt; 20 dias · COM empenho ativo',
    c6: 'Consumo médio mensal &lt; 2 unidades · Cobertura &gt; 45 dias (atenção especial)'
};

async function rodarProjecaoOPME() {
    // Oculta estado inicial, mostra loading
    document.getElementById('opme-initial-state').style.display  = 'none';
    document.getElementById('opme-loading-state').style.display  = 'flex';
    document.getElementById('opme-cards-wrap').style.display     = 'none';
    document.getElementById('opme-tabs-bar').style.display       = 'none';
    ['c2','c4','c5','c6','hist','cop'].forEach(t =>
        document.getElementById('opme-panel-' + t).style.display = 'none'
    );

    try {
        const res = await fetch(addAuth(URL_SCRIPT + '?action=projecaoOPME'));

        // Detecta resposta não-JSON (redirect, erro de autenticação, etc.)
        const ct = res.headers.get('content-type') || '';
        if (!ct.includes('json') && !ct.includes('plain')) {
            throw new Error('O servidor não retornou JSON. Verifique se o backend foi republicado com nova versão.');
        }

        let data;
        try { data = await res.json(); }
        catch (pe) { throw new Error('Resposta inválida do servidor. Publique uma nova versão no Apps Script.'); }

        // Erros explícitos do backend
        if (data && data.erro) throw new Error('Erro no backend: ' + data.erro);

        // Estrutura incompleta — backend antigo ou não republicado
        if (!data || !data.totais || !data.criterio2) {
            throw new Error(
                'Resposta incompleta. Causas comuns:\n' +
                '• Backend não foi republicado após adicionar opme_projecao.gs\n' +
                '• Aba "Alerta_Coop" não encontrada na planilha COP\n' +
                '• ID da planilha incorreto no opme_projecao.gs'
            );
        }

        opmeData = data;

        // Atualiza cards
        ['c2','c4','c5','c6'].forEach(k => {
            var numEl   = document.getElementById('opme-num-'   + k);
            var badgeEl = document.getElementById('opme-badge-' + k);
            if (numEl)   numEl.textContent   = data.totais[k] ?? 0;
            if (badgeEl) badgeEl.textContent = data.totais[k] ?? 0;
        });
        // card FORA COP
        var _bloqTotal = (data.bloqueados || []).length;
        var _nb1 = document.getElementById('opme-num-bloq');
        var _nb2 = document.getElementById('opme-badge-bloq');
        if (_nb1) _nb1.textContent = _bloqTotal;
        if (_nb2) _nb2.textContent = _bloqTotal;

        // Renderiza tabelas
        _renderTblOPME('c2', data.criterio2 || []);
        _renderTblOPME('c4', data.criterio4 || []);
        _renderTblOPME('c5', data.criterio5 || []);
        _renderTblOPME('c6', data.criterio6 || []);
        _renderTblOPME('bloq', data.bloqueados || []);

        document.getElementById('opme-loading-state').style.display = 'none';
        document.getElementById('opme-cards-wrap').style.display    = 'grid';
        document.getElementById('opme-tabs-bar').style.display      = 'flex';
        document.getElementById('btn-salvar-cop-opme').disabled     = false;
        switchOPMETab('c2');

    } catch(err) {
        document.getElementById('opme-loading-state').style.display = 'none';
        document.getElementById('opme-initial-state').style.display = 'block';
        document.getElementById('opme-initial-state').innerHTML =
            '<div class="opme-empty-icon">❌</div>' +
            '<p style="font-weight:700;color:var(--danger);">' + err.message + '</p>' +
            '<button class="opme-btn opme-btn-run" style="margin-top:16px" onclick="rodarProjecaoOPME()">Tentar novamente</button>';
    }
}

function _renderTblOPME(criterio, itens) {
    var panel = document.getElementById('opme-panel-' + criterio);
    var desc  = opmeDescs[criterio] || '';
    var isc5  = criterio === 'c5';

    if (!itens || !itens.length) {
        panel.innerHTML =
            '<div class="opme-empty-state">' +
            '<div class="opme-empty-icon">✅</div>' +
            '<p style="font-weight:700;color:var(--text-muted);">Nenhum item neste critério</p>' +
            '<p style="font-size:11px;color:var(--text-muted);margin-top:6px;">' + desc + '</p></div>';
        return;
    }

    var html =
        '<div style="padding:0 0 10px">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap;gap:8px;">' +
        '<div style="font-size:11px;color:var(--text-muted);">' + desc + '</div>' +
        '<input id="opme-filt-' + criterio + '" type="text" placeholder="🔍 Filtrar código ou descrição..." ' +
        'oninput="_filtrarOPME(\'' + criterio + '\')" ' +
        'style="background:var(--card-bg);border:1px solid var(--border);border-radius:8px;padding:6px 12px;' +
        'color:var(--text-main);font-size:12px;outline:none;min-width:220px;">' +
        '</div>' +
        '<div class="opme-tbl-wrap">' +
        '<table class="opme-tbl" id="opme-tbl-' + criterio + '">' +
        '<thead><tr>' +
        '<th onclick="_sortOPME(\'' + criterio + '\',\'codigo\')">Código ↕</th>' +
        '<th onclick="_sortOPME(\'' + criterio + '\',\'descricao\')">Descrição ↕</th>' +
        '<th onclick="_sortOPME(\'' + criterio + '\',\'cobertura\')">Cob. (dias) ↕</th>' +
        '<th onclick="_sortOPME(\'' + criterio + '\',\'cmmMensal\')">CMM/mês ↕</th>' +
        '<th>CMD/dia</th><th>RP</th><th>Empenho</th><th>COP</th>' +
        (isc5 ? '<th onclick="_sortOPME(\'' + criterio + '\',\'entregaPendente\')">Entrega Pend. ↕</th><th>Nº Empenho(s)</th>' : '') +
        '</tr></thead>' +
        '<tbody id="opme-tbody-' + criterio + '">' +
        itens.map(function(it){ return _linhaOPME(it, isc5); }).join('') +
        '</tbody></table></div>' +
        '<div style="text-align:right;padding:6px 4px;font-size:11px;color:var(--text-muted);">' +
        '<span id="opme-count-' + criterio + '">' + itens.length + '</span> item(s)' +
        '</div></div>';
    panel.innerHTML = html;
    panel.dataset.itens = JSON.stringify(itens);
}

function _linhaOPME(it, mostrarEmp) {
    var cob = it.cobertura || 0;
    var cobCls = cob === 0 ? 'cob-zero' : cob < 10 ? 'cob-crit' : cob < 30 ? 'cob-warn' : 'cob-ok';
    var cobTxt = cob === 0 ? '⚠ ZERO' : cob + ' d';
    var rpB  = it.temRP
        ? '<span class="bd-rp">COM RP</span>'
        : '<span class="bd-norp">SEM RP</span>';
    var empB = it.temEmpenho
        ? '<span class="bd-emp">ATIVO</span>'
        : '<span class="bd-noemp">NENHUM</span>';

    // Badge COP: bloqueioCompra = true → NÃO compra via COP (valor "N" no Alerta_Coop)
    var copB = it.bloqueioCompra
        ? '<span style="background:#ef4444;color:#fff;font-size:10px;font-weight:800;padding:2px 8px;border-radius:6px;letter-spacing:.5px;">FORA COP</span>'
        : '<span style="background:#16a34a;color:#fff;font-size:10px;font-weight:800;padding:2px 8px;border-radius:6px;letter-spacing:.5px;">VIA COP</span>';

    // Linha inteira destacada em vermelho claro quando fora da COP
    var rowStyle = it.bloqueioCompra
        ? ' style="background:rgba(239,68,68,0.08);border-left:3px solid #ef4444;"'
        : '';

    return '<tr data-cod="' + it.codigo + '" data-desc="' + (it.descricao||'').toLowerCase() + '"' + rowStyle + '>' +
        '<td style="font-family:monospace;font-size:11px;font-weight:700;color:var(--accent)">' + it.codigo + '</td>' +
        '<td style="max-width:300px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + (it.descricao||'') + '">' + (it.descricao||'') + '</td>' +
        '<td class="' + cobCls + '">' + cobTxt + '</td>' +
        '<td style="color:var(--text-muted)">' + (it.cmmMensal > 0 ? it.cmmMensal.toFixed(2) : '—') + '</td>' +
        '<td style="color:var(--text-muted);font-size:11px">' + (it.cmdDiario > 0 ? it.cmdDiario.toFixed(4) : '—') + '</td>' +
        '<td>' + rpB + '</td>' +
        '<td>' + empB + '</td>' +
        '<td>' + copB + '</td>' +
        (mostrarEmp
            ? '<td style="color:#f59e0b;font-weight:700">' + (it.entregaPendente||0) + '</td>' +
              '<td style="font-size:11px;color:var(--text-muted)">' + ((it.numerosEmpenho||[]).join(', ')||'—') + '</td>'
            : '') +
        '</tr>';
}

function _filtrarOPME(criterio) {
    var q   = document.getElementById('opme-filt-' + criterio).value.toLowerCase();
    var rows = document.querySelectorAll('#opme-tbody-' + criterio + ' tr');
    var n   = 0;
    rows.forEach(function(tr){
        var ok = !q || tr.dataset.cod.toLowerCase().includes(q) || tr.dataset.desc.includes(q);
        tr.style.display = ok ? '' : 'none';
        if (ok) n++;
    });
    var el = document.getElementById('opme-count-' + criterio);
    if (el) el.textContent = n;
}

var _opme_sort_dir = {};
function _sortOPME(criterio, campo) {
    var key = criterio + '_' + campo;
    var asc = _opme_sort_dir[key] !== true;
    _opme_sort_dir[key] = asc;
    var panel = document.getElementById('opme-panel-' + criterio);
    var itens = JSON.parse(panel.dataset.itens || '[]');
    itens.sort(function(a,b){
        var va = a[campo], vb = b[campo];
        if (typeof va === 'number') return asc ? va-vb : vb-va;
        return asc
            ? (va||'').toString().localeCompare((vb||'').toString())
            : (vb||'').toString().localeCompare((va||'').toString());
    });
    var isc5 = criterio === 'c5';
    document.getElementById('opme-tbody-' + criterio).innerHTML =
        itens.map(function(it){ return _linhaOPME(it, isc5); }).join('');
}

function switchOPMETab(tab) {
    opmeTabAtiva = tab;
    ['c2','c4','c5','c6','bloq','hist','cop'].forEach(function(t){
        var btn = document.getElementById('opme-tbtn-' + t);
        var pnl = document.getElementById('opme-panel-' + t);
        if (btn) btn.classList.toggle('opme-tab-active', t === tab);
        if (pnl) pnl.style.display = t === tab ? 'block' : 'none';
    });
    ['c2','c4','c5','c6','bloq'].forEach(function(t){
        var c = document.getElementById('opme-card-' + t);
        if (c) c.classList.toggle('card-active', t === tab);
    });
    if (tab === 'hist') carregarHistoricoCOP();
}

async function salvarCopOPME() {
    if (!opmeData) return;
    var btn = document.getElementById('btn-salvar-cop-opme');
    btn.disabled = true;
    btn.textContent = '⏳ Salvando...';
    try {
        var todos = [].concat(opmeData.criterio2, opmeData.criterio4, opmeData.criterio5, opmeData.criterio6);
        var res  = await fetch(URL_SCRIPT, {
            method: 'POST', mode: 'no-cors',
            body: JSON.stringify({ tipo: 'salvarCopOPME', login: loginAtual, itens: todos })
        });
        btn.textContent = '✅ Salvo!';
        btn.style.background = '#16a34a';
        mostrarToast('COP salva com sucesso — ' + todos.length + ' item(s).', 'success');
        setTimeout(function(){
            btn.textContent = '💾 Salvar COP';
            btn.style.background = '';
            btn.disabled = false;
        }, 4000);
    } catch(err) {
        btn.textContent = '❌ Erro';
        btn.style.background = 'var(--danger)';
        setTimeout(function(){
            btn.textContent = '💾 Salvar COP';
            btn.style.background = '';
            btn.disabled = false;
        }, 3000);
        mostrarToast('Erro ao salvar COP.', 'error');
    }
}

async function carregarHistoricoCOP() {
    var cont = document.getElementById('opme-hist-content');
    cont.innerHTML = '<div class="opme-spinner-wrap"><div class="opme-spin"></div></div>';
    try {
        var res  = await fetch(addAuth(URL_SCRIPT + '?action=getHistoricoCOP'));
        var data = await res.json();
        if (!data.historico || !data.historico.length) {
            cont.innerHTML =
                '<div class="opme-empty-state">' +
                '<div class="opme-empty-icon">📭</div>' +
                '<p style="font-weight:700;color:var(--text-muted)">Nenhuma COP salva ainda</p>' +
                '<p style="font-size:11px;color:var(--text-muted);margin-top:6px">Rode uma projeção e clique em "Salvar COP"</p></div>';
            return;
        }
        var cores = { C2:'#ef4444', C4:'#f59e0b', C5:'#f97316', C6:'#8b5cf6' };
        var html  = '';
        data.historico.forEach(function(cop, idx){
            var byC = {};
            cop.itens.forEach(function(it){ byC[it.criterio] = (byC[it.criterio]||0)+1; });
            var tags = Object.keys(byC).map(function(k){
                return '<span style="background:' + (cores[k]||'#94a3b8') + ';color:#fff;padding:1px 8px;border-radius:999px;font-size:10px;font-weight:800;margin-right:4px">' + k + ': ' + byC[k] + '</span>';
            }).join('');
            html +=
                '<div class="opme-hist-item">' +
                '<div class="opme-hist-hdr" onclick="_toggleHistOPME(' + idx + ')">' +
                '<div>' +
                '<div class="opme-hist-date">📋 COP — ' + cop.data + '</div>' +
                '<div class="opme-hist-meta">Por: ' + (cop.login||'—') + ' &nbsp;·&nbsp; ' + cop.itens.length + ' item(s) &nbsp;·&nbsp; ' + tags + '</div>' +
                '</div>' +
                '<div style="color:var(--text-muted);font-size:18px" id="opme-harrow-' + idx + '">▸</div>' +
                '</div>' +
                '<div class="opme-hist-body" id="opme-hbody-' + idx + '">' +
                '<div class="opme-tbl-wrap">' +
                '<table class="opme-tbl"><thead><tr>' +
                '<th>Critério</th><th>Código</th><th>Descrição</th><th>Cobertura</th><th>CMM</th><th>RP</th><th>Empenho</th>' +
                '</tr></thead><tbody>' +
                cop.itens.map(function(it){
                    var c = it.cobertura||0;
                    var cor = c<10?'var(--danger)':c<30?'var(--warning)':'var(--text-muted)';
                    return '<tr>' +
                        '<td><span style="background:' + (cores[it.criterio]||'#94a3b8') + ';color:#fff;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:800">' + it.criterio + '</span></td>' +
                        '<td style="font-family:monospace;font-size:11px;font-weight:700;color:var(--accent)">' + it.codigo + '</td>' +
                        '<td style="max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + (it.descricao||'') + '</td>' +
                        '<td style="color:' + cor + ';font-weight:700">' + c + ' d</td>' +
                        '<td style="color:var(--text-muted)">' + (it.cmmMensal||'—') + '</td>' +
                        '<td>' + (it.temRP==='SIM'?'<span class="bd-rp">SIM</span>':'<span class="bd-norp">NÃO</span>') + '</td>' +
                        '<td>' + (it.temEmpenho==='SIM'?'<span class="bd-emp">SIM</span>':'<span class="bd-noemp">NÃO</span>') + '</td>' +
                        '</tr>';
                }).join('') +
                '</tbody></table></div></div></div>';
        });
        cont.innerHTML = html;
    } catch(err) {
        cont.innerHTML = '<div class="opme-empty-state" style="color:var(--danger)">Erro: ' + err.message + '</div>';
    }
}

function _toggleHistOPME(idx) {
    var body  = document.getElementById('opme-hbody-'  + idx);
    var arrow = document.getElementById('opme-harrow-' + idx);
    var open  = body.style.display === 'block';
    body.style.display = open ? 'none' : 'block';
    if (arrow) arrow.textContent = open ? '▸' : '▾';
}

async function abrirCopOPME() {
    var btn = document.getElementById('opme-tbtn-cop');
    if (btn) btn.style.display = 'inline-flex';
    switchOPMETab('cop');
    var cont = document.getElementById('opme-cop-content');
    cont.innerHTML = '<div class="opme-spinner-wrap"><div class="opme-spin"></div><div style="font-size:13px;color:var(--text-muted)">Carregando Alerta COP...</div></div>';
    try {
        var res  = await fetch(addAuth(URL_SCRIPT + '?action=getCopOPME'));
        var data = await res.json();
        if (data.erro) throw new Error(data.erro);
        if (!data.itens || !data.itens.length) {
            cont.innerHTML = '<div class="opme-empty-state">Nenhum item OPME encontrado na planilha COP.</div>';
            return;
        }
        var cols = Object.keys(data.itens[0]);
        var sel  = '';
        if (data.abas && data.abas.length > 1) {
            sel = '<select onchange="_trocarAbaCOP(this.value)" style="background:var(--card-bg);border:1px solid var(--border);border-radius:8px;padding:6px 12px;color:var(--text-main);font-size:12px;outline:none">' +
                data.abas.map(function(a){ return '<option value="' + a + '"' + (a===data.abaAtual?' selected':'') + '>' + a + '</option>'; }).join('') +
                '</select>';
        }
        var html =
            '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:8px;">' +
            '<div style="font-size:11px;color:var(--text-muted)">' + data.total + ' item(s) OPME na COP atual</div>' + sel +
            '</div>' +
            '<div class="opme-tbl-wrap"><table class="opme-tbl">' +
            '<thead><tr>' + cols.map(function(c){ return '<th>' + c + '</th>'; }).join('') + '</tr></thead>' +
            '<tbody>' + data.itens.map(function(it){
                return '<tr>' + cols.map(function(c){
                    var v = it[c]||'';
                    return '<td style="max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="' + v + '">' + v + '</td>';
                }).join('') + '</tr>';
            }).join('') + '</tbody></table></div>';
        cont.innerHTML = html;
    } catch(err) {
        cont.innerHTML = '<div class="opme-empty-state" style="color:var(--danger)">Erro ao carregar COP: ' + err.message + '</div>';
    }
}

async function _trocarAbaCOP(aba) {
    var res  = await fetch(addAuth(URL_SCRIPT + '?action=getCopOPME&aba=' + encodeURIComponent(aba)));
    var data = await res.json();
    // re-renderiza
    abrirCopOPME();
}
// ── AUTOCOMPLETE FORNECEDOR ──────────────────────────────
var _acCache = [];
var _acIdx   = -1;

function acPopularCache() {
    const set = new Set();
    if (_histNotas && _histNotas.length) {
        _histNotas.forEach(n => { if (n.fornecedor && n.fornecedor.trim()) set.add(n.fornecedor.trim().toUpperCase()); });
    }
    try {
        const salvo = JSON.parse(sessionStorage.getItem('core_fornecedores') || '[]');
        salvo.forEach(f => set.add(f));
    } catch(e) {}
    _acCache = Array.from(set).sort();
    try { sessionStorage.setItem('core_fornecedores', JSON.stringify(_acCache)); } catch(e) {}
}

function acFornecedor(q) {
    const list = document.getElementById('ac-list');
    if (!list) return;
    if (!q || q.length < 2) { list.style.display = 'none'; _acIdx = -1; return; }
    if (!_acCache.length) acPopularCache();
    const upper    = q.toUpperCase();
    const starts   = _acCache.filter(f => f.startsWith(upper));
    const contains = _acCache.filter(f => !f.startsWith(upper) && f.includes(upper));
    const matches  = [...starts, ...contains].slice(0, 8);
    if (!matches.length) { list.style.display = 'none'; _acIdx = -1; return; }
    list.innerHTML = matches.map((f, i) => {
        const idx2 = f.toUpperCase().indexOf(upper);
        const hl = f.slice(0,idx2) + '<strong>' + f.slice(idx2, idx2+q.length) + '</strong>' + f.slice(idx2+q.length);
        return `<li data-val="${f}" onmousedown="acSelecionar('${f.replace(/'/g,"\\'")}')"> ${hl} </li>`;
    }).join('');
    _acIdx = -1;
    list.style.display = 'block';
}

function acSelecionar(valor) {
    const input = document.getElementById('f_fornecedor');
    const list  = document.getElementById('ac-list');
    if (input) input.value = valor;
    if (list)  list.style.display = 'none';
    _acIdx = -1;
    const next = document.getElementById('f_razao');
    if (next) next.focus();
}

function acNavegar(e) {
    const list  = document.getElementById('ac-list');
    if (!list || list.style.display === 'none') return;
    const items = list.querySelectorAll('li');
    if (!items.length) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); _acIdx = Math.min(_acIdx+1, items.length-1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); _acIdx = Math.max(_acIdx-1, 0); }
    else if (e.key === 'Enter' && _acIdx >= 0) { e.preventDefault(); acSelecionar(items[_acIdx].dataset.val); return; }
    else if (e.key === 'Escape') { list.style.display = 'none'; _acIdx = -1; return; }
    else return;
    items.forEach((li, i) => li.classList.toggle('ac-active', i === _acIdx));
    items[_acIdx].scrollIntoView({ block: 'nearest' });
}

document.addEventListener('click', e => {
    const input = document.getElementById('f_fornecedor');
    const list  = document.getElementById('ac-list');
    if (list && input && !input.contains(e.target) && !list.contains(e.target))
        list.style.display = 'none';
});