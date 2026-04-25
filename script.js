// --- CONFIGURAÇÃO ---
const URL_SCRIPT = "https://script.google.com/macros/s/AKfycbzrbc6xqFhpqRw2U9_1T4_rhscRJWTWlQPsCFH_5JM5Kedlq-DJj5IPpTkG3m9zcaHB2Q/exec";

let usuarioAtual = null;
let loginAtual   = null;
let sessaoAtual  = null; // { token, expira }
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
    b.setAttribute('data-theme', isDark ? 'light' : 'dark');
    const newIcon = isDark ? 'ph ph-moon' : 'ph ph-sun';
    const headerIcon = document.getElementById('themeIcon');
    const loginIcon  = document.getElementById('themeIconLogin');
    if (headerIcon) headerIcon.className = newIcon;
    if (loginIcon)  loginIcon.className  = newIcon;
}

function logout() { location.reload(); }

// --- BOTÃO DE REFRESH DINÂMICO ---
async function refreshData() {
    const icon = document.getElementById('refreshIcon');
    icon.classList.add('rotating');

    try {
        // Sempre recarrega blacklist e estatísticas do dashboard em background
        await carregarBlacklistCache();
        await carregarEstatisticas();

        // Recarrega dados específicos da aba atual
        if (abaAtual === 'Dashboard') {
            // já feito acima

        } else if (abaAtual === 'Projecao' || abaAtual.startsWith('Projecao_')) {
            dadosProjecao = []; // força novo fetch
            await carregarProjecao(abaAtual.startsWith('Projecao_') ? abaAtual.replace('Projecao_', '') : null);

        } else if (abaAtual === 'Adiantamento') {
            adiantamentosCarregados = []; // força novo fetch
            await carregarAdiantamentos();

        } else if (abaAtual === 'Admin') {
            await carregarUsuarios();

        }
        // Abas de notas: a tabela local já é a fonte da verdade, não precisa recarregar

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
            // Armazena token de sessão em memória
            sessaoAtual = {
                token:  data.token,
                expira: Date.now() + (8 * 60 * 60 * 1000) // 8 horas
            };
            // Timer de expiração automática
            setTimeout(() => {
                mostrarToast('⏱️ Sua sessão expirou. Faça login novamente.', 'warning', 6000);
                setTimeout(() => logout(), 3000);
            }, 8 * 60 * 60 * 1000);

            if (data.primeiroAcesso) {
                document.getElementById('loginScreen').style.display = 'none';
                document.getElementById('primeiroAcessoScreen').style.display = 'flex';
                document.getElementById('nomeBoasVindas').innerText = data.nome.split(' ')[0];
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
            const nomeUsuario = comp.nomeUsuario; // ex: CRISLENE, ERNESTO
            const canvasId   = `dashPizza_${nomeUsuario}`;
            const legendaId  = `dashLegenda_${nomeUsuario}`;
            const totalId    = `dashTotal_${nomeUsuario}`;

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
                    <div style="display:flex;align-items:center;gap:24px;">
                        <canvas id="${canvasId}" width="130" height="130" style="flex-shrink:0;"></canvas>
                        <div id="${legendaId}" style="display:flex;flex-direction:column;gap:7px;font-size:12px;flex:1;"></div>
                    </div>`;
                grid.appendChild(card);
            }

            // Carrega dados da projeção
            try {
                const res  = await fetch(addAuth(`${URL_SCRIPT}?action=projecao&usuario=${encodeURIComponent(nomeUsuario)}`));
                const data = await res.json();
                if (!data.erro && data.itens) {
                    desenharMiniPizza(canvasId, legendaId, totalId, data.itens);
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

    const podeVerProjecao = temPermissao('comprador') || temPermissao('diretor') || temPermissao('administrador');
    const isDiretor = temPermissao('diretor') && !temPermissao('comprador');

    // Comprador normal: mostra aba única com o nome dele
    if (podeVerProjecao && !temPermissao('diretor')) {
        const primeiroNome = usuarioAtual.nome.split(' ')[0];
        document.getElementById('btnProjecaoLabel').innerText = `PROJ. ${primeiroNome.toUpperCase()}`;
        document.getElementById('btn-projecao').style.display = 'inline-flex';
    } else if (temPermissao('comprador') && temPermissao('diretor')) {
        // Tem comprador + diretor: mostra a própria aba + abas dos outros
        const primeiroNome = usuarioAtual.nome.split(' ')[0];
        document.getElementById('btnProjecaoLabel').innerText = `PROJ. ${primeiroNome.toUpperCase()}`;
        document.getElementById('btn-projecao').style.display = 'inline-flex';
        criarTabsDiretor();
    } else if (temPermissao('diretor')) {
        // Diretor puro: só vê abas dos compradores
        document.getElementById('btn-projecao').style.display = 'none';
        criarTabsDiretor();
    } else {
        document.getElementById('btn-projecao').style.display = 'none';
    }

    document.getElementById('carrinhoHeaderBtn').style.display = temPermissao('comprador') || temPermissao('administrador') ? 'flex' : 'none';
    document.getElementById('btn-admin').style.display = temPermissao('administrador') ? 'inline-flex' : 'none';

    switchTab('Dashboard');
    carregarBlacklistCache();
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

        // Carrega dados de projeção em background
        await carregarProjecaoBackground();
        // Carrega pizzas das projeções nos cards da dashboard
        carregarPizzasDashboard();
    } catch (e) {
        document.getElementById('dash-loading').innerHTML =
            "<p style='color:var(--danger)'>Erro ao carregar dados do Google Sheets.</p>";
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
            // Guarda os prefixos de cada comprador para usar ao trocar de aba
            if (u.prefixos) {
                prefixosConfig[u.nomeUsuario] = u.prefixos.split('|').map(p => p.trim()).filter(Boolean);
            }

            // Não duplica se o próprio diretor também for comprador
            if (u.nomeUsuario === usuarioAtual.nome.split(' ')[0].toUpperCase()) return;

            const abaId = `Projecao_${u.nomeUsuario}`;
            const btn   = document.createElement('button');
            btn.className = 'nav-btn nav-btn-proj-dir';
            btn.id        = `btn-proj-${u.nomeUsuario.toLowerCase()}`;
            btn.innerHTML = `<i class="ph ph-package"></i> PROJ. ${u.nomeUsuario}`;
            btn.onclick   = () => switchTab(abaId);
            container.appendChild(btn);
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
let filtroProjecaoAtual = 'todos';
let ordenacaoCobertura = 'asc';
let paginaAtualProjecao = 1;
const ITENS_POR_PAGINA = 100;
let itensFiltradosProjecao = [];

// Cache de adiantamentos (compartilhado entre dashboard e aba Adiantamento)
let adiantamentosCarregados = [];

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

    // Se dados já foram carregados em background para este usuário, mostra logo
    if (!nomeUsuarioForcar && dadosProjecao.length > 0) {
        filtroProjecaoAtual = 'todos';
        renderizarChipsPrefixos(prefixosAtivos);
        atualizarTabelaProjecao(aplicarFiltroPrefixo(dadosProjecao));
        desenharPizzaProjecao(dadosProjecao);
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
        renderizarChipsPrefixos(prefixosAtivos);
        atualizarTabelaProjecao(aplicarFiltroPrefixo(dadosProjecao));
        desenharPizzaProjecao(dadosProjecao);

        document.getElementById('proj-loading').style.display = 'none';
        document.getElementById('proj-content').style.display = 'block';
    } catch (e) {
        document.getElementById('proj-loading').innerHTML =
            "<p style='color:var(--danger)'>Erro ao carregar dados de projeção.</p>";
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
    let filtrado = dadosProjecao;

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
            if (item.cobertura > 25) {
                statusBadges += '<span class="badge-status" style="background:rgba(99,102,241,0.15);color:#6366f1;font-weight:800;"><i class="ph ph-clock"></i> COMPRA EM BREVE</span>';
            } else {
                statusBadges += '<span class="badge-status zerado"><i class="ph ph-warning-diamond"></i> CRÍTICO</span>';
            }
            temAlgumStatus = true;
        }

        if (!temAlgumStatus) {
            statusBadges = '<span class="badge-status" style="background: rgba(234, 88, 12, 0.15); color: #ea580c;"><i class="ph ph-shopping-cart"></i> COMPRAR</span>';
        }

        // Texto limpo do status para o carrinho/CSV
        let statusTexto = 'COMPRAR';
        if (item.zeradoSemCobertura) {
            statusTexto = item.cobertura > 25 ? 'COMPRA EM BREVE' : 'CRÍTICO';
        } else if (item.temRP)       statusTexto = 'RP';
        else if (item.saldoCD > 0)   statusTexto = 'CD';
        else if (item.temEmpenho)    statusTexto = 'EMPENHO';

        tbody.innerHTML += `
            <tr>
                <td><b>${item.codigo}</b></td>
                <td>${item.descricao}</td>
                <td>${item.cobertura || 0}</td>
                <td>${item.saldoCD || 0}</td>
                <td>${item.temRP ? '✓' : '✗'}</td>
                <td>${item.temEmpenho ? '✓' : '✗'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoDiario > 0 ? parseFloat(item.consumoDiario.toFixed(3)) : '—'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoMensal > 0 ? parseFloat(item.consumoMensal.toFixed(3)) : '—'}</td>
                <td>${statusBadges}</td>
                <td style="text-align:center;">
                    <button onclick="abrirSaldoSub('${item.codigo}', ${item.consumoDiario || 0})"
                        style="background:transparent;border:1.5px solid var(--border);border-radius:8px;padding:5px 7px;cursor:pointer;color:var(--text-muted);" title="Ver saldo nos subestoques">
                        <i class="ph ph-warehouse"></i>
                    </button>
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
    // Remove active de todos os botões (fixos e dinâmicos)
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    // Ativa o botão correto
    if (aba.startsWith('Projecao_')) {
        const nomeUsuario = aba.replace('Projecao_', '').toLowerCase();
        const btnDir = document.getElementById(`btn-proj-${nomeUsuario}`);
        if (btnDir) btnDir.classList.add('active');
    } else {
        const btnEl = document.getElementById('btn-' + aba.toLowerCase());
        if (btnEl) btnEl.classList.add('active');
    }

    const monitorSec = document.getElementById('monitorAdiantamentosSection');
    if (monitorSec) monitorSec.style.display = 'none';

    if (aba === 'Dashboard') {
        document.getElementById('view-dashboard').style.display = 'block';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'none';
        document.getElementById('view-admin').style.display = 'none';
        const isGestorOuAdmin = temPermissao('gestor') || temPermissao('administrador');
        document.getElementById('dash-gestor').style.display = isGestorOuAdmin ? 'block' : 'none';
        carregarEstatisticas();
    } else if (aba === 'Projecao' || aba.startsWith('Projecao_')) {
        document.getElementById('view-dashboard').style.display = 'none';
        document.getElementById('view-forms').style.display = 'none';
        document.getElementById('view-projecao').style.display = 'block';
        document.getElementById('view-admin').style.display = 'none';

        if (aba.startsWith('Projecao_')) {
            const nomeUsuario = aba.replace('Projecao_', '');
            document.getElementById('projTitulo').innerText = `Projeção de Compras — ${nomeUsuario.charAt(0) + nomeUsuario.slice(1).toLowerCase()}`;
            prefixosAtivos = new Set(prefixosConfig[nomeUsuario] || []);
            dadosProjecao = [];
            carregarProjecao(nomeUsuario);
        } else {
            const nomeUsuario = usuarioAtual.nome.split(' ')[0];
            document.getElementById('projTitulo').innerText = `Projeção de Compras — ${nomeUsuario}`;
            prefixosAtivos = new Set(usuarioAtual.prefixos || []);
            carregarProjecao();
        }
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
        if (aba === 'Adiantamento') {
            document.getElementById('historicoNotasSection').style.display = 'none';
            carregarAdiantamentos();
        } else if (aba === 'Digitadas' || aba === 'Recebimento') {
            setTimeout(() => carregarHistoricoNotas(), 200);
        } else {
            document.getElementById('historicoNotasSection').style.display = 'none';
        }
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
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoDiario > 0 ? parseFloat(item.consumoDiario.toFixed(3)) : '—'}</td>
                <td style="font-size:12px;color:var(--text-muted);">${item.consumoMensal > 0 ? parseFloat(item.consumoMensal.toFixed(3)) : '—'}</td>
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

function aplicarFiltroAtual() {
    let itens = aplicarFiltroPrefixo(dadosProjecao);
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
    document.getElementById('tabBtnUsuarios').classList.toggle('ativo',  tab === 'usuarios');
    document.getElementById('tabBtnBlacklist').classList.toggle('ativo', tab === 'blacklist');
    document.getElementById('tabBtnLog').classList.toggle('ativo',       tab === 'log');
    if (tab === 'blacklist') { carregarBlacklist(); preencherCompradoresAddProj(); }
    if (tab === 'log')       carregarLog();
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