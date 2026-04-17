const USUARIOS = [{ login: "mateus", senha: "123", nome: "Mateus" }];
const URL_SCRIPT = "SUA_URL_AQUI"; 

let abaAtual = "Digitadas";
let listas = { "Digitadas": [], "Recebimento": [] };
let usuarioLogado = "";

function toggleTheme() {
    const body = document.body;
    const icon = document.getElementById('themeIcon');
    const text = document.getElementById('themeText');
    const isDark = body.getAttribute('data-theme') === 'dark';

    if (isDark) {
        body.removeAttribute('data-theme');
        icon.className = 'ph ph-moon';
        text.innerText = 'MODO ESCURO';
        localStorage.setItem('theme', 'light');
    } else {
        body.setAttribute('data-theme', 'dark');
        icon.className = 'ph ph-sun';
        text.innerText = 'MODO CLARO';
        localStorage.setItem('theme', 'dark');
    }
}

if (localStorage.getItem('theme') === 'dark') toggleTheme();

function realizarLogin() {
    const u = document.getElementById('userInput').value.toLowerCase();
    const s = document.getElementById('passInput').value;
    const user = USUARIOS.find(x => x.login === u && x.senha === s);

    if (user) {
        usuarioLogado = user.nome;
        document.getElementById('loginScreen').style.display = 'none';
        document.getElementById('mainHeader').style.display = 'flex';
        document.getElementById('mainFooter').style.display = 'block'; // Mostra rodapé
        document.getElementById('app').style.display = 'block';
        document.getElementById('userDisplay').innerText = usuarioLogado.toUpperCase();
        switchTab('Digitadas');
    } else {
        alert("Usuário ou senha incorretos.");
    }
}

function logout() {
    usuarioLogado = "";
    document.getElementById('app').style.display = 'none';
    document.getElementById('mainHeader').style.display = 'none';
    document.getElementById('mainFooter').style.display = 'none'; // Esconde rodapé
    document.getElementById('loginScreen').style.display = 'block';
    document.getElementById('userInput').value = "";
    document.getElementById('passInput').value = "";
    listas = { "Digitadas": [], "Recebimento": [] };
}

function switchTab(aba) {
    abaAtual = aba;
    document.getElementById('tabTitle').innerText = "Lote de Notas " + aba;
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + aba.toLowerCase()).classList.add('active');

    document.getElementById('statusDigitacaoArea').style.display = (aba === 'Digitadas') ? 'block' : 'none';
    document.getElementById('statusRecebimentoArea').style.display = (aba === 'Recebimento') ? 'block' : 'none';
    document.getElementById('btnCopiar').style.display = (aba === 'Digitadas') ? 'flex' : 'none';
    document.getElementById('thStatus').innerText = (aba === 'Digitadas') ? 'STATUS MV' : 'LOGÍSTICA';

    atualizarTabela();
}

function adicionarNota() {
    const nf = document.getElementById('f_nf').value;
    if (!nf) return alert("NF é obrigatória!");

    let statusFinal = "";
    if (abaAtual === 'Digitadas') {
        const s = document.querySelector('input[name="gSistema"]:checked');
        const p = document.querySelector('input[name="gProcesso"]:checked');
        if (!s || !p) return alert("Selecione o Sistema e o Processo!");
        statusFinal = s.value + " | " + p.value;
    } else {
        const l = document.querySelector('input[name="gLogistica"]:checked');
        if (!l) return alert("Selecione a situação logística!");
        statusFinal = l.value;
    }

    const nota = {
        destino: abaAtual,
        responsavel: usuarioLogado,
        data: document.getElementById('f_data').value,
        nf: nf,
        fornecedor: document.getElementById('f_fornecedor').value,
        razaoSocial: document.getElementById('f_razao').value,
        vencimento: document.getElementById('f_vencimento').value,
        valor: document.getElementById('f_valor').value,
        setor: document.getElementById('f_setor').value || "GERAL",
        possuiLote: document.getElementById('f_lote').value,
        statusDigitacao: (abaAtual === 'Digitadas') ? statusFinal : "",
        situacaoMaterial: (abaAtual === 'Recebimento') ? statusFinal : ""
    };

    listas[abaAtual].push(nota);
    atualizarTabela();
    document.getElementById('f_nf').value = "";
    document.getElementById('f_valor').value = "";
    document.getElementById('f_nf').focus();
}

function atualizarTabela() {
    const tbody = document.querySelector("#tabelaDados tbody");
    tbody.innerHTML = "";
    listas[abaAtual].forEach((n, i) => {
        const st = (abaAtual === 'Digitadas') ? n.statusDigitacao : n.situacaoMaterial;
        tbody.innerHTML += `<tr>
            <td><b>${n.nf}</b></td>
            <td>${n.fornecedor}</td>
            <td>${n.razaoSocial === 'Fundação Zerbini' ? 'FZ' : 'HC'}</td>
            <td>${n.vencimento.split('-').reverse().join('/')}</td>
            <td>${n.setor}</td>
            <td>${n.possuiLote}</td>
            <td>${st}</td>
            <td><button onclick="removerNota(${i})" style="border:none; background:none; cursor:pointer; color: var(--text-muted)"><i class="ph ph-trash" style="font-size:20px"></i></button></td>
        </tr>`;
    });
    document.getElementById('areaAcoes').style.display = listas[abaAtual].length > 0 ? 'grid' : 'none';
}

function removerNota(i) {
    listas[abaAtual].splice(i, 1);
    atualizarTabela();
}

function copiarProtocolo() {
    let texto = `*PROTOCOLO DE DIGITAÇÃO - RESPONSÁVEL: ${usuarioLogado.toUpperCase()}*\n\n`;
    texto += `NF\t| FORNECEDOR\t| RAZÃO\t| VENCIMENTO\t| SETOR\t| LOTE\n`;
    listas['Digitadas'].forEach(n => {
        texto += `${n.nf}\t| ${n.fornecedor}\t| ${n.razaoSocial === 'Fundação Zerbini' ? 'FZ' : 'HC'}\t| ${n.vencimento.split('-').reverse().join('/')}\t| ${n.setor}\t| ${n.possuiLote}\n`;
    });
    navigator.clipboard.writeText(texto).then(() => alert("✅ Protocolo copiado!"));
}

function enviarTudo() {
    const btn = document.getElementById('btnEnviar');
    btn.disabled = true;
    btn.innerText = "SINCRONIZANDO...";

    fetch(URL_SCRIPT, { method: 'POST', mode: 'no-cors', body: JSON.stringify(listas[abaAtual]) })
        .then(() => {
            alert("🚀 Sucesso! Lote sincronizado.");
            listas[abaAtual] = [];
            atualizarTabela();
        })
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = '<i class="ph-bold ph-cloud-arrow-up"></i> ENVIAR PARA PLANILHA';
        });
}