// ================================================================
//  NFControl — Código.gs
//  Cole este código no Apps Script da sua planilha Google Sheets
//  (Extensões > Apps Script)
// ================================================================

// ──────────────────────────────────────────────────────────────────
// CONFIGURAÇÃO
// ──────────────────────────────────────────────────────────────────

// Nome da aba onde os dados serão salvos
var NOME_DA_ABA = "Lançamentos";

// Cabeçalhos das colunas (devem existir na primeira linha da planilha)
var CABECALHOS = [
  "Responsável",
  "Data",
  "NF",
  "Fornecedor",
  "Razão Social",
  "Vencimento",
  "Valor",
  "Setor",
  "Timestamp"   // coluna extra para controle interno
];

// ──────────────────────────────────────────────────────────────────
// PONTO DE ENTRADA — recebe requisições POST do frontend
// ──────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    // 1. Parse do payload JSON enviado pelo frontend
    var body = JSON.parse(e.postData.contents);

    // 2. Abre a planilha ativa e localiza (ou cria) a aba correta
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(NOME_DA_ABA);

    if (!aba) {
      // Cria a aba se não existir e insere os cabeçalhos
      aba = ss.insertSheet(NOME_DA_ABA);
      aba.appendRow(CABECALHOS);
      formatarCabecalhos(aba);
    }

    // 3. Garante que os cabeçalhos existem (segurança extra)
    if (aba.getLastRow() === 0) {
      aba.appendRow(CABECALHOS);
      formatarCabecalhos(aba);
    }

    // 4. Monta a linha de dados na mesma ordem dos cabeçalhos
    var timestamp = Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      "dd/MM/yyyy HH:mm:ss"
    );

    var linha = [
      body.responsavel  || "",
      body.data         || "",
      body.nf           || "",
      body.fornecedor   || "",
      body.razao_social || "",
      body.vencimento   || "",
      body.valor        || "",
      body.setor        || "",
      timestamp
    ];

    // 5. Insere na próxima linha disponível
    aba.appendRow(linha);

    // 6. Formata a coluna de Valor como moeda (opcional)
    var ultimaLinha = aba.getLastRow();
    var colunaValor = 7; // coluna G (1-indexed)
    aba.getRange(ultimaLinha, colunaValor)
       .setNumberFormat('"R$" #,##0.00');

    // 7. Retorna resposta de sucesso com headers CORS
    return criarResposta({ status: "ok", linha: ultimaLinha });

  } catch (erro) {
    Logger.log("Erro em doPost: " + erro.toString());
    return criarResposta({ status: "erro", mensagem: erro.toString() }, true);
  }
}

// ──────────────────────────────────────────────────────────────────
// RESPOSTA JSON com headers CORS
// ──────────────────────────────────────────────────────────────────
function criarResposta(obj, isErro) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  return output;
}

// ──────────────────────────────────────────────────────────────────
// FORMATAÇÃO DOS CABEÇALHOS
// ──────────────────────────────────────────────────────────────────
function formatarCabecalhos(aba) {
  var range = aba.getRange(1, 1, 1, CABECALHOS.length);
  range.setBackground("#1a1a2e")
       .setFontColor("#c8f050")
       .setFontWeight("bold")
       .setFontSize(10);
  aba.setFrozenRows(1); // Congela linha de cabeçalho
}

// ──────────────────────────────────────────────────────────────────
// GET (opcional) — útil para testar se o script está publicado
// ──────────────────────────────────────────────────────────────────
function doGet(e) {
  return criarResposta({ status: "online", mensagem: "NFControl Apps Script ativo." });
}
