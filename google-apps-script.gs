/**
 * Backend do pedido "Coleção Comemorativa 1º Esqd C Pqdt" (camisas + copos) — Google Apps Script.
 *
 * O que faz a cada pedido recebido do index.html:
 *  - Aba "Pedidos": 1 linha por pedido (militar, qtd de camisas, qtd de copos,
 *    total, link do comprovante, status do pagamento e uma caixa "Conferido").
 *  - Aba "Camisas": 1 linha por item de camisa (modelo, cor, tamanho, quantidade).
 *  - Aba "Resumo": tabela Modelo × Cor × Tamanho das camisas + total de copos,
 *    pronta para mandar ao fornecedor.
 *  - Salva o comprovante numa pasta do seu Google Drive
 *    ("Comprovantes — Camisa 1º Esqd C Pqdt"), visível só para você.
 *  - Ignora reenvios do mesmo pedido (mesmo número), evitando duplicatas.
 *
 * COMO IMPLANTAR (uma única vez):
 * 1. Crie uma planilha nova no Google Sheets (sheets.new).
 * 2. Menu Extensões → Apps Script.
 * 3. Apague o conteúdo do editor e cole todo o conteúdo deste arquivo.
 * 4. Clique em Implantar → Nova implantação → "App da Web",
 *    Executar como "Eu", Quem tem acesso "Qualquer pessoa".
 * 5. Autorize as permissões (Planilhas e Drive) e copie a URL do App da Web.
 * 6. Cole essa URL em CONFIG.WEBHOOK_URL no topo do script do index.html.
 *
 * PARA ATUALIZAR SEM TROCAR A URL: Implantar → Gerenciar implantações →
 * lápis (editar) → Versão: "Nova versão" → Implantar.
 *
 * Se apagar/editar linhas na mão (ex: pedido cancelado), use o menu
 * "Camisas → Atualizar resumo" na planilha para recalcular a aba Resumo.
 */

var CONFIG = {
  PRECO_CAMISA: 50,
  PRECO_COPO: 45,
  PASTA_COMPROVANTES: "Comprovantes — Camisa 1º Esqd C Pqdt",
  MODELOS: ["Masculina", "Feminina", "Infantil"],
  CORES: ["Bordô", "Bege"],
  TAMANHOS: ["PP", "P", "M", "G", "GG"]
};

var CAB_PEDIDOS = ["Recebido em", "Pedido", "Nome do militar", "Nome de guerra", "Qtd camisas",
                   "Qtd copos", "Total (R$)", "Itens", "Comprovante", "Pagamento", "Conferido"];
var CAB_CAMISAS = ["Recebido em", "Pedido", "Nome do militar", "Modelo", "Cor", "Tamanho", "Qtd"];
var COL = { COPOS: 6, TOTAL: 7, CONFERIDO: 11 };

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var data = JSON.parse(e.postData.contents);
    var itens = (data.itens || []).map(function (i) {
      return { modelo: i.modelo, cor: i.cor, tamanho: i.tamanho,
               quantidade: Math.max(1, parseInt(i.quantidade, 10) || 1) };
    });
    var copos = Math.max(0, parseInt(data.copos, 10) || 0);
    if (!data.id || !data.nome || (!itens.length && !copos)) {
      return json_({ status: "error", message: "Pedido incompleto." });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaPedidos = obterAbaPedidos_(ss);
    var abaCamisas = obterAba_(ss, "Camisas", CAB_CAMISAS);

    if (pedidoExiste_(abaPedidos, data.id)) {
      return json_({ status: "ok", id: data.id, duplicado: true });
    }

    var link = "";
    if (data.comprovante && data.comprovante.base64) {
      link = salvarComprovante_(data);
    }

    var agora = new Date();
    var qtdCamisas = itens.reduce(function (s, i) { return s + i.quantidade; }, 0);
    // calculado aqui, não confia no navegador
    var total = qtdCamisas * CONFIG.PRECO_CAMISA + copos * CONFIG.PRECO_COPO;
    var linhasTexto = itens.map(function (i) {
      return i.quantidade + "× Camisa " + i.modelo + " · " + i.cor + " · " + i.tamanho;
    });
    if (copos) linhasTexto.push(copos + "× Copo Térmico 4 em 1");

    abaPedidos.appendRow([
      agora, data.id, data.nome, data.nome_de_guerra || "", qtdCamisas, copos,
      total, linhasTexto.join("\n"), link, link ? "Comprovante enviado" : "Pendente", false
    ]);
    var linha = abaPedidos.getLastRow();
    abaPedidos.getRange(linha, COL.TOTAL).setNumberFormat('"R$" #,##0.00');
    abaPedidos.getRange(linha, COL.CONFERIDO).insertCheckboxes();

    if (itens.length) {
      var linhasCamisas = itens.map(function (i) {
        return [agora, data.id, data.nome, i.modelo, i.cor, i.tamanho, i.quantidade];
      });
      abaCamisas.getRange(abaCamisas.getLastRow() + 1, 1, linhasCamisas.length, CAB_CAMISAS.length)
        .setValues(linhasCamisas);
    }

    atualizarResumo();

    return json_({ status: "ok", id: data.id });
  } catch (err) {
    return json_({ status: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Só para testar a implantação abrindo a URL do App da Web direto no navegador.
function doGet() {
  return ContentService
    .createTextOutput("Backend da Coleção Comemorativa 1º Esqd C Pqdt (camisas + copos) está no ar.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Camisas")
    .addItem("Atualizar resumo", "atualizarResumo")
    .addToUi();
}

/** Recria a aba "Resumo" a partir das abas "Camisas" e "Pedidos". */
function atualizarResumo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCamisas = obterAba_(ss, "Camisas", CAB_CAMISAS);
  var abaPedidos = obterAbaPedidos_(ss);
  var resumo = ss.getSheetByName("Resumo") || ss.insertSheet("Resumo");

  var n = abaCamisas.getLastRow() - 1;
  var dados = n > 0 ? abaCamisas.getRange(2, 4, n, 4).getValues() : []; // Modelo, Cor, Tamanho, Qtd
  var cont = {};
  dados.forEach(function (r) {
    var k = r[0] + "|" + r[1] + "|" + r[2];
    cont[k] = (cont[k] || 0) + (Number(r[3]) || 0);
  });

  var np = abaPedidos.getLastRow() - 1;
  var totalCopos = 0, totalValor = 0;
  if (np > 0) {
    abaPedidos.getRange(2, COL.COPOS, np, 2).getValues().forEach(function (r) {
      totalCopos += Number(r[0]) || 0;
      totalValor += Number(r[1]) || 0;
    });
  }

  var tabela = [["Modelo", "Cor"].concat(CONFIG.TAMANHOS).concat(["Total"])];
  var totaisColuna = CONFIG.TAMANHOS.map(function () { return 0; });
  var totalCamisas = 0;
  CONFIG.MODELOS.forEach(function (m) {
    CONFIG.CORES.forEach(function (c) {
      var linha = [m, c], soma = 0;
      CONFIG.TAMANHOS.forEach(function (t, i) {
        var q = cont[m + "|" + c + "|" + t] || 0;
        linha.push(q); soma += q; totaisColuna[i] += q;
      });
      linha.push(soma); totalCamisas += soma;
      tabela.push(linha);
    });
  });
  tabela.push(["TOTAL", ""].concat(totaisColuna).concat([totalCamisas]));

  resumo.clear();
  resumo.getRange(1, 1).setValue("Resumo para os fornecedores").setFontWeight("bold").setFontSize(14);
  resumo.getRange(2, 1).setValue("Atualizado em " +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm"));

  resumo.getRange(4, 1).setValue("CAMISAS").setFontWeight("bold");
  var ini = 5;
  resumo.getRange(ini, 1, tabela.length, tabela[0].length).setValues(tabela).setHorizontalAlignment("center");
  resumo.getRange(ini, 1, 1, tabela[0].length).setFontWeight("bold").setBackground("#8FD3F4");
  resumo.getRange(ini + tabela.length - 1, 1, 1, tabela[0].length).setFontWeight("bold").setBackground("#FADC1C");

  var linhaInfo = ini + tabela.length + 1;
  resumo.getRange(linhaInfo, 1, 4, 2).setValues([
    ["Camisas", totalCamisas],
    ["Copos Térmicos 4 em 1", totalCopos],
    ["", ""],
    ["Valor total dos pedidos (R$)", totalValor]
  ]);
  resumo.getRange(linhaInfo + 1, 1, 1, 2).setBackground("#FADC1C");
  resumo.getRange(linhaInfo + 3, 2).setNumberFormat('"R$" #,##0.00');
  resumo.getRange(linhaInfo, 1, 4, 1).setFontWeight("bold");
}

/* ---------------- auxiliares ---------------- */

/** Aba "Pedidos" com o cabeçalho atual; migra a versão antiga (só camisas) inserindo a coluna de copos. */
function obterAbaPedidos_(ss) {
  var aba = obterAba_(ss, "Pedidos", CAB_PEDIDOS);
  var cab = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 1)).getValues()[0];
  if (cab[COL.COPOS - 1] !== "Qtd copos") {
    aba.insertColumnAfter(COL.COPOS - 1); // pedidos antigos ficam com copos em branco
    aba.getRange(1, 1, 1, CAB_PEDIDOS.length).setValues([CAB_PEDIDOS])
      .setFontWeight("bold").setBackground("#8FD3F4");
  }
  return aba;
}

function obterAba_(ss, nome, cabecalho) {
  var aba = ss.getSheetByName(nome);
  if (!aba) {
    // Reaproveita a "Página1"/"Sheet1" vazia criada junto com a planilha.
    var abas = ss.getSheets();
    if (abas.length === 1 && abas[0].getLastRow() === 0 && !/^(Pedidos|Camisas|Resumo)$/.test(abas[0].getName())) {
      aba = abas[0].setName(nome);
    } else {
      aba = ss.insertSheet(nome);
    }
  }
  if (aba.getLastRow() === 0) {
    aba.appendRow(cabecalho);
    aba.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#8FD3F4");
    aba.setFrozenRows(1);
  }
  return aba;
}

function pedidoExiste_(aba, id) {
  if (aba.getLastRow() < 2) return false;
  return !!aba.getRange(2, 2, aba.getLastRow() - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
}

function pastaComprovantes_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("PASTA_ID");
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* pasta apagada: cria outra */ }
  }
  var pasta = DriveApp.createFolder(CONFIG.PASTA_COMPROVANTES);
  props.setProperty("PASTA_ID", pasta.getId());
  return pasta;
}

function salvarComprovante_(data) {
  var c = data.comprovante;
  var ext = (String(c.nome || "").match(/\.[a-z0-9]+$/i) || [c.mime === "application/pdf" ? ".pdf" : ".jpg"])[0];
  var nomeArquivo = data.id + " — " + data.nome + ext;
  var blob = Utilities.newBlob(Utilities.base64Decode(c.base64), c.mime || "application/octet-stream", nomeArquivo);
  return pastaComprovantes_().createFile(blob).getUrl();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
