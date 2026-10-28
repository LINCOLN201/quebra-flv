/**
 * Quebra FLV -> CONTROLE DE APROVEITAMENTO DIARIO - FLV
 *
 * Recebe do app Quebra FLV os itens marcados com R (aproveitamento do refeitório)
 * e grava na aba do dia da semana (SEGUNDA, TERÇA, ... DOMINGO):
 *   B = Cod. Produto, C = Descrição, D = Quantidade, E = Preço de custo, F = Valor total.
 * O custo vem da aba CUSTO pelo código do produto (e, se não achar, pelo nome).
 *
 * Instalação: na planilha, Extensões > Apps Script, cole este arquivo e implante como
 * "App da Web" (Executar como: Eu / Quem pode acessar: Qualquer pessoa).
 */

var ABAS_POR_DIA = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SABADO"];
var NOMES_ALTERNATIVOS = { "TERÇA": ["TERCA"], "SABADO": ["SÁBADO"] };
var PRIMEIRA_LINHA = 9;
var ULTIMA_LINHA = 42;   // a linha 43 é a do VALOR TOTAL
var CELULA_DATA = "C5";
var CELULA_TOTAL = "F43";

function doGet() {
  return resposta({ ok: true, info: "Quebra FLV - aproveitamento do refeitório" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var req = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (req.acao === "teste") return resposta({ ok: true, planilha: ss.getName() });
    if (req.acao !== "lancar") return resposta({ ok: false, erro: "Ação desconhecida" });
    return resposta(lancar(ss, req));
  } catch (err) {
    return resposta({ ok: false, erro: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function lancar(ss, req) {
  var partes = String(req.data || "").split("-").map(Number);
  if (partes.length !== 3 || partes.some(isNaN)) return { ok: false, erro: "Data inválida" };
  var dia = new Date(partes[0], partes[1] - 1, partes[2]);
  var nomeAba = ABAS_POR_DIA[dia.getDay()];
  var sh = acharAba(ss, nomeAba);
  if (!sh) return { ok: false, erro: "Não achei a aba " + nomeAba + " na planilha" };

  var itens = (req.itens || []).filter(function (i) { return i && i.cod && Number(i.qtd) > 0; });
  if (!itens.length) return { ok: false, erro: "Nenhum item para lançar" };

  // A aba é da semana certa? (a data em C5 tem que ser a do dia enviado)
  var atual = sh.getRange(CELULA_DATA).getValue();
  var mesmaData = atual instanceof Date && atual.getFullYear() === dia.getFullYear() &&
                  atual.getMonth() === dia.getMonth() && atual.getDate() === dia.getDate();
  var n = ULTIMA_LINHA - PRIMEIRA_LINHA + 1;
  var colB = sh.getRange(PRIMEIRA_LINHA, 2, n, 1).getValues();
  var ocupadas = colB.filter(function (r) { return r[0] !== "" && r[0] !== null; }).length;
  if (!mesmaData && ocupadas > 0 && !req.novaSemana) {
    return { ok: false, outraData: true, aba: nomeAba,
             abaData: atual instanceof Date ? Utilities.formatDate(atual, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : String(atual || "") };
  }
  if (!mesmaData) {
    if (ocupadas > 0) limparAba(sh);
    sh.getRange(CELULA_DATA).setValue(dia);
    colB = sh.getRange(PRIMEIRA_LINHA, 2, n, 1).getValues();
  }

  var nomes = nomesPorCodigo(ss);
  var gravados = [], linhas = [], naoCouberam = [];
  var livre = 0;
  itens.forEach(function (item) {
    while (livre < n && colB[livre][0] !== "" && colB[livre][0] !== null) livre++;
    if (livre >= n) { naoCouberam.push(String(item.desc || item.cod)); return; }
    var lin = PRIMEIRA_LINHA + livre;
    var cod = String(item.cod).trim();
    var codValor = /^\d+$/.test(cod) && cod.length < 12 ? Number(cod) : cod;
    var descricao = nomes[cod] || String(item.desc || "");
    sh.getRange(lin, 2, 1, 3).setValues([[codValor, descricao, Number(item.qtd)]]);
    sh.getRange(lin, 5).setFormula(formulaCusto(lin));
    sh.getRange(lin, 6).setFormula("=D" + lin + "*E" + lin);
    colB[livre][0] = codValor;
    gravados.push(cod); linhas.push(lin);
    livre++;
  });
  SpreadsheetApp.flush();
  var total = sh.getRange(CELULA_TOTAL).getValue();
  return { ok: true, aba: nomeAba, gravados: gravados, linhas: linhas, naoCouberam: naoCouberam,
           total: typeof total === "number" ? total : null };
}

// Custo unitário da aba CUSTO: primeiro pelo código, depois pelo nome (como a planilha já fazia).
function formulaCusto(lin) {
  return "=IFERROR(VLOOKUP(B" + lin + "&\"\",CUSTO!A:E,5,0),IFERROR(VLOOKUP(B" + lin +
         ",CUSTO!A:E,5,0),IFERROR(VLOOKUP(C" + lin + ",CUSTO!B:E,4,0),0)))";
}

// Nome do produto pelo código: CUSTO primeiro (mais atual), depois BANCO.
function nomesPorCodigo(ss) {
  var mapa = {};
  ["BANCO", "CUSTO"].forEach(function (aba) {
    var sh = ss.getSheetByName(aba); if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    vals.forEach(function (r) {
      if (r[0] === "" || r[0] === null || !r[1]) return;
      var c = typeof r[0] === "number" ? String(Math.round(r[0])) : String(r[0]).trim();
      mapa[c] = String(r[1]).trim();
    });
  });
  return mapa;
}

// Começa a semana nova nessa aba: apaga código e quantidade e refaz as fórmulas.
function limparAba(sh) {
  for (var lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) {
    sh.getRange(lin, 2).clearContent();
    sh.getRange(lin, 4).clearContent();
    sh.getRange(lin, 3).setFormula("=IFERROR(VLOOKUP(B" + lin + ",BANCO!A:B,2,0),\"\")");
    sh.getRange(lin, 5).setFormula(formulaCusto(lin));
    sh.getRange(lin, 6).setFormula("=D" + lin + "*E" + lin);
  }
}

function acharAba(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (sh) return sh;
  var alt = NOMES_ALTERNATIVOS[nome] || [];
  for (var i = 0; i < alt.length; i++) { sh = ss.getSheetByName(alt[i]); if (sh) return sh; }
  return null;
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
