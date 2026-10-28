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

// Com o acesso restrito à empresa, o app manda os dados pela URL (parâmetro d, JSON em base64url),
// carregando o script com o login do navegador. Respostas:
//   callback=nome -> JavaScript nome({...})  (o app lê o resultado)
//   sem callback  -> página com o resultado e botão para voltar ao app
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out = { ok: true, info: "Quebra FLV - aproveitamento do refeitório" };
  if (p.d) {
    try {
      out = processar(JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.d)).getDataAsString()));
    } catch (err) { out = { ok: false, erro: String(err && err.message || err) }; }
  }
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + "(" + JSON.stringify(out) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  if (p.d) return paginaResultado(out, p.volta);
  return resposta(out);
}

function doPost(e) {
  try { return resposta(processar(JSON.parse(e.postData.contents))); }
  catch (err) { return resposta({ ok: false, erro: String(err && err.message || err) }); }
}

function processar(req) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (req.acao === "teste") return { ok: true, planilha: ss.getName() };
    if (req.acao === "iniciarSemana") return iniciarSemana(ss, req);
    if (req.acao !== "lancar") return { ok: false, erro: "Ação desconhecida" };
    return lancar(ss, req);
  } catch (err) {
    return { ok: false, erro: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function paginaResultado(out, volta) {
  var titulo, texto;
  if (out.ok && out.aba) {
    titulo = "Enviado ✓";
    texto = (out.gravados || []).length + " itens na aba " + out.aba +
      (typeof out.total === "number" ? " · total do dia R$ " + out.total.toFixed(2).replace(".", ",") : "");
    if (out.naoCouberam && out.naoCouberam.length) texto += " · não couberam: " + out.naoCouberam.join(", ");
  } else if (out.ok && out.planilha) { titulo = "Ligado ✓"; texto = "Planilha: " + out.planilha; }
  else if (out.ok) { titulo = "Pronto ✓"; texto = out.segunda ? "Semana iniciada em " + out.segunda : ""; }
  else if (out.outraData) { titulo = "Aba de outra semana"; texto = "A aba " + out.aba + " está com a data " + out.abaData + ". Volte ao app e confirme para apagar e lançar."; }
  else { titulo = "Erro"; texto = out.erro || "desconhecido"; }
  var link = /^https:\/\/[\w.-]+\//.test(volta || "") ? volta : "https://lincoln201.github.io/quebra-flv/";
  var html = '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:Arial,sans-serif;max-width:420px;margin:40px auto;padding:20px;text-align:center">' +
    '<h2 style="color:' + (out.ok ? '#23603a' : '#b42318') + '">' + esc_(titulo) + '</h2><p>' + esc_(texto) + '</p>' +
    '<p><a href="' + esc_(link) + '" target="_top" style="display:inline-block;background:#23603a;color:#fff;padding:12px 18px;border-radius:12px;text-decoration:none;font-weight:bold">Voltar ao app</a></p></div>';
  return HtmlService.createHtmlOutput(html).setTitle("Quebra FLV").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function esc_(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

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

// Começa a semana: limpa as 7 abas e põe as datas de segunda a domingo.
// req.segunda = "aaaa-mm-dd" (qualquer dia da semana serve; volta até a segunda).
function iniciarSemana(ss, req) {
  var p = String(req.segunda || "").split("-").map(Number);
  if (p.length !== 3 || p.some(isNaN)) return { ok: false, erro: "Data inválida" };
  var seg = new Date(p[0], p[1] - 1, p[2]);
  seg.setDate(seg.getDate() - ((seg.getDay() + 6) % 7));
  var feitas = [];
  ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SABADO", "DOMINGO"].forEach(function (nome, i) {
    var sh = acharAba(ss, nome); if (!sh) return;
    limparAba(sh);
    var d = new Date(seg); d.setDate(seg.getDate() + i);
    sh.getRange(CELULA_DATA).setValue(d);
    feitas.push(nome);
  });
  SpreadsheetApp.flush();
  return { ok: true, abas: feitas, segunda: Utilities.formatDate(seg, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") };
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
  var linhas = [];
  for (var lin = PRIMEIRA_LINHA; lin <= ULTIMA_LINHA; lin++) {
    linhas.push(["", "=IFERROR(VLOOKUP(B" + lin + ",BANCO!A:B,2,0),\"\")", "", formulaCusto(lin), "=D" + lin + "*E" + lin]);
  }
  sh.getRange(PRIMEIRA_LINHA, 2, linhas.length, 5).setFormulas(linhas);
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
