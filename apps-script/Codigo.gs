/**
 * Frio Estoque — API da planilha (Apps Script)
 *
 * Cole este arquivo inteiro em Extensões → Apps Script da planilha
 * "Frio Estoque 2026" e publique em Implantar → Nova implantação → App da Web
 * (Executar como: Eu · Quem pode acessar: Qualquer pessoa).
 *
 * O site (GitHub Pages) lê e grava por aqui:
 *   GET  ?action=data                → estoque + lançamentos
 *   POST {action:'move', ...}         → baixa (venda, transferência ou avaria) ou entrada
 *   POST {action:'revert', ...}       → estorno de um lançamento
 *   POST {action:'addProduct', ...}   → cadastro de modelo novo (senha da gerência)
 * As senhas são conferidas aqui no servidor, não só na tela.
 *
 * Aba Lançamentos: a coluna K guarda o motivo (Venda, Avaria, Transferência → loja).
 * Aba Histórico (criada sozinha): vendas de meses anteriores ao app, uma linha por
 * mês/loja/produto, para o comparativo da aba Vendas. Colunas: Mês (AAAA-MM), Loja, Código, Quantidade.
 * Uma transferência grava duas linhas: a baixa na origem e a entrada no destino (id com final -T).
 */

var ABA_ESTOQUE = 'Estoque';
var ABA_LANC = 'Lançamentos';
var LOJAS = ['02', '41', '51', '85', '87'];
var SAL = 'ar-romanos:';
var VERSAO = 3;
var ABA_HIST = 'Histórico';
var COL_MOTIVO = 11; // coluna K da aba Lançamentos

// Mesmas senhas do layout (guardadas só como hash SHA-256)
var HASH_LOJA = {
  '02': 'e38f2ce6b89faa4e150948d501fbcc088e46b027929cfe9fe7a598d8d1d14b22',
  '41': '8e96a3fdf828eb0467366cf41c94cea2fcc7293276282a306a4340cb674ab911',
  '51': 'dcd57bc1da24788001f2ef9b72e1319704fbf49287b5a4b5a0bdb4ab327790f6',
  '85': 'f899afa7fea7ede3e65915f7216b7f495755d38d69f4b889e531951f1325e022',
  '87': '500c9f8cc78e3874dbe23100e42bdea0089f9ef4ecfaae7cfbe3ca965a5a3b8e'
};
var HASH_ENTRADA = 'a800d3ccf1bebaf8def115b161386f61051d3eb696b7fe4a25076d7ec69f39cd';
var HASH_GERENCIA = 'ad8e75c1103994c2c3b8e4a33d958ea1c84ea8fc03a70c118e5844ffd53147a0';

function doGet(e) {
  return json_(lerTudo_());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'move') movimentar_(req);
    else if (req.action === 'revert') estornar_(req);
    else if (req.action === 'addProduct') cadastrar_(req);
    else throw new Error('Ação desconhecida');
    SpreadsheetApp.flush();
    var out = lerTudo_();
    out.ok = true;
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------------- leitura ---------------- */

function lerTudo_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var est = ss.getSheetByName(ABA_ESTOQUE).getDataRange().getDisplayValues();
  var cab = est[0];
  var col = colunasLojas_(cab);
  var products = [];
  for (var i = 1; i < est.length; i++) {
    var r = est[i];
    var code = String(r[0]).trim();
    if (!code) continue;
    var qty = {};
    LOJAS.forEach(function (l) { qty[l] = toInt_(r[col[l]]); });
    products.push({ code: code, name: String(r[1]).trim(), type: String(r[2]).trim(), qty: qty });
  }

  var lan = ss.getSheetByName(ABA_LANC).getDataRange().getDisplayValues();
  var log = [];
  for (var j = lan.length - 1; j >= 1 && log.length < 500; j--) {
    var l = lan[j];
    if (!l[0]) continue;
    log.push({
      id: l[0], ts: Number(String(l[0]).split('-')[0]) || 0, data: l[1], store: l[2], code: l[3],
      type: l[4], qty: toInt_(l[5]), seller: l[6], order: l[7], note: l[8], reverted: !!l[9],
      reason: l[10] || ''
    });
  }
  // vendas por mês/loja/produto: todas as baixas por venda do app + aba Histórico
  var mapa = {};
  function soma(mes, loja, cod, q) {
    if (!mes || !cod || !(q > 0) || LOJAS.indexOf(loja) < 0) return;
    var k = mes + '|' + loja + '|' + cod;
    mapa[k] = (mapa[k] || 0) + q;
  }
  for (var m = 1; m < lan.length; m++) {
    var x = lan[m];
    if (!x[0] || x[4] !== 'Baixa' || x[9]) continue;
    if (x[10] && x[10] !== 'Venda') continue;
    var ts = Number(String(x[0]).split('-')[0]) || 0;
    if (!ts) continue;
    soma(new Date(ts - 3 * 3600000).toISOString().slice(0, 7), x[2], String(x[3]).replace(/\D/g, ''), toInt_(x[5]));
  }
  var hist = abaHistorico_(ss).getDataRange().getDisplayValues();
  for (var h = 1; h < hist.length; h++) {
    var lj = String(hist[h][1]).replace(/\D/g, '');
    soma(mes_(hist[h][0]), lj ? ('0' + lj).slice(-2) : '', String(hist[h][2]).replace(/\D/g, ''), toInt_(hist[h][3]));
  }
  var sales = Object.keys(mapa).map(function (k) { var p = k.split('|'); return [p[0], p[1], p[2], mapa[k]]; });
  return { products: products, log: log, sales: sales, updated: Date.now(), v: VERSAO };
}

function abaHistorico_(ss) {
  var sh = ss.getSheetByName(ABA_HIST);
  if (sh) return sh;
  try {
    sh = ss.insertSheet(ABA_HIST);
    sh.getRange('A:C').setNumberFormat('@');
    sh.getRange(1, 1, 1, 4).setValues([['Mês (AAAA-MM)', 'Loja', 'Código', 'Quantidade']]);
  } catch (x) { sh = ss.getSheetByName(ABA_HIST); }
  return sh;
}

// aceita 2026-10, 10/2026 ou 01/10/2026
function mes_(v) {
  v = String(v);
  var a = v.match(/(\d{4})-(\d{1,2})/);
  if (a) return a[1] + '-' + ('0' + a[2]).slice(-2);
  a = v.match(/(\d{1,2})\/(\d{4})/);
  return a ? a[2] + '-' + ('0' + a[1]).slice(-2) : '';
}

function colunasLojas_(cab) {
  var col = {};
  LOJAS.forEach(function (l) {
    for (var c = 0; c < cab.length; c++) {
      if (String(cab[c]).trim().indexOf(l + ' ') === 0 || String(cab[c]).trim() === l) { col[l] = c; break; }
    }
    if (col[l] === undefined) throw new Error('Coluna da loja ' + l + ' não encontrada na aba Estoque');
  });
  return col;
}

/* ---------------- gravação ---------------- */

function movimentar_(req) {
  var loja = String(req.store || '');
  var tipo = req.type === 'Entrada' ? 'Entrada' : 'Baixa';
  var n = parseInt(req.qty, 10);
  if (LOJAS.indexOf(loja) < 0) throw new Error('Loja inválida');
  if (!(n > 0)) throw new Error('Quantidade inválida');
  if (!confere_(req.pw, HASH_LOJA[loja])) throw new Error('Senha da loja incorreta');
  if (tipo === 'Entrada' && !confere_(req.entryPw, HASH_ENTRADA)) throw new Error('Senha de entrada incorreta');

  var motivo = '';
  var destino = '';
  if (tipo === 'Baixa') {
    motivo = ['Venda', 'Transferência', 'Avaria'].indexOf(req.reason) >= 0 ? req.reason : 'Venda';
    if (!limpa_(req.seller).trim()) throw new Error('Informe o nome do vendedor');
    if (motivo === 'Venda' && !limpa_(req.order).trim()) throw new Error('Informe o nome ou nº do pedido');
    if (motivo === 'Transferência') {
      destino = String(req.dest || '');
      if (LOJAS.indexOf(destino) < 0 || destino === loja) throw new Error('Escolha a loja de destino');
    }
  }

  var celula = celulaEstoque_(String(req.code), loja);
  var atual = toInt_(celula.getValue());
  var novo = tipo === 'Baixa' ? atual - n : atual + n;
  if (novo < 0) throw new Error('Só há ' + atual + ' unidade(s) em estoque');
  var celDest = destino ? celulaEstoque_(String(req.code), destino) : null;
  celula.setValue(novo);
  if (celDest) celDest.setValue(toInt_(celDest.getValue()) + n);

  var agora = new Date();
  var data = Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  var id = agora.getTime() + '-' + loja + '-' + Math.floor(Math.random() * 1e4);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LANC);
  if (!sh.getRange(1, COL_MOTIVO).getValue()) sh.getRange(1, COL_MOTIVO).setValue('Motivo');
  sh.appendRow([
    id, data, loja, String(req.code), tipo, n, limpa_(req.seller), limpa_(req.order), limpa_(req.note), '',
    destino ? 'Transferência → ' + destino : motivo
  ]);
  if (destino) {
    sh.appendRow([
      id + '-T', data, destino, String(req.code), 'Entrada', n, limpa_(req.seller), limpa_(req.order), limpa_(req.note), '',
      'Transferência ← ' + loja
    ]);
  }
}

function cadastrar_(req) {
  if (!confereGerencia_(req.mgrPw)) throw new Error('Senha da gerência incorreta');
  var code = String(req.code || '').trim().slice(0, 20);
  var name = String(req.name || '').trim().toUpperCase().slice(0, 120);
  var num = code.replace(/\D/g, '');
  if (!num) throw new Error('Informe o código do produto');
  if (!name) throw new Error('Informe o nome do produto');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  var dados = sh.getDataRange().getDisplayValues();
  var col = colunasLojas_(dados[0]);
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).replace(/\D/g, '') === num) throw new Error('Já existe um produto com o código ' + dados[i][0]);
  }
  var linha = [];
  for (var c = 0; c < dados[0].length; c++) linha.push('');
  linha[0] = code;
  linha[1] = name;
  linha[2] = String(req.ptype || '').trim().toUpperCase().slice(0, 40);
  LOJAS.forEach(function (l) { linha[col[l]] = 0; });
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1).setNumberFormat('@'); // código como texto, para 54.334 não virar número
  sh.getRange(r, 1, 1, linha.length).setValues([linha]);
}

function estornar_(req) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LANC);
  var dados = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] !== String(req.id)) continue;
    var r = dados[i], loja = r[2], tipo = r[4], n = toInt_(r[5]);
    var transf = String(r[10] || '').indexOf('Transferência') === 0;
    var agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    if (r[9]) throw new Error('Esse lançamento já foi estornado');
    if (transf && tipo === 'Entrada') throw new Error('Transferência: o estorno é feito pela loja de origem');
    if (!confere_(req.pw, HASH_LOJA[loja])) throw new Error('Senha da loja incorreta');
    if (tipo === 'Entrada' && !confere_(req.entryPw, HASH_ENTRADA)) throw new Error('Senha de entrada incorreta');
    var celula = celulaEstoque_(r[3], loja);
    var atual = toInt_(celula.getValue());
    var novo = tipo === 'Baixa' ? atual + n : atual - n;
    if (novo < 0) throw new Error('Não dá para estornar: o saldo ficaria negativo');
    if (transf) {
      // devolve da loja de destino para a de origem, nas duas linhas
      var par = -1;
      for (var k = 1; k < dados.length; k++) if (dados[k][0] === String(req.id) + '-T') { par = k; break; }
      if (par < 0) throw new Error('Entrada da transferência não encontrada');
      var celDest = celulaEstoque_(r[3], dados[par][2]);
      var saldoDest = toInt_(celDest.getValue());
      if (saldoDest < n) throw new Error('Não dá para estornar: a loja ' + dados[par][2] + ' só tem ' + saldoDest + ' unidade(s)');
      celDest.setValue(saldoDest - n);
      sh.getRange(par + 1, 10).setValue(agora);
    }
    celula.setValue(novo);
    sh.getRange(i + 1, 10).setValue(agora);
    return;
  }
  throw new Error('Lançamento não encontrado');
}

function celulaEstoque_(code, loja) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ESTOQUE);
  var dados = sh.getDataRange().getDisplayValues();
  var col = colunasLojas_(dados[0]);
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === code) return sh.getRange(i + 1, col[loja] + 1);
  }
  throw new Error('Produto ' + code + ' não encontrado');
}

/* ---------------- utilidades ---------------- */

function confere_(senha, hash) {
  if (!senha || !hash) return false;
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    SAL + String(senha).trim().toUpperCase(), Utilities.Charset.UTF_8);
  var hex = b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
  return hex === hash;
}

function confereGerencia_(senha) {
  if (!senha) return false;
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    SAL + String(senha).replace(/\s+/g, '').toUpperCase(), Utilities.Charset.UTF_8);
  var hex = b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
  return hex === HASH_GERENCIA;
}

function toInt_(v) { var n = parseInt(String(v).replace(/\D/g, ''), 10); return isNaN(n) ? 0 : n; }
function limpa_(v) { return String(v || '').slice(0, 120); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
