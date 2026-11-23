// Site "João Siqueira convida!": contas e Lista VIP gravando na planilha.
// Cole este código em Extensões > Apps Script, dentro da planilha "Site - Festas".
// Ele faz três coisas: entrega as festas para o site, guarda as contas (aba "Contas")
// e grava os nomes da Lista VIP em uma aba por festa ("VIP dd-mm Festa", nomes na coluna A, prontos para colar no AZList).
// As senhas NUNCA ficam visíveis: só um código embaralhado (hash) com sal.

function abaFestas_() {
  var abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < abas.length; i++) {
    if (String(abas[i].getRange(1, 1).getDisplayValue()).trim().toLowerCase() === 'festa') return abas[i];
  }
  return abas[0];
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// O site pede as festas aqui (só a aba das festas; contas e listas nunca saem da planilha)
function doGet() {
  return json_({ ok: true, linhas: abaFestas_().getDataRange().getDisplayValues() });
}

function limpo_(v, max) {
  v = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function chave_(v) {
  return String(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function hash_(sal, senha) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + senha, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

// ---- aba Contas: uma linha por WhatsApp ----
var COLS = ['Nome', 'Instagram', 'WhatsApp', 'E-mail', 'Quer agenda', 'Criada em', 'Última lista', 'Última festa', 'Senha (embaralhada)', 'Sal', 'Sessões'];
function abaContas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), aba = ss.getSheetByName('Contas');
  if (!aba) {
    aba = ss.insertSheet('Contas', ss.getSheets().length);
    aba.appendRow(COLS); aba.setFrozenRows(1); aba.hideColumns(9, 3);
  }
  return aba;
}
function contas_() {
  var aba = abaContas_(), n = aba.getLastRow() - 1;
  if (n < 1) return [];
  return aba.getRange(2, 1, n, COLS.length).getDisplayValues().map(function (r, i) {
    return { linha: i + 2, nome: r[0], ig: r[1], wa: String(r[2]).replace(/\D/g, ''), email: String(r[3]).toLowerCase(), agenda: r[4] === 'sim', senha: r[8], sal: r[9], sessoes: r[10] ? r[10].split(' ') : [] };
  });
}
function publica_(c) { return { nome: c.nome, ig: c.ig, wa: c.wa, email: c.email, agenda: c.agenda }; }
function novaSessao_(c) {
  var t = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  var s = [t].concat(c.sessoes).slice(0, 5);   // até 5 aparelhos logados
  abaContas_().getRange(c.linha, 11).setValue(s.join(' '));
  return t;
}
function porToken_(token) {
  if (!token) return null;
  var cs = contas_();
  for (var i = 0; i < cs.length; i++) if (cs[i].sessoes.indexOf(String(token)) >= 0) return cs[i];
  return null;
}
function dadosConta_(d) {
  var c = { nome: limpo_(d.nome, 80), ig: limpo_(d.ig, 40), wa: String(d.wa || '').replace(/\D/g, '').slice(0, 13), email: limpo_(d.email, 80).toLowerCase(), agenda: !!d.agenda };
  if (c.nome.split(' ').length < 2) throw new Error('Escreva nome e sobrenome.');
  if (c.wa.length < 10) throw new Error('Confira o WhatsApp, com DDD.');
  return c;
}

function doPost(e) {
  var trava = LockService.getScriptLock();
  try {
    trava.waitLock(20000);
    var d = JSON.parse(e.postData.contents), aba = abaContas_(), agora = new Date();

    if (d.acao === 'cadastrar') {
      var c = dadosConta_(d);
      if (String(d.senha || '').length < 6) return json_({ ok: false, erro: 'A senha precisa ter pelo menos 6 caracteres.' });
      if (contas_().some(function (x) { return x.wa === c.wa; })) return json_({ ok: false, erro: 'Esse WhatsApp já tem conta. Use "Já tenho conta".' });
      var sal = Utilities.getUuid();
      aba.appendRow([c.nome, c.ig, "'" + c.wa, c.email, c.agenda ? 'sim' : 'não', agora, '', '', hash_(sal, d.senha), sal, '']);
      var nova = contas_().filter(function (x) { return x.wa === c.wa; })[0];
      return json_({ ok: true, token: novaSessao_(nova), conta: publica_(nova) });
    }

    if (d.acao === 'entrar') {
      var login = String(d.login || '').trim().toLowerCase(), wa = login.replace(/\D/g, '');
      var achada = contas_().filter(function (x) { return (wa.length >= 10 && x.wa === wa) || (x.email && x.email === login); })[0];
      if (!achada || hash_(achada.sal, String(d.senha || '')) !== achada.senha) return json_({ ok: false, erro: 'WhatsApp/e-mail ou senha não conferem.' });
      return json_({ ok: true, token: novaSessao_(achada), conta: publica_(achada) });
    }

    var eu = porToken_(d.token);
    if (!eu) return json_({ ok: false, erro: 'sessao' });

    if (d.acao === 'editar') {
      var n = dadosConta_(d);
      aba.getRange(eu.linha, 1, 1, 5).setValues([[n.nome, n.ig, "'" + n.wa, n.email, n.agenda ? 'sim' : 'não']]);
      n.linha = eu.linha; n.sessoes = eu.sessoes;
      return json_({ ok: true, token: d.token, conta: publica_(n) });
    }

    if (d.acao === 'lista') {
      var nomes = (d.nomes || []).slice(0, 40).map(function (x) { return limpo_(x, 80); }).filter(String);
      if (!nomes.length) return json_({ ok: false, erro: 'Escreva pelo menos um nome.' });
      // só aceita festa que existe na planilha com Lista VIP = sim
      var linhas = abaFestas_().getDataRange().getDisplayValues(), cab = linhas[0].map(chave_);
      var cFesta = cab.indexOf('festa'), cData = -1, cVip = -1;
      cab.forEach(function (h, i) { if (h.indexOf('data inicio') === 0) cData = i; if (h.indexOf('lista vip') === 0) cVip = i; });
      var achou = null;
      for (var i = 1; i < linhas.length; i++) {
        if (chave_(linhas[i][cFesta]) === chave_(d.festa) && /^s(im)?$/i.test(String(linhas[i][cVip]).trim())) { achou = linhas[i]; break; }
      }
      if (!achou) return json_({ ok: false, erro: 'Essa festa não está com Lista VIP aberta.' });
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var dia = String(achou[cData] || '').slice(0, 5).replace('/', '-');
      var nomeAba = ('VIP ' + dia + ' ' + achou[cFesta]).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
      var lista = ss.getSheetByName(nomeAba);
      if (!lista) {
        lista = ss.insertSheet(nomeAba, ss.getSheets().length);
        lista.appendRow(['Nome na lista', 'Enviado por', 'Instagram', 'WhatsApp', 'Quando']);
        lista.setFrozenRows(1);
      }
      var ja = {};
      if (lista.getLastRow() > 1) lista.getRange(2, 1, lista.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) { ja[chave_(r[0])] = true; });
      var novos = nomes.filter(function (x) { var k = chave_(x); if (ja[k]) return false; ja[k] = true; return true; });
      if (novos.length) lista.getRange(lista.getLastRow() + 1, 1, novos.length, 5).setValues(novos.map(function (x) { return [x, eu.nome, eu.ig, "'" + eu.wa, agora]; }));
      aba.getRange(eu.linha, 7, 1, 2).setValues([[agora, achou[cFesta]]]);
      return json_({ ok: true, novos: novos.length, total: nomes.length });
    }

    return json_({ ok: false, erro: 'ação desconhecida' });
  } catch (err) {
    return json_({ ok: false, erro: err && err.message ? err.message : String(err) });
  } finally {
    try { trava.releaseLock(); } catch (_) {}
  }
}

// Para o Jonnes, quando alguém esquecer a senha: escreva o WhatsApp da pessoa (só números)
// abaixo e clique em Executar. A conta é apagada e a pessoa cria de novo com o mesmo WhatsApp.
function liberarConta() {
  var WHATSAPP = '';   // ex.: '11987654321'
  var c = contas_().filter(function (x) { return x.wa === WHATSAPP; })[0];
  if (!c) throw new Error('Não achei esse WhatsApp na aba Contas.');
  abaContas_().deleteRow(c.linha);
}
