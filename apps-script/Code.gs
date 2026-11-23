// Site "João Siqueira convida!": contas, Lista VIP e base para cruzar com os check-ins.
// Cole este código em Extensões > Apps Script, dentro da planilha "Site - Festas".
//
// Abas que ele cria e mantém (uma linha = um fato, nunca apague linhas):
//   Contas      uma linha por pessoa cadastrada (ID C0001...). Senha só embaralhada (hash com sal).
//   Lista VIP   uma linha por NOME colocado em lista, de todas as festas: quem é o convidado,
//               quem colocou (ID da conta), festa, marca, data, e se entrou (check-in).
//   Check-ins   você cola aqui a exportação do botão do AZList (Festa, Edição, Convidado...).
//               Depois, em Extensões > Apps Script, rode "cruzarCheckins" para marcar quem entrou.
//   AZList      escolha a festa na célula C1 e copie a coluna A (de A2 para baixo) para colar no AZList.

var ABA_CONTAS = 'Contas', ABA_LISTA = 'Lista VIP', ABA_CHECKINS = 'Check-ins', ABA_AZ = 'AZList';
var COLS_CONTAS = ['ID', 'Nome', 'Chave do nome', 'Instagram', 'WhatsApp', 'E-mail', 'Nascimento', 'Estilos favoritos',
  'Agenda no WhatsApp', 'Agenda por e-mail', 'Veio de', 'Criada em', 'Listas enviadas', 'Última lista', 'Última festa',
  'Senha (embaralhada)', 'Sal', 'Sessões'];
var COLS_LISTA = ['ID envio', 'Data da festa', 'Marca da festa', 'Festa (edição)', 'Convidado', 'Chave do convidado',
  'É o próprio titular', 'Enviado por (ID)', 'Enviado por', 'Instagram de quem enviou', 'WhatsApp de quem enviou',
  'Enviado em', 'Entrou', 'Check-in'];
var COLS_CHECKINS = ['Festa', 'Edição', 'Data e horário da festa', 'Convidado', 'Lista', 'Promoter', 'Check-in (data e horário)', 'Hostess do check-in', 'Nome subiu na lista (data e horário)'];

// ---------- utilidades ----------
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function limpo_(v, max) {
  v = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
// chave do nome: sem acento, minúsculo, sem "de/da/do" (assim "Ana de Souza" = "ana souza")
function chave_(v) {
  return String(v).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ')
    .split(' ').filter(function (t) { return t && ['de', 'da', 'do', 'das', 'dos', 'e'].indexOf(t) < 0; }).join(' ');
}
function hash_(sal, senha) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + senha, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function aba_(nome, cols, esconder) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), a = ss.getSheetByName(nome);
  if (!a) {
    a = ss.insertSheet(nome, ss.getSheets().length);
    a.appendRow(cols); a.setFrozenRows(1); a.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    if (esconder) a.hideColumns(esconder[0], esconder[1]);
  }
  return a;
}
function abaFestas_() {
  var abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < abas.length; i++) if (String(abas[i].getRange(1, 1).getDisplayValue()).trim().toLowerCase() === 'festa') return abas[i];
  return abas[0];
}
function col_(cab, inicio) { for (var i = 0; i < cab.length; i++) if (cab[i].indexOf(inicio) === 0) return i; return -1; }
// "12/12/2026 18h" ou "12/12/2026 18:30" -> Date
function dataHora_(v) {
  var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\D+(\d{1,2})(?:h|:)?(\d{2})?)?/.exec(String(v || ''));
  return m ? new Date(+m[3], +m[2] - 1, +m[1], m[4] ? +m[4] : 23, m[5] ? +m[5] : 59) : null;
}

// ---------- o site pede as festas aqui (contas e listas nunca saem da planilha) ----------
function doGet() {
  return json_({ ok: true, linhas: abaFestas_().getDataRange().getDisplayValues() });
}

// ---------- contas ----------
function contas_() {
  var a = aba_(ABA_CONTAS, COLS_CONTAS, [16, 3]), n = a.getLastRow() - 1;
  if (n < 1) return [];
  return a.getRange(2, 1, n, COLS_CONTAS.length).getDisplayValues().map(function (r, i) {
    return { linha: i + 2, id: r[0], nome: r[1], ig: r[3], wa: String(r[4]).replace(/\D/g, ''), email: String(r[5]).toLowerCase(),
      nascimento: r[6], estilos: r[7], agenda: r[8] === 'sim', agendaEmail: r[9] === 'sim', listas: +r[12] || 0,
      senha: r[15], sal: r[16], sessoes: r[17] ? r[17].split(' ') : [] };
  });
}
function publica_(c) { return { id: c.id, nome: c.nome, ig: c.ig, wa: c.wa, email: c.email, nascimento: c.nascimento, estilos: c.estilos, agenda: c.agenda, agendaEmail: c.agendaEmail }; }
function novaSessao_(c) {
  var t = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  aba_(ABA_CONTAS, COLS_CONTAS).getRange(c.linha, 18).setValue([t].concat(c.sessoes).slice(0, 5).join(' '));   // até 5 aparelhos
  return t;
}
function porToken_(token) {
  if (!token) return null;
  var cs = contas_();
  for (var i = 0; i < cs.length; i++) if (cs[i].sessoes.indexOf(String(token)) >= 0) return cs[i];
  return null;
}
function dadosConta_(d) {
  var c = { nome: limpo_(d.nome, 80), ig: limpo_(String(d.ig || '').toLowerCase(), 40), wa: String(d.wa || '').replace(/\D/g, '').slice(0, 13),
    email: limpo_(d.email, 80).toLowerCase(), nascimento: limpo_(d.nascimento, 10), estilos: limpo_([].concat(d.estilos || []).join(', '), 200),
    agenda: !!d.agenda, agendaEmail: !!d.agendaEmail, origem: limpo_(d.origem, 60) };
  if (c.nome.split(' ').length < 2) throw new Error('Escreva nome e sobrenome.');
  if (c.wa.length < 10) throw new Error('Confira o WhatsApp, com DDD.');
  var n = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(c.nascimento);
  if (!n) throw new Error('Confira a data de nascimento.');
  var nasc = new Date(+n[3], +n[2] - 1, +n[1]), hoje = new Date(), idade = hoje.getFullYear() - nasc.getFullYear() - (hoje < new Date(hoje.getFullYear(), nasc.getMonth(), nasc.getDate()) ? 1 : 0);
  if (idade < 18) throw new Error('As festas são para maiores de 18 anos.');
  return c;
}
function linhaConta_(c) {
  return [c.nome, chave_(c.nome), c.ig, "'" + c.wa, c.email, "'" + c.nascimento, c.estilos, c.agenda ? 'sim' : 'não', c.agendaEmail ? 'sim' : 'não'];
}

// ---------- ações do site ----------
function doPost(e) {
  var trava = LockService.getScriptLock();
  try {
    trava.waitLock(20000);
    var d = JSON.parse(e.postData.contents), contas = aba_(ABA_CONTAS, COLS_CONTAS, [16, 3]), agora = new Date();

    if (d.acao === 'cadastrar') {
      var c = dadosConta_(d), todas = contas_();
      if (String(d.senha || '').length < 6) return json_({ ok: false, erro: 'A senha precisa ter pelo menos 6 caracteres.' });
      var existe = todas.filter(function (x) { return x.wa === c.wa; })[0], sal = Utilities.getUuid(), id;
      if (existe && existe.senha) return json_({ ok: false, erro: 'Esse WhatsApp já tem conta. Use "Já tenho conta".' });
      if (existe) {   // senha liberada pelo Jonnes: mesma conta, mesmo ID, senha nova
        id = existe.id;
        contas.getRange(existe.linha, 2, 1, 9).setValues([linhaConta_(c)]);
        contas.getRange(existe.linha, 16, 1, 3).setValues([[hash_(sal, d.senha), sal, '']]);
      } else {
        id = 'C' + ('000' + (todas.length + 1)).slice(-4);
        contas.appendRow([id].concat(linhaConta_(c), [c.origem, agora, 0, '', '', hash_(sal, d.senha), sal, '']));
      }
      var nova = contas_().filter(function (x) { return x.id === id; })[0];
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
      contas.getRange(eu.linha, 2, 1, 9).setValues([linhaConta_(n)]);
      n.id = eu.id;
      return json_({ ok: true, token: d.token, conta: publica_(n) });
    }

    if (d.acao === 'lista') {
      var nomes = (d.nomes || []).slice(0, 40).map(function (x) { return limpo_(x, 80); }).filter(String);
      if (!nomes.length) return json_({ ok: false, erro: 'Escreva pelo menos um nome.' });
      // a festa precisa existir na planilha com Lista VIP = sim e a lista não pode ter fechado
      var linhas = abaFestas_().getDataRange().getDisplayValues(), cab = linhas[0].map(chave_);
      var cF = cab.indexOf('festa'), cD = col_(cab, 'data inicio'), cV = col_(cab, 'lista vip'), cM = col_(cab, 'marca'), cX = col_(cab, 'lista fecha');
      var f = null;
      for (var i = 1; i < linhas.length; i++) {
        var iso = String(linhas[i][cD]).replace(/^(\d{1,2})\/(\d{1,2})\/(\d{4}).*$/, function (_, a, b, c) { return c + '-' + ('0' + b).slice(-2) + '-' + ('0' + a).slice(-2); });
        if (chave_(linhas[i][cF]) === chave_(d.festa) && (!d.data || iso === d.data) && /^s(im)?$/i.test(String(linhas[i][cV]).trim())) { f = linhas[i]; break; }
      }
      if (!f) return json_({ ok: false, erro: 'Essa festa não está com Lista VIP aberta.' });
      var fecha = cX >= 0 ? dataHora_(f[cX]) : null;
      if (fecha && agora > fecha) return json_({ ok: false, erro: 'A Lista VIP dessa festa já fechou.' });

      var lista = aba_(ABA_LISTA, COLS_LISTA), dataFesta = f[cD], marca = cM >= 0 && f[cM] ? f[cM] : f[cF];
      var ja = {};
      if (lista.getLastRow() > 1) lista.getRange(2, 1, lista.getLastRow() - 1, 6).getDisplayValues().forEach(function (r) {
        if (r[1] === dataFesta && chave_(r[3]) === chave_(f[cF])) ja[r[5]] = true;
      });
      var envio = 'E' + Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyMMddHHmmss') + '-' + eu.id;
      var minha = chave_(eu.nome);
      var novos = [];
      nomes.forEach(function (x) {
        var k = chave_(x); if (!k || ja[k]) return; ja[k] = true;
        novos.push([envio, "'" + dataFesta, marca, f[cF], x, k, k === minha ? 'sim' : 'não', eu.id, eu.nome, eu.ig, "'" + eu.wa, agora, '', '']);
      });
      if (novos.length) lista.getRange(lista.getLastRow() + 1, 1, novos.length, COLS_LISTA.length).setValues(novos);
      contas.getRange(eu.linha, 13, 1, 3).setValues([[eu.listas + 1, agora, f[cF]]]);
      prepararAZList_();
      return json_({ ok: true, novos: novos.length, total: nomes.length });
    }

    return json_({ ok: false, erro: 'ação desconhecida' });
  } catch (err) {
    return json_({ ok: false, erro: err && err.message ? err.message : String(err) });
  } finally {
    try { trava.releaseLock(); } catch (_) {}
  }
}

// ---------- aba AZList: escolha a festa em B1 e copie a coluna A ----------
function prepararAZList_() {
  var a = aba_(ABA_AZ, ['Nomes para colar no AZList', 'Festa:']);
  var lista = aba_(ABA_LISTA, COLS_LISTA), n = lista.getLastRow() - 1;
  if (n < 1) return;
  var opcoes = {};
  lista.getRange(2, 2, n, 3).getDisplayValues().forEach(function (r) { opcoes[r[0] + ' · ' + r[2]] = true; });
  var regra = SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(opcoes), true).build();
  a.getRange('C1').setDataValidation(regra);
  if (!a.getRange('C1').getValue()) a.getRange('C1').setValue(Object.keys(opcoes).pop());
  a.getRange('A2').setFormula("=IFERROR(SORT(FILTER('Lista VIP'!E2:E, 'Lista VIP'!B2:B&\" · \"&'Lista VIP'!D2:D=C1)), \"\")");
}

// ---------- cruzar check-ins: cole a exportação do AZList na aba Check-ins e rode esta função ----------
function cruzarCheckins() {
  var lista = aba_(ABA_LISTA, COLS_LISTA), chk = aba_(ABA_CHECKINS, COLS_CHECKINS);
  var nl = lista.getLastRow() - 1, nc = chk.getLastRow() - 1;
  if (nl < 1 || nc < 1) return;
  var cab = chk.getRange(1, 1, 1, chk.getLastColumn()).getDisplayValues()[0].map(chave_);
  var cC = col_(cab, 'convidado'), cCk = col_(cab, 'check in'), cDH = -1;
  cab.forEach(function (h, i) { if (cDH < 0 && h.indexOf('data') === 0 && h.indexOf('festa') > 0) cDH = i; });
  var entrou = {}, diasComCheckin = {};
  chk.getRange(2, 1, nc, chk.getLastColumn()).getDisplayValues().forEach(function (r) {
    var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(r[cDH]); if (!m) return;
    var dia = ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
    diasComCheckin[dia] = true;
    entrou[dia + '|' + chave_(r[cC])] = r[cCk] || 'sim';
  });
  var dados = lista.getRange(2, 1, nl, COLS_LISTA.length).getDisplayValues();
  var saida = dados.map(function (r) {
    var m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(r[1]), dia = m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3] : '';
    if (entrou[dia + '|' + r[5]]) return ['sim', entrou[dia + '|' + r[5]]];
    return diasComCheckin[dia] ? ['não', ''] : [r[12], r[13]];   // só marca "não" se já colou os check-ins dessa data
  });
  lista.getRange(2, 13, nl, 2).setValues(saida);
}

// ---------- quando alguém esquecer a senha ----------
// Escreva o WhatsApp da pessoa (só números) abaixo e clique em Executar.
// A senha é apagada; a pessoa cria de novo em "Criar conta" com o mesmo WhatsApp e o histórico continua no mesmo ID.
function liberarSenha() {
  var WHATSAPP = '';   // ex.: '11987654321'
  var c = contas_().filter(function (x) { return x.wa === WHATSAPP; })[0];
  if (!c) throw new Error('Não achei esse WhatsApp na aba Contas.');
  aba_(ABA_CONTAS, COLS_CONTAS).getRange(c.linha, 16, 1, 3).setValues([['', '', '']]);
}
