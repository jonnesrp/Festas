// Lista VIP do site "João Siqueira convida!"
// Cole este código em Extensões > Apps Script, dentro da planilha "Site - Festas".
// Ele faz duas coisas: entrega as festas para o site e grava os nomes da Lista VIP em abas da planilha.

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

// O site pede as festas aqui (só a aba das festas; as abas de nomes nunca saem da planilha)
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

// O site manda os nomes aqui
function doPost(e) {
  var trava = LockService.getScriptLock();
  try {
    trava.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    var festa = limpo_(d.festa, 80), nome = limpo_(d.nome, 80), ig = limpo_(d.ig, 40);
    var wa = String(d.wa || '').replace(/\D/g, '').slice(0, 13);
    var nomes = (d.nomes || []).slice(0, 40).map(function (n) { return limpo_(n, 80); }).filter(String);
    if (!festa || !nome || !nomes.length) return json_({ ok: false, erro: 'dados incompletos' });

    // só aceita festa que existe na planilha com Lista VIP = sim
    var linhas = abaFestas_().getDataRange().getDisplayValues();
    var cab = linhas[0].map(chave_);
    var cFesta = cab.indexOf('festa'), cData = -1, cVip = -1;
    cab.forEach(function (h, i) { if (h.indexOf('data inicio') === 0) cData = i; if (h.indexOf('lista vip') === 0) cVip = i; });
    var achou = null;
    for (var i = 1; i < linhas.length; i++) {
      if (chave_(linhas[i][cFesta]) === chave_(festa) && /^s(im)?$/i.test(String(linhas[i][cVip]).trim())) { achou = linhas[i]; break; }
    }
    if (!achou) return json_({ ok: false, erro: 'festa sem lista' });

    var ss = SpreadsheetApp.getActiveSpreadsheet(), agora = new Date();
    var dia = String(achou[cData] || '').slice(0, 5).replace('/', '-');
    var nomeAba = ('VIP ' + dia + ' ' + achou[cFesta]).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
    var aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      aba = ss.insertSheet(nomeAba, ss.getSheets().length);
      aba.appendRow(['Nome na lista', 'Enviado por', 'Instagram', 'WhatsApp', 'Quando']);
      aba.setFrozenRows(1);
    }
    var ja = {};
    if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) { ja[chave_(r[0])] = true; });
    var novos = nomes.filter(function (n) { var k = chave_(n); if (ja[k]) return false; ja[k] = true; return true; });
    if (novos.length) {
      aba.getRange(aba.getLastRow() + 1, 1, novos.length, 5).setValues(novos.map(function (n) { return [n, nome, ig, "'" + wa, agora]; }));
    }

    // cadastro único: uma linha por WhatsApp
    var cad = ss.getSheetByName('Cadastros');
    if (!cad) {
      cad = ss.insertSheet('Cadastros', ss.getSheets().length);
      cad.appendRow(['Nome', 'Instagram', 'WhatsApp', 'Quer agenda', 'Primeiro cadastro', 'Última lista', 'Última festa']);
      cad.setFrozenRows(1);
    }
    var linha = -1;
    if (cad.getLastRow() > 1) {
      var was = cad.getRange(2, 3, cad.getLastRow() - 1, 1).getDisplayValues();
      for (var j = 0; j < was.length; j++) if (String(was[j][0]).replace(/\D/g, '') === wa) { linha = j + 2; break; }
    }
    var quer = d.agenda ? 'sim' : 'não';
    if (linha < 0) cad.appendRow([nome, ig, "'" + wa, quer, agora, agora, achou[cFesta]]);
    else { cad.getRange(linha, 1, 1, 4).setValues([[nome, ig, "'" + wa, quer]]); cad.getRange(linha, 6, 1, 2).setValues([[agora, achou[cFesta]]]); }

    return json_({ ok: true, novos: novos.length, total: nomes.length });
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    try { trava.releaseLock(); } catch (_) {}
  }
}
