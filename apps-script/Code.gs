/**
 * Rilievi Seme: riepilogo su foglio Google.
 * Una riga per campo (stesse colonne del file Excel importato) e una colonna per ogni data di visita.
 * Va incollato in Estensioni > Apps Script di un foglio Google e pubblicato come "App web".
 * Salva anche su Google Drive le foto dei rilievi (un file JPG per foto), una sottocartella per campo.
 */
var NOME_CARTELLA = 'Rilievi Seme - Report';
// 'LINK' = chiunque abbia il link alla cartella può vedere i PDF (sola lettura); 'PRIVATA' = solo chi condividi a mano da Drive.
var CONDIVISIONE = 'LINK';
var NOME_FOGLIO = 'Rilievi';
var INTESTAZIONI = ['Ditta sementiera', 'Tipo di coltura', 'Codice', 'Ettari', 'Nome agricoltore', 'Coordinate', 'Zona', 'Chiave'];

function doGet() {
  return ContentService.createTextOutput('Rilievi Seme: collegamento attivo.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.azione === 'foto' || d.azione === 'pdf') return salvaFile(d);
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(NOME_FOGLIO) || ss.insertSheet(NOME_FOGLIO);
    var N = INTESTAZIONI.length;

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, N).setValues([INTESTAZIONI]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setFrozenColumns(3);
      sh.hideColumns(N);
    }

    // Righe: aggiorna i campi esistenti (chiave = agricoltore + coltura, colonna nascosta) e aggiunge i nuovi.
    var lastRow = sh.getLastRow();
    var righe = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, N).getDisplayValues() : [];
    var indice = {};
    righe.forEach(function (r, i) { indice[String(r[N - 1])] = i; });
    (d.campi || []).forEach(function (c) {
      var v = [c.ditta, c.coltura, String(c.codice), c.ettari, c.agricoltore, c.coordinate, c.zona, c.chiave];
      var key = String(c.chiave);
      if (key in indice) {
        righe[indice[key]] = v;
      } else {
        indice[key] = righe.length;
        righe.push(v);
      }
    });
    if (righe.length) {
      sh.getRange(2, 3, righe.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, righe.length, N).setValues(righe);
    }

    // Colonne: una per data di visita.
    var celle = d.celle || [];
    if (celle.length) {
      var lastCol = Math.max(sh.getLastColumn(), N);
      var intest = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      celle.forEach(function (c) {
        if (!(String(c.chiave) in indice)) return;
        var col = -1;
        for (var j = N; j < intest.length; j++) { if (intest[j] === c.data) { col = j + 1; break; } }
        if (col < 0) {
          intest.push(c.data);
          col = intest.length;
          sh.getRange(1, col).setNumberFormat('@').setValue(c.data).setFontWeight('bold');
          sh.setColumnWidth(col, 280);
        }
        sh.getRange(indice[String(c.chiave)] + 2, col)
          .setNumberFormat('@').setWrap(true).setVerticalAlignment('top').setValue(c.testo);
      });
    }

    SpreadsheetApp.flush();
    return risposta({ ok: true });
  } catch (err) {
    return risposta({ ok: false, errore: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function risposta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function cartellaRadice() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARTELLA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(NOME_CARTELLA);
  if (CONDIVISIONE === 'LINK') f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('CARTELLA_ID', f.getId());
  return f;
}

function sottocartella(radice, nome) {
  var it = radice.getFoldersByName(nome);
  return it.hasNext() ? it.next() : radice.createFolder(nome);
}

function pulisci(t) { return String(t || '').replace(/[\\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim(); }

function salvaFile(d) {
  var radice = cartellaRadice();
  var c = d.campo || {};
  var nome = [pulisci(c.codice), pulisci(c.agricoltore), pulisci(c.coltura)].filter(String).join(' - ') || 'Senza nome';
  var cart = sottocartella(radice, nome);
  (d.elimina || []).forEach(function (n) {
    var it = cart.getFilesByName(n);
    while (it.hasNext()) it.next().setTrashed(true);
  });
  (d.file || []).forEach(function (f) {
    var it = cart.getFilesByName(f.nome);
    while (it.hasNext()) it.next().setTrashed(true);
    var blob = Utilities.newBlob(Utilities.base64Decode(f.base64), f.tipo || 'application/pdf', f.nome);
    cart.createFile(blob);
  });
  // v: 2 = questa versione dello script sa salvare le foto (l'app lo controlla).
  return risposta({ ok: true, v: 2, cartella: radice.getUrl() });
}
