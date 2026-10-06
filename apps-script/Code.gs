/**
 * Rilievi Seme: riepilogo su foglio Google.
 * Una riga per campo (stesse colonne del file Excel importato) e una colonna per ogni data di visita.
 * Va incollato in Estensioni > Apps Script di un foglio Google e pubblicato come "App web".
 */
var NOME_FOGLIO = 'Rilievi';
var INTESTAZIONI = ['Ditta sementiera', 'Tipo di coltura', 'Codice', 'Ettari', 'Nome agricoltore', 'Coordinate', 'Zona'];

function doGet() {
  return ContentService.createTextOutput('Rilievi Seme: collegamento attivo.');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(NOME_FOGLIO) || ss.insertSheet(NOME_FOGLIO);
    var N = INTESTAZIONI.length;

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, N).setValues([INTESTAZIONI]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setFrozenColumns(3);
    }

    // Righe: aggiorna i campi esistenti (chiave = Codice) e aggiunge i nuovi.
    var lastRow = sh.getLastRow();
    var righe = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, N).getDisplayValues() : [];
    var indice = {};
    righe.forEach(function (r, i) { indice[String(r[2])] = i; });
    (d.campi || []).forEach(function (c) {
      var v = [c.ditta, c.coltura, String(c.codice), c.ettari, c.agricoltore, c.coordinate, c.zona];
      var key = String(c.codice);
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
        if (!(String(c.codice) in indice)) return;
        var col = -1;
        for (var j = N; j < intest.length; j++) { if (intest[j] === c.data) { col = j + 1; break; } }
        if (col < 0) {
          intest.push(c.data);
          col = intest.length;
          sh.getRange(1, col).setNumberFormat('@').setValue(c.data).setFontWeight('bold');
          sh.setColumnWidth(col, 280);
        }
        sh.getRange(indice[String(c.codice)] + 2, col)
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
