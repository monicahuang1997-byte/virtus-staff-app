/**
 * VIRTUS TENNIS - STAFF · "Ponte" tra l'app e il foglio Google
 * ---------------------------------------------------------------
 * Copia di sicurezza: questo file vive anche in app-staff/backend/ sul Mac.
 *
 * COSA FA
 *  - GET  ?sheet=NomeScheda      -> restituisce le righe della scheda come elenco
 *  - POST { sheet, data }        -> aggiunge una riga (l'ID lo crea lo script)
 *  - POST { sheet, action:'update', id | matchField+matchValue, data } -> modifica
 *  - POST { sheet, action:'delete', id } -> cancella una riga
 *  - Invia le notifiche push con OneSignal (nuovo messaggio in bacheca, nuovo turno)
 *
 * PRIMA VOLTA: seleziona la funzione "setup" in alto e premi "Esegui".
 * NOTIFICHE: la chiave OneSignal va in Impostazioni progetto -> Proprietà script
 *            con nome ONESIGNAL_API_KEY (mai scriverla qui nel codice).
 */

var ONESIGNAL_APP_ID = '7ad21254-c48a-4b34-8581-08d2abd4d3c9';
var APP_URL = 'https://monicahuang1997-byte.github.io/virtus-staff-app/';

// Schede e colonne che l'app si aspetta
var SCHEMA = {
  Utenti:             ['Nome', 'Ruolo', 'Pin', 'Gestisce turni'],
  Turni:              ['ID', 'Data', 'Ora inizio', 'Ora fine', 'Dipendente', 'Note', 'Creato da'],
  Note_Turni:         ['ID', 'Dipendente', 'Nota', 'Creato da'],
  Ferie:              ['ID', 'Data', 'Dipendente', 'Note', 'Creato da'],
  Checklist:          ['Locale', 'Fase', 'Attività'],
  Checklist_Log:      ['ID', 'Data', 'Dipendente', 'Locale', 'Fase', 'Attività', 'Ora'],
  Bacheca:            ['ID', 'Data', 'Autore', 'Messaggio', 'Priorità', 'Destinatario'],
  Vendita_Prodotto:   ['Categoria', 'Nome prodotto', 'Descrizione', 'Abbinamento', 'Prezzo'],
  Scorte:             ['ID', 'Data', 'Segnalato da', 'Fornitore', 'Prodotto', 'Note', 'Stato', 'Risolto da'],
  Catalogo_Fornitori: ['Fornitore', 'Prodotto'],
  Manutenzione:       ['ID', 'Data', 'Segnalato da', 'Locale', 'Problema', 'Note', 'Stato', 'Risolto da'],
  Documenti:          ['Titolo', 'Categoria', 'Contenuto']
};

// ---------- SETUP (da eseguire una volta) ----------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (nome) {
    var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
    var cols = SCHEMA[nome];
    var existing = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    cols.forEach(function (c) {
      if (existing.indexOf(c) === -1) { existing.push(c); sh.getRange(1, existing.length).setValue(c); }
    });
    // Tutto come testo: così date, ore e PIN non vengono trasformati da Google
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(existing.length, 1)).setNumberFormat('@');
    sh.getRange(1, 1, 1, existing.length).setFontWeight('bold').setBackground('#e8e6f2');
    sh.setFrozenRows(1);
  });
  // Elimina il "Foglio1" vuoto creato da Google, se c'è
  ['Foglio1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // Primo utente, così si può entrare subito
  var ut = ss.getSheetByName('Utenti');
  if (ut.getLastRow() < 2) ut.appendRow(['Monica', 'titolare', '', '']);
  Logger.log('Setup completato: ' + Object.keys(SCHEMA).length + ' schede pronte.');
}

// ---------- LETTURA ----------
function doGet(e) {
  var nome = e && e.parameter && e.parameter.sheet;
  if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return json_([]);
  var values = sh.getDataRange().getValues();
  var headers = values.shift().map(function (h) { return String(h).trim(); });
  var tz = Session.getScriptTimeZone();
  var rows = values
    .filter(function (r) { return r.some(function (c) { return c !== ''; }); })
    .map(function (r) {
      var o = {};
      headers.forEach(function (h, i) {
        if (!h) return;
        var v = r[i];
        if (v instanceof Date) {
          // Celle che Google ha trasformato in data/ora: le riportiamo a testo
          v = (v.getFullYear() < 1900) ? Utilities.formatDate(v, tz, 'HH:mm') : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        }
        o[h] = v;
      });
      return o;
    });
  return json_(rows);
}

// ---------- SCRITTURA ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita che due salvataggi contemporanei si pestino i piedi
  try {
    var body = JSON.parse(e.postData.contents);
    var nome = body.sheet;
    if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
    if (!sh) return json_({ error: 'Scheda "' + nome + '" non trovata. Esegui setup.' });
    var action = body.action || 'add';
    var data = body.data || {};

    if (action === 'add') {
      if (SCHEMA[nome].indexOf('ID') !== -1 && !data.ID) {
        data.ID = Date.now() + '-' + Math.floor(Math.random() * 10000);
      }
      var headers = ensureHeaders_(sh, Object.keys(data));
      var row = headers.map(function (h) { return data[h] != null ? String(data[h]) : ''; });
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
      notifyAfterAdd_(nome, data);
      return json_({ ok: true, id: data.ID || null });
    }

    var rowIndex = findRow_(sh, body);
    if (rowIndex < 0) return json_({ error: 'Riga non trovata' });

    if (action === 'update') {
      var hs = ensureHeaders_(sh, Object.keys(data));
      Object.keys(data).forEach(function (k) {
        sh.getRange(rowIndex, hs.indexOf(k) + 1).setNumberFormat('@').setValue(String(data[k]));
      });
      return json_({ ok: true });
    }
    if (action === 'delete') {
      sh.deleteRow(rowIndex);
      return json_({ ok: true });
    }
    return json_({ error: 'Azione sconosciuta' });
  } catch (err) {
    return json_({ error: 'Errore: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

// ---------- NOTIFICHE ----------
function notifyAfterAdd_(nome, data) {
  try {
    if (nome === 'Bacheca') {
      var urgente = String(data['Priorità'] || '').toLowerCase() === 'urgente';
      var titolo = (urgente ? '🔴 Urgente · ' : '📣 ') + 'Bacheca · ' + (data.Autore || '');
      var dest = String(data.Destinatario || '').trim();
      sendPush_(titolo, data.Messaggio || '', dest ? [dest] : null);
    } else if (nome === 'Turni' && data.Dipendente && data.Dipendente !== data['Creato da']) {
      sendPush_('🗓️ Nuovo turno', data.Data + ' · ' + data['Ora inizio'] + '–' + data['Ora fine'], [data.Dipendente]);
    }
  } catch (err) {
    Logger.log('Notifica non inviata: ' + err.message); // la riga è comunque salvata
  }
}

// persone = elenco di nomi (come nel foglio Utenti) oppure null = tutti
function sendPush_(titolo, testo, persone) {
  var key = PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY');
  if (!key) return; // notifiche non ancora configurate
  var payload = {
    app_id: ONESIGNAL_APP_ID,
    target_channel: 'push',
    headings: { en: titolo, it: titolo },
    contents: { en: testo || ' ', it: testo || ' ' },
    url: APP_URL
  };
  if (persone && persone.length) payload.include_aliases = { external_id: persone };
  else payload.included_segments = ['Total Subscriptions'];
  var res = UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Key ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('OneSignal: ' + res.getResponseCode() + ' ' + res.getContentText());
}

// Prova le notifiche: seleziona "testNotifica" e premi Esegui
function testNotifica() {
  sendPush_('Prova Virtus Staff', 'Se leggi questo, le notifiche funzionano 🎾', null);
}

// ---------- UTILITÀ ----------
function ensureHeaders_(sh, keys) {
  var last = sh.getLastColumn();
  var headers = last ? sh.getRange(1, 1, 1, last).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  keys.forEach(function (k) {
    if (headers.indexOf(k) === -1) { headers.push(k); sh.getRange(1, headers.length).setValue(k).setFontWeight('bold'); }
  });
  return headers;
}

function findRow_(sh, body) {
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var field = body.id != null ? 'id' : String(body.matchField || '').trim().toLowerCase();
  var target = String(body.id != null ? body.id : body.matchValue).trim();
  var col = headers.indexOf(field);
  if (col < 0) return -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]).trim() === target) return i + 1;
  }
  return -1;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
