/**
 * VIRTUS TENNIS - STAFF · "Ponte" tra l'app e il foglio Google
 * ---------------------------------------------------------------
 * Copia di sicurezza: questo file vive anche in app-staff/backend/ sul Mac e su GitHub.
 *
 * COSA FA
 *  - GET  ?sheet=NomeScheda[&token=..] -> righe della scheda (i PIN non escono mai)
 *  - POST { sheet, data }              -> aggiunge una riga (l'ID lo crea lo script)
 *  - POST { sheet, action:'update', id | matchField+matchValue, data } -> modifica
 *  - POST { sheet, action:'delete', id | matchField+matchValue }       -> cancella
 *  - POST { action:'login', nome, pin }      -> controlla il PIN e dà una chiave di sessione
 *  - POST { action:'set_pin', nome, pin }    -> imposta/cambia il PIN (salvato cifrato)
 *  - POST { action:'notifica_turni', dal, al } -> avvisa chi ha turni in quel periodo
 *  - Turni, Turni_Modello, Utenti, Ferie: modificabili solo da titolare/responsabili
 *    entrati con il PIN.
 *
 * DOPO OGNI AGGIORNAMENTO DEL CODICE: esegui "setup" una volta, poi
 *   Esegui il deployment -> Gestisci deployment -> ✏️ -> Versione: "Nuova versione" -> Esegui il deployment
 * NOTIFICHE: chiave OneSignal in Impostazioni progetto -> Proprietà script, nome ONESIGNAL_API_KEY.
 */

var ONESIGNAL_APP_ID = '7ad21254-c48a-4b34-8581-08d2abd4d3c9';
var APP_URL = 'https://monicahuang1997-byte.github.io/virtus-staff-app/';
var TOKEN_DAYS = 30;

// Schede e colonne che l'app si aspetta
var SCHEMA = {
  Utenti:             ['Nome', 'Ruolo', 'Pin', 'Gestisce turni', 'Tipo', 'Ore contratto', 'Telefono'],
  Turni:              ['ID', 'Data', 'Ora inizio', 'Ora fine', 'Dipendente', 'Modello', 'Note', 'Creato da'],
  Turni_Modello:      ['ID', 'Nome', 'Ora inizio', 'Ora fine', 'Colore'],
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

// Schede che solo titolare e responsabili possono modificare
var PROTETTE = { Turni: true, Turni_Modello: true, Utenti: true, Ferie: true };

var MODELLI_INIZIALI = [
  ['Colazione',  '07:00', '12:00', 'giallo'],
  ['Pranzo',     '11:00', '15:00', 'verde'],
  ['Pomeriggio', '14:00', '19:00', 'azzurro'],
  ['Cena',       '18:00', '23:00', 'lilla'],
  ['Giornata',   '09:00', '17:00', 'rosa']
];

// ---------- SETUP (da eseguire una volta, e dopo ogni aggiornamento) ----------
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
  ['Foglio1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  var ut = ss.getSheetByName('Utenti');
  if (ut.getLastRow() < 2) appendObj_(ut, { Nome: 'Monica', Ruolo: 'titolare', Tipo: 'fisso' });
  var tm = ss.getSheetByName('Turni_Modello');
  if (tm.getLastRow() < 2) {
    MODELLI_INIZIALI.forEach(function (m, i) {
      appendObj_(tm, { ID: (Date.now() + i) + '-m', Nome: m[0], 'Ora inizio': m[1], 'Ora fine': m[2], Colore: m[3] });
    });
  }
  secret_();
  Logger.log('Setup completato: ' + Object.keys(SCHEMA).length + ' schede pronte.');
}

// ---------- LETTURA ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  var nome = p.sheet;
  if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
  var rows = readRows_(nome);
  if (nome === 'Utenti') {
    var me = userFromToken_(p.token);
    var manager = !!(me && isManager_(me));
    rows = rows.map(function (r) {
      var o = {};
      Object.keys(r).forEach(function (k) {
        var kl = k.toLowerCase();
        if (kl === 'pin') return;                     // il PIN non esce mai
        if (kl === 'telefono' && !manager) return;    // il telefono lo vede solo chi gestisce
        o[k] = r[k];
      });
      o.HaPin = String(r.Pin || '').trim() ? 'si' : '';
      return o;
    });
  }
  return json_(rows);
}

// ---------- SCRITTURA ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita che due salvataggi contemporanei si pestino i piedi
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'add';

    if (action === 'login') return json_(login_(body));
    if (action === 'set_pin') return json_(setPin_(body));
    if (action === 'notifica_turni') return json_(notificaTurni_(body));

    var nome = body.sheet;
    if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
    if (!sh) return json_({ error: 'Scheda "' + nome + '" non trovata. Esegui setup.' });

    if (PROTETTE[nome]) {
      var me = userFromToken_(body.token);
      if (!me || !isManager_(me)) {
        return json_({ error: 'Questa modifica la possono fare solo titolare e responsabili. Esci e rientra con il tuo PIN.', auth: true });
      }
    }

    var data = body.data || {};
    Object.keys(data).forEach(function (k) { if (k.toLowerCase() === 'pin') delete data[k]; }); // il PIN si cambia solo con set_pin

    if (action === 'add') {
      if (SCHEMA[nome].indexOf('ID') !== -1 && !data.ID) {
        data.ID = Date.now() + '-' + Math.floor(Math.random() * 10000);
      }
      appendObj_(sh, data);
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

// ---------- ACCESSO (PIN e chiave di sessione) ----------
function login_(body) {
  var u = getUser_(body.nome);
  if (!u) return { error: 'Nome non trovato.' };
  if (!u.pin) return { error: 'Non hai ancora un PIN: impostalo.', needPin: true };
  if (!pinOk_(u.pin, u.nome, body.pin)) return { error: 'PIN errato, riprova.' };
  return { ok: true, token: makeToken_(u.nome), ruolo: u.ruolo };
}

function setPin_(body) {
  var pin = String(body.pin || '').trim();
  if (!/^\d{4,8}$/.test(pin)) return { error: 'Il PIN deve avere da 4 a 8 cifre.' };
  var u = getUser_(body.nome);
  if (!u) return { error: 'Nome non trovato.' };
  if (u.pin) {
    var me = userFromToken_(body.token);
    if (!me || me.nome !== u.nome) return { error: 'Per cambiare il PIN devi prima entrare con quello attuale.' };
  }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Utenti');
  var col = u.headers.indexOf('pin') + 1;
  if (col < 1) { ensureHeaders_(sh, ['Pin']); col = sh.getLastColumn(); }
  sh.getRange(u.row, col).setNumberFormat('@').setValue(hashPin_(u.nome, pin));
  return { ok: true, token: makeToken_(u.nome), ruolo: u.ruolo };
}

function secret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('TOKEN_SECRET', s); }
  return s;
}
function sign_(txt) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(txt, secret_()));
}
function hashPin_(nome, pin) {
  return 'h:' + sign_('pin|' + String(nome).trim().toLowerCase() + '|' + String(pin).trim());
}
function pinOk_(stored, nome, pin) {
  stored = String(stored || '').trim();
  if (!stored || !String(pin || '').trim()) return false;
  if (stored.indexOf('h:') === 0) return stored === hashPin_(nome, pin);
  return stored === String(pin).trim(); // PIN vecchi scritti a mano nel foglio
}
function makeToken_(nome) {
  var payload = Utilities.base64EncodeWebSafe(nome + '|' + (Date.now() + TOKEN_DAYS * 86400000), Utilities.Charset.UTF_8);
  return payload + '.' + sign_('tok|' + payload);
}
function userFromToken_(token) {
  try {
    if (!token) return null;
    var parts = String(token).split('.');
    if (parts.length !== 2 || sign_('tok|' + parts[0]) !== parts[1]) return null;
    var dec = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    var i = dec.lastIndexOf('|');
    var exp = Number(dec.slice(i + 1));
    if (!exp || exp < Date.now()) return null;
    return getUser_(dec.slice(0, i));
  } catch (err) { return null; }
}
function getUser_(nome) {
  var key = String(nome || '').trim().toLowerCase();
  if (!key) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Utenti');
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var c = function (name) { return h.indexOf(name); };
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][c('nome')]).trim().toLowerCase() === key) {
      return {
        row: i + 1, headers: h,
        nome: String(v[i][c('nome')]).trim(),
        ruolo: String(c('ruolo') >= 0 ? v[i][c('ruolo')] : '').trim().toLowerCase(),
        pin: String(c('pin') >= 0 ? v[i][c('pin')] : '').trim(),
        gestisce: String(c('gestisce turni') >= 0 ? v[i][c('gestisce turni')] : '').trim()
      };
    }
  }
  return null;
}
function isManager_(u) {
  return u.ruolo === 'titolare' || u.ruolo === 'responsabile' || !!u.gestisce;
}

// ---------- NOTIFICHE ----------
function notifyAfterAdd_(nome, data) {
  try {
    if (nome === 'Bacheca') {
      var urgente = String(data['Priorità'] || '').toLowerCase() === 'urgente';
      var titolo = (urgente ? '🔴 Urgente · ' : '📣 ') + 'Bacheca · ' + (data.Autore || '');
      var dest = String(data.Destinatario || '').trim();
      sendPush_(titolo, data.Messaggio || '', dest ? [dest] : null);
    }
  } catch (err) {
    Logger.log('Notifica non inviata: ' + err.message); // la riga è comunque salvata
  }
}

function notificaTurni_(body) {
  var me = userFromToken_(body.token);
  if (!me || !isManager_(me)) return { error: 'Solo titolare e responsabili possono avvisare lo staff.', auth: true };
  var dal = String(body.dal || ''), al = String(body.al || '');
  var nomi = {};
  readRows_('Turni').forEach(function (r) {
    var d = String(r.Data || '').slice(0, 10);
    var n = String(r.Dipendente || '').trim();
    if (n && d >= dal && d <= al) nomi[n] = true;
  });
  var lista = Object.keys(nomi);
  if (!lista.length) return { ok: true, inviati: 0 };
  var attive = !!PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY');
  sendPush_('🗓️ Turni pubblicati', 'Sono pronti i turni dal ' + fmtIt_(dal) + ' al ' + fmtIt_(al) + '. Apri l\'app per vederli.', lista);
  return { ok: true, inviati: lista.length, notificheAttive: attive };
}
function fmtIt_(s) { var p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] : s; }

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
function readRows_(nome) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values.shift().map(function (h) { return String(h).trim(); });
  var tz = Session.getScriptTimeZone();
  return values
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
}

function appendObj_(sh, data) {
  var headers = ensureHeaders_(sh, Object.keys(data));
  var row = headers.map(function (h) { return data[h] != null ? String(data[h]) : ''; });
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

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
  var target = String(body.id != null ? body.id : body.matchValue).trim().toLowerCase();
  var col = headers.indexOf(field);
  if (col < 0 || !target) return -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][col]).trim().toLowerCase() === target) return i + 1;
  }
  return -1;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
