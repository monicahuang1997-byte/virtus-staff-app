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

var ONESIGNAL_APP_ID = '71900c39-f127-4a4d-9635-d0aaefc73706';
var APP_URL = 'https://monicahuang1997-byte.github.io/virtus-staff-app/';
var TOKEN_DAYS = 30;

// Schede e colonne che l'app si aspetta
var SCHEMA = {
  Utenti:             ['Nome', 'Ruolo', 'Pin', 'Gestisce turni', 'Tipo', 'Ore contratto', 'Telefono', 'Giorni ferie anno', 'Ore permesso anno'],
  Assenze:            ['ID', 'Dipendente', 'Tipo', 'Dal', 'Al', 'Dalle', 'Alle', 'Giorni', 'Ore', 'Stato', 'Nota', 'Motivo rifiuto', 'Richiesto il', 'Gestito da'],
  Tipi_Assenza:       ['ID', 'Nome', 'Colore', 'A ore', 'Scala ferie', 'Scala permessi', 'Approvazione'],
  Richieste_Turno:    ['ID', 'Dipendente', 'Data', 'Modello', 'Dalle', 'Alle', 'Riposo', 'Motivo', 'Stato', 'Risposta', 'Richiesto il', 'Gestito da'],
  Festivi:            ['ID', 'Data', 'Nome'],
  Eventi:             ['ID', 'Titolo', 'Dal', 'Al', 'Ora', 'Luogo', 'Descrizione', 'Foto', 'Note', 'Creato da', 'Pubblicato il'],
  Reazioni_Eventi:    ['ID', 'Evento', 'Persona', 'Emoji'],
  Da_Fare:            ['ID', 'Testo', 'Per', 'Entro', 'Urgente', 'Creato da', 'Creato il', 'Fatto', 'Fatto da', 'Fatto il'],
  Impostazioni:       ['Chiave', 'Valore'],
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
// Schede che solo la titolare può modificare (regole di ferie e permessi)
var SOLO_TITOLARE = { Tipi_Assenza: true, Festivi: true, Impostazioni: true, Eventi: true };

var TIPI_ASSENZA_INIZIALI = [
  // Nome, Colore, A ore, Scala ferie, Scala permessi, Approvazione
  ['Ferie',               'verde',   '',   'si', '',   'si'],
  ['Permesso',            'azzurro', 'si', '',   'si', 'si'],
  ['Malattia',            'rosa',    '',   '',   '',   ''],
  ['Riposo compensativo', 'lilla',   '',   '',   '',   'si'],
  ['Altro',               'giallo',  '',   '',   '',   'si']
];
var IMPOSTAZIONI_INIZIALI = [
  ['Giorni ferie base', '26'],
  ['Max assenti stesso giorno', '2'],
  ['Conta domeniche', 'no']
];

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
  var ta = ss.getSheetByName('Tipi_Assenza');
  if (ta.getLastRow() < 2) {
    TIPI_ASSENZA_INIZIALI.forEach(function (t, i) {
      appendObj_(ta, { ID: (Date.now() + i) + '-t', Nome: t[0], Colore: t[1], 'A ore': t[2], 'Scala ferie': t[3], 'Scala permessi': t[4], Approvazione: t[5] });
    });
  }
  var im = ss.getSheetByName('Impostazioni');
  var giaPresenti = readRows_('Impostazioni').map(function (r) { return String(r.Chiave).trim(); });
  IMPOSTAZIONI_INIZIALI.forEach(function (x) {
    if (giaPresenti.indexOf(x[0]) === -1) appendObj_(im, { Chiave: x[0], Valore: x[1] });
  });
  // Promemoria automatico ogni mattina alle 8
  var haTrigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'promemoriaAssenze'; });
  if (!haTrigger) ScriptApp.newTrigger('promemoriaAssenze').timeBased().everyDays(1).atHour(8).inTimezone('Europe/Rome').create();
  secret_();
  Logger.log('Setup completato: ' + Object.keys(SCHEMA).length + ' schede pronte.');
}

// ---------- LETTURA ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  var nome = p.sheet;
  if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
  var rows = readRows_(nome);
  if (nome === 'Assenze' || nome === 'Richieste_Turno') {
    // Privacy: la titolare e i responsabili vedono tutto, gli altri solo le proprie righe
    var io = userFromToken_(p.token);
    if (!(io && isManager_(io))) {
      var chi = String(io ? io.nome : (p.nome || '')).trim().toLowerCase();
      rows = chi ? rows.filter(function (r) { return String(r.Dipendente || '').trim().toLowerCase() === chi; }) : [];
    }
  }
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
    if (action === 'reazione') return json_(reazione_(body));
    if (action === 'foto_evento') return json_(fotoEvento_(body));

    var nome = body.sheet;
    if (!nome || !SCHEMA[nome]) return json_({ error: 'Scheda non valida' });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
    if (!sh) return json_({ error: 'Scheda "' + nome + '" non trovata. Esegui setup.' });

    if (nome === 'Assenze') return json_(assenze_(sh, body, action));
    if (nome === 'Richieste_Turno') return json_(richiesteTurno_(sh, body, action));

    if (SOLO_TITOLARE[nome]) {
      var tit = userFromToken_(body.token);
      if (!pieniPoteri_(tit)) {
        return json_({ error: 'Questa modifica la possono fare solo titolare e responsabile. Esci e rientra con il tuo PIN.', auth: !tit });
      }
    }

    if (PROTETTE[nome]) {
      var me = userFromToken_(body.token);
      if (!me || !isManager_(me)) {
        return json_({ error: 'Questa modifica la possono fare solo titolare e responsabili. Esci e rientra con il tuo PIN.', auth: true });
      }
    }

    var data = body.data || {};
    Object.keys(data).forEach(function (k) { if (k.toLowerCase() === 'pin') delete data[k]; }); // il PIN si cambia solo con set_pin

    if (action === 'add') {
      if (nome === 'Eventi') {
        var autore = userFromToken_(body.token);
        data['Creato da'] = autore ? autore.nome : (data['Creato da'] || '');
        data['Pubblicato il'] = Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd HH:mm');
      }
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
// Titolare e responsabile hanno gli stessi poteri (approvazioni, eventi, impostazioni)
function pieniPoteri_(u) {
  return !!u && (u.ruolo === 'titolare' || u.ruolo === 'responsabile');
}

// ---------- FERIE E PERMESSI ----------
function assenze_(sh, body, action) {
  var me = userFromToken_(body.token);
  var titolare = pieniPoteri_(me);
  var data = body.data || {};

  if (action === 'add') {
    var chi = String(data.Dipendente || '').trim();
    if (!titolare) {
      if (me) chi = me.nome;                         // chi è entrato col PIN chiede solo per sé
      else if (!getUser_(chi)) return { error: 'Nome non trovato.' };
    }
    if (!chi) return { error: 'Manca il nome della persona.' };
    var tipo = tipoAssenza_(data.Tipo);
    if (!tipo) return { error: 'Tipo di assenza non valido.' };
    var serveOk = String(tipo.Approvazione || '').trim().toLowerCase() === 'si';
    var row = {
      ID: Date.now() + '-' + Math.floor(Math.random() * 10000),
      Dipendente: chi, Tipo: tipo.Nome,
      Dal: String(data.Dal || ''), Al: String(data.Al || data.Dal || ''),
      Dalle: String(data.Dalle || ''), Alle: String(data.Alle || ''),
      Giorni: String(data.Giorni || ''), Ore: String(data.Ore || ''),
      Stato: titolare || !serveOk ? 'approvata' : 'in attesa',
      Nota: String(data.Nota || ''), 'Motivo rifiuto': '',
      'Richiesto il': Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd HH:mm'),
      'Gestito da': titolare ? me.nome : ''
    };
    if (!row.Dal) return { error: 'Manca la data.' };
    appendObj_(sh, row);
    try {
      var periodo = periodoTxt_(row);
      if (titolare && chi !== me.nome) {
        sendPush_('🌴 Assenza registrata', row.Tipo + ' · ' + periodo, [chi]);
      } else if (row.Stato === 'in attesa') {
        sendPush_('📝 Nuova richiesta da ' + chi, row.Tipo + ' · ' + periodo, titolari_());
      } else {
        sendPush_('ℹ️ ' + chi + ': ' + row.Tipo, periodo, titolari_());
      }
    } catch (err) { Logger.log('Notifica non inviata: ' + err.message); }
    return { ok: true, id: row.ID, stato: row.Stato };
  }

  var rowIndex = findRow_(sh, body);
  if (rowIndex < 0) return { error: 'Richiesta non trovata.' };
  var attuale = rowAsObj_(sh, rowIndex);

  if (action === 'delete') {
    var mia = me && String(attuale.Dipendente).trim().toLowerCase() === me.nome.toLowerCase();
    if (!titolare && !(mia && attuale.Stato === 'in attesa')) {
      return { error: 'Puoi annullare solo le tue richieste ancora in attesa.', auth: !me };
    }
    sh.deleteRow(rowIndex);
    return { ok: true };
  }

  if (action === 'update') {
    if (!titolare) return { error: 'Solo la titolare può approvare o modificare le assenze.', auth: !me };
    var upd = {};
    ['Stato', 'Motivo rifiuto', 'Nota', 'Dal', 'Al', 'Dalle', 'Alle', 'Giorni', 'Ore', 'Tipo'].forEach(function (k) {
      if (data[k] != null) upd[k] = String(data[k]);
    });
    if (upd.Stato) upd['Gestito da'] = me.nome;
    var hs = ensureHeaders_(sh, Object.keys(upd));
    Object.keys(upd).forEach(function (k) {
      sh.getRange(rowIndex, hs.indexOf(k) + 1).setNumberFormat('@').setValue(upd[k]);
    });
    try {
      if (upd.Stato === 'approvata') sendPush_('✅ Richiesta approvata', attuale.Tipo + ' · ' + periodoTxt_(attuale), [attuale.Dipendente]);
      if (upd.Stato === 'rifiutata') sendPush_('❌ Richiesta rifiutata', attuale.Tipo + ' · ' + periodoTxt_(attuale) + (upd['Motivo rifiuto'] ? ' — ' + upd['Motivo rifiuto'] : ''), [attuale.Dipendente]);
    } catch (err) { Logger.log('Notifica non inviata: ' + err.message); }
    return { ok: true };
  }
  return { error: 'Azione sconosciuta' };
}

// ---------- EVENTI: reazioni e foto ----------
// Reazione (🎉 👏 ❤️ 🎾): se c'è già la toglie, altrimenti la aggiunge
function reazione_(body) {
  var me = userFromToken_(body.token);
  var chi = me ? me.nome : String(body.nome || '').trim();
  if (!me && !getUser_(chi)) return { error: 'Nome non trovato.' };
  var evento = String(body.evento || ''), emoji = String(body.emoji || '');
  if (!evento || !emoji) return { error: 'Dati mancanti.' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reazioni_Eventi');
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  var ce = h.indexOf('Evento'), cp = h.indexOf('Persona'), cm = h.indexOf('Emoji');
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][ce]) === evento && String(v[i][cp]).toLowerCase() === chi.toLowerCase() && String(v[i][cm]) === emoji) {
      sh.deleteRow(i + 1);
      return { ok: true, attiva: false };
    }
  }
  appendObj_(sh, { ID: Date.now() + '-' + Math.floor(Math.random() * 10000), Evento: evento, Persona: chi, Emoji: emoji });
  return { ok: true, attiva: true };
}

// Foto dell'evento: salvata nella cartella Drive "Virtus Staff – Foto eventi", visibile a chi ha il link
function fotoEvento_(body) {
  var me = userFromToken_(body.token);
  if (!pieniPoteri_(me)) return { error: 'Solo titolare e responsabile possono caricare le foto degli eventi.', auth: !me };
  var b64 = String(body.dati || '').replace(/^data:image\/\w+;base64,/, '');
  if (!b64) return { error: 'Foto mancante.' };
  if (b64.length > 4000000) return { error: 'La foto è troppo grande.' };
  var props = PropertiesService.getScriptProperties();
  var folder = null, id = props.getProperty('FOTO_FOLDER_ID');
  if (id) { try { folder = DriveApp.getFolderById(id); } catch (e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder('Virtus Staff – Foto eventi'); props.setProperty('FOTO_FOLDER_ID', folder.getId()); }
  var nomeFile = 'evento-' + Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyyMMdd-HHmmss') + '.jpg';
  var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nomeFile));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200' };
}

// ---------- RICHIESTE DI TURNO (es. "il 3 dicembre vorrei fare la mattina") ----------
function richiesteTurno_(sh, body, action) {
  var me = userFromToken_(body.token);
  var titolare = pieniPoteri_(me);
  var data = body.data || {};

  if (action === 'add') {
    var chi = me ? me.nome : String(data.Dipendente || '').trim();
    if (!me && !getUser_(chi)) return { error: 'Nome non trovato.' };
    if (!data.Data) return { error: 'Manca il giorno.' };
    var row = {
      ID: Date.now() + '-' + Math.floor(Math.random() * 10000),
      Dipendente: chi, Data: String(data.Data).slice(0, 10),
      Modello: String(data.Modello || ''), Dalle: String(data.Dalle || ''), Alle: String(data.Alle || ''),
      Riposo: data.Riposo ? 'si' : '', Motivo: String(data.Motivo || ''),
      Stato: 'in attesa', Risposta: '',
      'Richiesto il': Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd HH:mm'), 'Gestito da': ''
    };
    appendObj_(sh, row);
    try { sendPush_('📝 Richiesta turno da ' + chi, richiestaTxt_(row), titolari_()); } catch (err) { Logger.log(err.message); }
    return { ok: true, id: row.ID, stato: row.Stato };
  }

  var rowIndex = findRow_(sh, body);
  if (rowIndex < 0) return { error: 'Richiesta non trovata.' };
  var attuale = rowAsObj_(sh, rowIndex);

  if (action === 'delete') {
    var mia = me && String(attuale.Dipendente).trim().toLowerCase() === me.nome.toLowerCase();
    if (!titolare && !(mia && attuale.Stato === 'in attesa')) return { error: 'Puoi annullare solo le tue richieste ancora in attesa.', auth: !me };
    sh.deleteRow(rowIndex);
    return { ok: true };
  }

  if (action === 'update') {
    if (!titolare) return { error: 'Solo la titolare può approvare le richieste di turno.', auth: !me };
    var upd = {};
    ['Stato', 'Risposta'].forEach(function (k) { if (data[k] != null) upd[k] = String(data[k]); });
    if (upd.Stato) upd['Gestito da'] = me.nome;
    var hs = ensureHeaders_(sh, Object.keys(upd));
    Object.keys(upd).forEach(function (k) { sh.getRange(rowIndex, hs.indexOf(k) + 1).setNumberFormat('@').setValue(upd[k]); });
    try {
      if (upd.Stato === 'approvata') sendPush_('✅ Richiesta turno approvata', richiestaTxt_(attuale), [attuale.Dipendente]);
      if (upd.Stato === 'rifiutata') sendPush_('❌ Richiesta turno non accolta', richiestaTxt_(attuale) + (upd.Risposta ? ' — ' + upd.Risposta : ''), [attuale.Dipendente]);
    } catch (err) { Logger.log(err.message); }
    return { ok: true };
  }
  return { error: 'Azione sconosciuta' };
}
function richiestaTxt_(r) {
  var cosa = r.Riposo === 'si' ? 'giorno di riposo' : ((r.Modello ? r.Modello + ' ' : '') + r.Dalle + '–' + r.Alle);
  return fmtIt_(String(r.Data).slice(0, 10)) + ' · ' + cosa + (r.Motivo ? ' (' + r.Motivo + ')' : '');
}

function tipoAssenza_(nome) {
  var key = String(nome || '').trim().toLowerCase();
  return readRows_('Tipi_Assenza').filter(function (t) { return String(t.Nome).trim().toLowerCase() === key; })[0] || null;
}
function titolari_() {
  return readRows_('Utenti')
    .filter(function (u) { var r = String(u.Ruolo || '').trim().toLowerCase(); return r === 'titolare' || r === 'responsabile'; })
    .map(function (u) { return String(u.Nome).trim(); });
}
function periodoTxt_(r) {
  var dal = fmtIt_(String(r.Dal).slice(0, 10)), al = fmtIt_(String(r.Al || r.Dal).slice(0, 10));
  var txt = dal === al ? dal : 'dal ' + dal + ' al ' + al;
  if (r.Dalle && r.Alle) txt += ' · ' + r.Dalle + '–' + r.Alle;
  return txt;
}
function rowAsObj_(sh, rowIndex) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var v = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  var o = {};
  h.forEach(function (k, i) { if (String(k).trim()) o[String(k).trim()] = String(v[i]); });
  return o;
}

// Parte da solo ogni mattina alle 8 (creato da setup): avvisa chi inizia un'assenza domani
function promemoriaAssenze() {
  var domani = Utilities.formatDate(new Date(Date.now() + 86400000), 'Europe/Rome', 'yyyy-MM-dd');
  var chi = [];
  readRows_('Assenze').forEach(function (r) {
    if (r.Stato !== 'approvata' || String(r.Dal).slice(0, 10) !== domani) return;
    sendPush_('⏰ Promemoria: domani ' + r.Tipo, periodoTxt_(r), [String(r.Dipendente).trim()]);
    chi.push(r.Dipendente + ' (' + r.Tipo + ')');
  });
  if (chi.length) sendPush_('⏰ Domani assenti', chi.join(', '), titolari_());
}

// ---------- NOTIFICHE ----------
function notifyAfterAdd_(nome, data) {
  try {
    if (nome === 'Bacheca') {
      var urgente = String(data['Priorità'] || '').toLowerCase() === 'urgente';
      var titolo = (urgente ? '🔴 Urgente · ' : '📣 ') + 'Bacheca · ' + (data.Autore || '');
      var dest = String(data.Destinatario || '').trim();
      sendPush_(titolo, data.Messaggio || '', dest ? [dest] : null);
    } else if (nome === 'Da_Fare') {
      // Per una persona: avviso solo a lei. Per tutti: avviso solo se urgente
      var per = String(data.Per || '').trim(), urg = String(data.Urgente || '') === 'si';
      var tit = (urg ? '🔴 Da fare urgente' : '✅ Nuova cosa da fare') + (data['Creato da'] ? ' · da ' + data['Creato da'] : '');
      var txt = String(data.Testo || '') + (data.Entro ? ' (entro il ' + fmtIt_(String(data.Entro).slice(0, 10)) + ')' : '');
      if (per) sendPush_(tit, txt, [per]);
      else if (urg) sendPush_(tit, txt, null);
    } else if (nome === 'Eventi') {
      var quando = fmtIt_(String(data.Dal || '').slice(0, 10)) + (data.Al && data.Al !== data.Dal ? '–' + fmtIt_(String(data.Al).slice(0, 10)) : '');
      sendPush_('📅 Nuovo evento: ' + (data.Titolo || ''), quando + (data.Ora ? ' · ' + data.Ora : '') + (data.Luogo ? ' · ' + data.Luogo : ''), null);
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
  var esito = sendPush_('🗓️ Turni pubblicati', 'Sono pronti i turni dal ' + fmtIt_(dal) + ' al ' + fmtIt_(al) + '. Apri l\'app per vederli.', lista);
  // OneSignal dice quali nomi non ha trovato (persone che non hanno ancora attivato le notifiche)
  var non = esitoNonRaggiunti_(esito, lista);
  return { ok: true, inviati: lista.length, raggiunti: lista.length - non.length, nonRaggiunti: non, notificheAttive: attive };
}
function fmtIt_(s) { var p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] : s; }

// persone = elenco di nomi (come nel foglio Utenti) oppure null = tutti
function sendPush_(titolo, testo, persone) {
  var key = String(PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY') || '').replace(/\s+/g, '');
  if (!key) return; // notifiche non ancora configurate
  // Chiavi nuove (os_v2_…) usano "Key", quelle vecchie (Legacy REST API Key) usano "Basic"
  var auth = /^os_v2_/.test(key) ? 'Key ' + key : 'Basic ' + key;
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
    headers: { Authorization: auth },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('OneSignal: ' + res.getResponseCode() + ' ' + res.getContentText());
  try { return JSON.parse(res.getContentText()); } catch (e) { return null; }
}

// Prova le notifiche: seleziona "testNotifica" e premi Esegui
// Chiede a OneSignal se una persona (per nome) ha almeno un dispositivo con le notifiche attive
function haNotificheAttive_(nome) {
  var key = String(PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY') || '').replace(/\s+/g, '');
  if (!key) return false;
  var auth = /^os_v2_/.test(key) ? 'Key ' + key : 'Basic ' + key;
  var res = UrlFetchApp.fetch('https://api.onesignal.com/apps/' + ONESIGNAL_APP_ID + '/users/by/external_id/' + encodeURIComponent(nome), {
    method: 'get', headers: { Authorization: auth }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return false;
  try {
    var subs = JSON.parse(res.getContentText()).subscriptions || [];
    return subs.some(function (x) { return /Push/i.test(String(x.type)) && x.enabled !== false && x.notification_types !== -2 && x.notification_types !== 0; });
  } catch (e) { return false; }
}
// Elenco dei nomi che non hanno notifiche attive (quindi non raggiunti)
function esitoNonRaggiunti_(esito, lista) {
  return lista.filter(function (n) { return !haNotificheAttive_(n); });
}

// Prova: notifica a UNA persona (cambia il nome qui sotto e premi Esegui)
function testNotificaPersona() {
  var nome = 'Monica';
  var esito = sendPush_('Prova per ' + nome, 'Se leggi questo, il tuo telefono è collegato al tuo nome ✅', [nome]);
  var non = esitoNonRaggiunti_(esito, [nome]);
  Logger.log('Risposta OneSignal: ' + JSON.stringify(esito));
  Logger.log(non.length ? '❌ ' + nome + ' NON raggiunta: il telefono non è collegato al nome. Apri l\'app → Altro → Attiva notifiche.' : '✅ Notifica inviata a ' + nome);
}

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
