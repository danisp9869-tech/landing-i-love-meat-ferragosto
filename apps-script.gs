// ID del foglio della CENA
var SHEET_ID = '1HzGRC4Jp5lSoJVxDIMrL-msBMYKkljgjkw9CsjdDg8w';
// Email del ristorante che riceve la notifica
var RESTAURANT_EMAIL = 'simoneignazzi1@gmail.com';
// Dati mostrati al cliente nell'email di conferma
var RESTAURANT_NOME     = 'I Love Meat';
var RESTAURANT_TEL      = '+39 345 688 1360';
var RESTAURANT_INDIRIZZO = 'Largo II Giugno 2, Pianezza (TO)';
var RESTAURANT_MAPS     = 'https://www.google.com/maps/search/?api=1&query=Largo+II+Giugno+2+Pianezza+TO';

// "Offerta" è l'ULTIMA colonna, dopo "Creata": messa in fondo, le righe già
// presenti nel foglio restano allineate e avranno solo la cella vuota.
var INTESTAZIONI = ['Data', 'Ora', 'Persone', 'Nome', 'Telefono', 'Email', 'Richieste', 'Privacy', 'Creata', 'Offerta'];

// Fuso del ristorante. Un progetto Apps Script senza fuso impostato gira in
// UTC: in agosto (ora legale italiana) scriverebbe due ore indietro, e una
// prenotazione ricevuta alle 00:30 finirebbe datata al giorno prima.
var TZ = 'Europe/Rome';

// "2026-08-15" → "15/08/2026". Le landing mandano la data in ISO perché è il
// formato dell'<input type="date">. Se arriva già in altra forma, la lascia.
function dataIta_(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || ''));
  return m ? m[3] + '/' + m[2] + '/' + m[1] : (v || '');
}

// Data e ora di arrivo della prenotazione, ora di Roma, giorno/mese/anno.
function creataIta_() {
  return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var p = e.parameter || {};
    var sheet = ss.getSheetByName('Prenotazioni') || ss.insertSheet('Prenotazioni');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(INTESTAZIONI);
      sheet.setFrozenRows(1);
    } else {
      // Foglio nato prima che esistesse la colonna: la aggiunge una volta sola
      var ultimaCol = sheet.getLastColumn();
      var testate = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0];
      if (testate.indexOf('Offerta') === -1) sheet.getRange(1, ultimaCol + 1).setValue('Offerta');
    }
    sheet.appendRow([
      dataIta_(p.data), p.ora || '', p.persone || '', p.nome || '', p.telefono || '',
      p.email || '', p.richieste || '', p.privacy || '', creataIta_(),
      p.offerta || ''   // "Ferragosto", "Bombette" o "Degustazione", dalla landing
    ]);

    try {
      var subj = '🍖 ' + (p.offerta ? p.offerta + ' — ' : '') + 'Nuova prenotazione — ' + (p.nome || '') + ' · ' + dataIta_(p.data) + ' ' + (p.ora || '') + ' · ' + (p.persone || '') + 'p';
      var html =
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.5">' +
          '<h2 style="margin:0 0 4px">Nuova prenotazione — I Love Meat</h2>' +
          '<p style="margin:0 0 14px;color:#666">Ricevuta dal sito delle prenotazioni</p>' +
          '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
            trEmail_('Offerta', p.offerta) +
            trEmail_('Data', dataIta_(p.data)) + trEmail_('Orario', p.ora) + trEmail_('Persone', p.persone) +
            trEmail_('Nome', p.nome) + trEmail_('Telefono', p.telefono) + trEmail_('Email', p.email) +
            trEmail_('Richieste', p.richieste) +
          '</table>' +
        '</div>';
      var opts = { htmlBody: html, name: 'Prenotazioni I Love Meat' };
      if (p.email && p.email.indexOf('@') > 0) opts.replyTo = p.email;
      MailApp.sendEmail(RESTAURANT_EMAIL, subj, 'Nuova prenotazione ricevuta.', opts);
    } catch (mailErr) {}

    // Conferma al cliente, in un try/catch suo: se il suo indirizzo è sbagliato
    // o finisce la quota di invio, la riga sul foglio e l'avviso al ristorante
    // sono già andati e non devono saltare per colpa di questa.
    try { inviaConfermaCliente_(p); } catch (clienteErr) {}

    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// L'email del cliente è facoltativa nel form: un controllo minimo evita di
// bruciare quota di invio su indirizzi scritti a caso.
function emailValida_(v) {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(String(v || '').trim());
}

// Conferma al cliente. Deve reggere tutte le landing che condividono questo
// endpoint: mostra l'orario solo se è stato scelto (a Ferragosto non si chiede)
// e il nome dell'offerta solo se la landing lo ha mandato.
function inviaConfermaCliente_(p) {
  var dest = String(p.email || '').trim();
  if (!emailValida_(dest)) return false;

  var nome = String(p.nome || '').trim();
  var primoNome = nome ? nome.split(' ')[0] : '';
  var data = dataIta_(p.data);
  var quando = data + (p.ora ? ' alle ' + p.ora : '');

  var oggetto = 'Prenotazione confermata da ' + RESTAURANT_NOME + (data ? ' — ' + data : '');

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.55;max-width:520px">' +
      '<h2 style="margin:0 0 6px">Ci vediamo da ' + RESTAURANT_NOME + '!</h2>' +
      '<p style="margin:0 0 16px">' +
        (primoNome ? 'Ciao ' + primoNome + ', abbiamo' : 'Abbiamo') +
        ' ricevuto la tua prenotazione' + (quando ? ' per <b>' + quando + '</b>' : '') + '.' +
      '</p>' +
      '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee;margin-bottom:16px">' +
        trEmail_('Offerta', p.offerta) +
        trEmail_('Data', data) +
        trEmail_('Orario', p.ora) +
        trEmail_('Persone', p.persone) +
        trEmail_('A nome di', nome) +
        trEmail_('Telefono', p.telefono) +
        trEmail_('Richieste', p.richieste) +
      '</table>' +
      '<p style="margin:0 0 6px"><b>Dove siamo</b><br>' +
        '<a href="' + RESTAURANT_MAPS + '" style="color:#d10a1c">' + RESTAURANT_INDIRIZZO + '</a></p>' +
      '<p style="margin:0 0 16px"><b>Telefono</b><br>' +
        '<a href="tel:' + RESTAURANT_TEL.replace(/\s/g, '') + '" style="color:#d10a1c">' + RESTAURANT_TEL + '</a></p>' +
      '<p style="margin:0;color:#666;font-size:13.5px">' +
        'Se devi modificare o annullare, rispondi a questa email o chiamaci. A presto!' +
      '</p>' +
    '</div>';

  var testo =
    'Ci vediamo da ' + RESTAURANT_NOME + '!\n\n' +
    (primoNome ? 'Ciao ' + primoNome + ', abbiamo' : 'Abbiamo') +
    ' ricevuto la tua prenotazione' + (quando ? ' per ' + quando : '') + '.\n\n' +
    (p.offerta ? 'Offerta: ' + p.offerta + '\n' : '') +
    (data ? 'Data: ' + data + '\n' : '') +
    (p.ora ? 'Orario: ' + p.ora + '\n' : '') +
    (p.persone ? 'Persone: ' + p.persone + '\n' : '') +
    (nome ? 'A nome di: ' + nome + '\n' : '') +
    '\n' + RESTAURANT_INDIRIZZO + '\n' + RESTAURANT_TEL + '\n\n' +
    'Se devi modificare o annullare, rispondi a questa email o chiamaci.';

  MailApp.sendEmail(dest, oggetto, testo, {
    htmlBody: html,
    name: 'Prenotazioni ' + RESTAURANT_NOME,
    replyTo: RESTAURANT_EMAIL   // le risposte del cliente vanno al ristorante
  });
  return true;
}

function trEmail_(label, value) {
  if (!value) return '';
  return '<tr><td style="border:1px solid #eee;color:#666;font-weight:bold">' + label + '</td><td style="border:1px solid #eee">' + value + '</td></tr>';
}

// La versione dice QUALE codice è davvero pubblicato: salvare non basta,
// bisogna ridistribuire. Se qui non leggi "v2", la distribuzione è vecchia.
function doGet() {
  return ContentService.createTextOutput('I Love Meat — endpoint prenotazioni attivo · v4 (conferma email al cliente)');
}
