// ID del foglio della CENA
var SHEET_ID = '1HzGRC4Jp5lSoJVxDIMrL-msBMYKkljgjkw9CsjdDg8w';
// Email del ristorante che riceve la notifica
var RESTAURANT_EMAIL = 'simoneignazzi1@gmail.com';

// Indirizzo con cui FIRMARE le email in uscita.
//
// Di default le email partono dall'account che possiede questo script
// (dani.sp9869@gmail.com): è Gmail a deciderlo, non il codice. Per farle
// partire da un altro indirizzo NON basta scriverlo qui: deve essere un alias
// VERIFICATO di quell'account (Gmail → Impostazioni → Account e importazione →
// "Invia messaggi come" → aggiungi indirizzo e conferma il codice ricevuto).
//
// Lascialo VUOTO finché l'alias non è verificato. Se ci metti un indirizzo che
// non è un alias, Gmail rifiuta l'invio e il cliente non riceve niente: per
// questo il codice controlla prima, e in caso di dubbio spedisce dall'account.
// Per vedere quali alias sono già utilizzabili: Esegui → aliasDisponibili.
var MITTENTE_ALIAS = '';
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
      var opts = opzioniInvio_({ htmlBody: html });
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

// Alias verificati dell'account, letti una volta sola per esecuzione: chiederli
// a Gmail ad ogni invio sarebbe una chiamata di rete sprecata.
var _alias = null;
function aliasUsabile_(indirizzo) {
  if (!indirizzo) return false;
  if (_alias === null) {
    // Se lo script non ha (ancora) il permesso su Gmail, si prosegue senza
    // alias invece di far fallire tutta la prenotazione.
    try { _alias = GmailApp.getAliases(); } catch (e) { _alias = []; }
  }
  return _alias.indexOf(indirizzo) !== -1;
}

// Opzioni comuni a tutte le email: nome mostrato e, se disponibile, mittente.
function opzioniInvio_(extra) {
  var o = extra || {};
  o.name = 'Prenotazioni ' + RESTAURANT_NOME;
  if (aliasUsabile_(MITTENTE_ALIAS)) o.from = MITTENTE_ALIAS;
  return o;
}

// L'email del cliente è facoltativa nel form: un controllo minimo evita di
// bruciare quota di invio su indirizzi scritti a caso.
function emailValida_(v) {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(String(v || '').trim());
}

// Testo della conferma, scelto in base all'offerta.
//
// Questo endpoint è UNO SOLO per tutte le landing, che girano in contemporanea:
// Ferragosto e Bombette sono attive insieme. Ogni landing manda il proprio tag
// nel campo "offerta" (Ferragosto / Bombette / Degustazione; la landing pranzo
// non ne manda nessuno), ed è quel tag a decidere il testo. Chi prenota da una
// landing non può quindi ricevere il messaggio di un'altra.
//
// Il confronto è normalizzato (spazi e maiuscole) perché un tag scritto
// "ferragosto " manderebbe altrimenti il messaggio sbagliato. Se il tag manca o
// non è riconosciuto si cade sul testo neutro: mai su quello di un'altra offerta.
function messaggioCliente_(p, primoNome) {
  var ciao = primoNome ? 'Ciao ' + primoNome + ' ☺️' : 'Ciao ☺️';
  var offerta = String(p.offerta || '').trim().toLowerCase();

  if (offerta === 'ferragosto') {
    return {
      oggetto: 'La tua prenotazione per Ferragosto è confermata 🔥',
      righe: [
        ciao,
        'La tua prenotazione per Ferragosto è confermata: il tuo posto è riservato. 🔥',
        'Ti aspetta la nostra Grigliata Mista a 20€: carne alla brace, tavolate piene e quella bella atmosfera da giorno di festa in compagnia.',
        '📅 15 agosto — dalle ore 12:30\n📍 ' + RESTAURANT_INDIRIZZO,
        "Un'unica cosa: se dovesse succedere qualcosa e non riuscissi a venire, faccelo sapere almeno una settimana prima, così possiamo liberare il tavolo per qualcun altro.",
        'Ci vediamo il 15! 🍖'
      ]
    };
  }

  // Bombette, degustazione, pranzo e qualsiasi tag non previsto: testo neutro,
  // con data e orario presi dalla prenotazione. Nessun riferimento a Ferragosto.
  // Per dare a Bombette un testo suo basta aggiungere qui sopra un ramo
  // `if (offerta === 'bombette')`, sulla falsariga di quello di Ferragosto.
  var data = dataIta_(p.data);
  var quando = data + (p.ora ? ' alle ' + p.ora : '');
  return {
    oggetto: 'Prenotazione confermata da ' + RESTAURANT_NOME + (data ? ' — ' + data : ''),
    righe: [
      ciao,
      'La tua prenotazione è confermata: il tuo posto è riservato. 🔥',
      (quando ? '📅 ' + quando + '\n' : '') + '📍 ' + RESTAURANT_INDIRIZZO,
      "Un'unica cosa: se dovesse succedere qualcosa e non riuscissi a venire, faccelo sapere il prima possibile, così possiamo liberare il tavolo per qualcun altro.",
      'Ci vediamo presto! 🍖'
    ]
  };
}

// Conferma al cliente: prima il messaggio, sotto il riepilogo della prenotazione.
function inviaConfermaCliente_(p) {
  var dest = String(p.email || '').trim();
  if (!emailValida_(dest)) return false;

  var nome = String(p.nome || '').trim();
  var primoNome = nome ? nome.split(' ')[0] : '';
  var msg = messaggioCliente_(p, primoNome);

  // L'indirizzo diventa un link a Maps ovunque compaia nel messaggio.
  var linkMaps = '<a href="' + RESTAURANT_MAPS + '" style="color:#d10a1c">' + RESTAURANT_INDIRIZZO + '</a>';

  var paragrafi = msg.righe.map(function (r) {
    return '<p style="margin:0 0 14px">' +
      r.split(RESTAURANT_INDIRIZZO).join(linkMaps).split('\n').join('<br>') +
      '</p>';
  }).join('');

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.55;max-width:520px">' +
      paragrafi +
      '<p style="margin:22px 0 8px;font-weight:bold">Riepilogo della prenotazione</p>' +
      '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
        trEmail_('Data', dataIta_(p.data)) +
        trEmail_('Orario', p.ora) +
        trEmail_('Persone', p.persone) +
        trEmail_('A nome di', nome) +
        trEmail_('Telefono', p.telefono) +
        trEmail_('Richieste', p.richieste) +
      '</table>' +
      '<p style="margin:16px 0 0;color:#666;font-size:13.5px">' +
        'Per modifiche rispondi a questa email o chiamaci allo ' +
        '<a href="tel:' + RESTAURANT_TEL.replace(/\s/g, '') + '" style="color:#d10a1c">' + RESTAURANT_TEL + '</a>.' +
      '</p>' +
    '</div>';

  var testo =
    msg.righe.join('\n\n') + '\n\n' +
    '— Riepilogo della prenotazione —\n' +
    (dataIta_(p.data) ? 'Data: ' + dataIta_(p.data) + '\n' : '') +
    (p.ora ? 'Orario: ' + p.ora + '\n' : '') +
    (p.persone ? 'Persone: ' + p.persone + '\n' : '') +
    (nome ? 'A nome di: ' + nome + '\n' : '') +
    (p.telefono ? 'Telefono: ' + p.telefono + '\n' : '') +
    (p.richieste ? 'Richieste: ' + p.richieste + '\n' : '') +
    '\nPer modifiche rispondi a questa email o chiamaci allo ' + RESTAURANT_TEL + '.';

  MailApp.sendEmail(dest, msg.oggetto, testo, opzioniInvio_({
    htmlBody: html,
    replyTo: RESTAURANT_EMAIL   // le risposte del cliente vanno al ristorante
  }));
  return true;
}

function trEmail_(label, value) {
  if (!value) return '';
  return '<tr><td style="border:1px solid #eee;color:#666;font-weight:bold">' + label + '</td><td style="border:1px solid #eee">' + value + '</td></tr>';
}

// Da lanciare A MANO dall'editor (Esegui → chiMandaLeMail): dice da quale
// indirizzo partono le email e quanti invii restano oggi. Non è esposta
// sull'URL pubblico apposta: l'indirizzo non deve finire in chiaro sul web.
function chiMandaLeMail() {
  var account = Session.getEffectiveUser().getEmail();
  var alias;
  try { alias = GmailApp.getAliases(); } catch (e) { alias = ['(permesso Gmail non ancora concesso)']; }
  var msg =
    'Account dello script:  ' + account +
    '\nAlias verificati:      ' + (alias.length ? alias.join(', ') : '(nessuno)') +
    '\nMITTENTE_ALIAS:        ' + (MITTENTE_ALIAS || '(vuoto)') +
    '\nLe email partiranno da: ' + (aliasUsabile_(MITTENTE_ALIAS) ? MITTENTE_ALIAS : account) +
    '\nInvii disponibili oggi: ' + MailApp.getRemainingDailyQuota();
  Logger.log(msg);
  return msg;
}

// Comodo da lanciare a mano dopo aver verificato un alias in Gmail.
function aliasDisponibili() {
  var a = GmailApp.getAliases();
  var msg = a.length ? 'Alias utilizzabili: ' + a.join(', ')
                     : 'Nessun alias verificato su questo account.';
  Logger.log(msg);
  return msg;
}

// La versione dice QUALE codice è davvero pubblicato: salvare non basta,
// bisogna ridistribuire. Se qui non leggi "v5", la distribuzione è vecchia.
function doGet() {
  return ContentService.createTextOutput('I Love Meat — endpoint prenotazioni attivo · v5 (conferma cliente, testo Ferragosto)');
}
