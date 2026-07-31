/**
 * I Love Meat — Prenotazioni CENA → Google Sheet
 * Incolla questo codice in Extensions → Apps Script, salva,
 * poi pubblicalo come Web App (vedi google-sheet-setup.md).
 */

// ID del foglio della CENA (scrive sempre qui, senza ambiguità)
var SHEET_ID = '1HzGRC4Jp5lSoJVxDIMrL-msBMYKkljgjkw9CsjdDg8w';

// Email del ristorante a cui arriva la notifica di ogni prenotazione
var RESTAURANT_EMAIL = 'simoneignazzi1@gmail.com';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // evita scritture sovrapposte
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var p = e.parameter || {};

    // Tutte le prenotazioni di questa landing in un'unica lista
    var sheet = ss.getSheetByName('Prenotazioni') || ss.insertSheet('Prenotazioni');

    // "Offerta" è l'ULTIMA colonna, dopo "Creata": aggiungendola in fondo le
    // righe già presenti nel foglio restano allineate. Le prenotazioni arrivate
    // prima di questa modifica avranno la cella vuota.
    var INTESTAZIONI = ['Data', 'Ora', 'Persone', 'Nome', 'Telefono', 'Email', 'Richieste', 'Privacy', 'Creata', 'Offerta'];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(INTESTAZIONI);
      sheet.setFrozenRows(1);
    } else {
      // Foglio creato prima che esistesse la colonna: la aggiunge una volta sola
      var ultimaCol = sheet.getLastColumn();
      var testate = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0];
      if (testate.indexOf('Offerta') === -1) {
        sheet.getRange(1, ultimaCol + 1).setValue('Offerta');
      }
    }

    sheet.appendRow([
      p.data || '',
      p.ora || '',
      p.persone || '',
      p.nome || '',
      p.telefono || '',
      p.email || '',
      p.richieste || '',
      p.privacy || '',
      new Date(),
      p.offerta || ''   // "Bombette" o "Degustazione", inviata dalla landing
    ]);

    // Notifica email al ristorante (se fallisce, la prenotazione resta comunque salvata)
    try {
      var subj = '🍖 ' + (p.offerta ? p.offerta + ' — ' : '') + 'Nuova prenotazione — ' + (p.nome || '') +
                 ' · ' + (p.data || '') + ' ' + (p.ora || '') +
                 ' · ' + (p.persone || '') + 'p';
      var html =
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#111;line-height:1.5">' +
          '<h2 style="margin:0 0 4px">Nuova prenotazione — I Love Meat</h2>' +
          '<p style="margin:0 0 14px;color:#666">Ricevuta dal sito delle prenotazioni</p>' +
          '<table cellpadding="7" style="border-collapse:collapse;border:1px solid #eee">' +
            trEmail_('Offerta', p.offerta) +
            trEmail_('Data', p.data) +
            trEmail_('Orario', p.ora) +
            trEmail_('Persone', p.persone) +
            trEmail_('Nome', p.nome) +
            trEmail_('Telefono', p.telefono) +
            trEmail_('Email', p.email) +
            trEmail_('Richieste', p.richieste) +
          '</table>' +
        '</div>';
      var opts = { htmlBody: html, name: 'Prenotazioni I Love Meat' };
      if (p.email && p.email.indexOf('@') > 0) opts.replyTo = p.email; // rispondi diretto al cliente
      MailApp.sendEmail(RESTAURANT_EMAIL, subj, 'Nuova prenotazione ricevuta.', opts);
    } catch (mailErr) {}

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Riga della tabella nell'email (salta i campi vuoti)
function trEmail_(label, value) {
  if (!value) return '';
  return '<tr>' +
           '<td style="border:1px solid #eee;color:#666;font-weight:bold">' + label + '</td>' +
           '<td style="border:1px solid #eee">' + value + '</td>' +
         '</tr>';
}

// Utile per verificare che l'app sia pubblicata: apri l'URL nel browser.
// La versione serve a capire da fuori QUALE codice è effettivamente pubblicato:
// salvare in Apps Script non basta, bisogna anche ridistribuire. Se qui non
// leggi "v2", la distribuzione è ferma alla versione senza colonna Offerta.
function doGet() {
  return ContentService.createTextOutput(
    'I Love Meat — endpoint prenotazioni CENA attivo · v2 (colonna Offerta)'
  );
}
