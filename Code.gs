// ============================================================
// Chilbi Herrliberg – Schichtplanung Backend
// Google Apps Script  |  Cl1.069-dev
// Schema Konfiguration: ID|Datum|Von|Bis|Schicht|Aufgabe|Max Personen|Farbe|Informationen|Geschlossen
// Schema Anmeldungen:   ID|Name|Schicht|Aufgabe|Timestamp
// Schema Tage:          Datum|Typ
// Schema Gast:          Kürzel|Vorname|Name|Email|Tel
// ============================================================

const SHEET_ID       = '1GAxPj-6QtpkzVjNwww1L6ZERTJUje6pmsCBbtBzYq3Y';
const KUERZEL_SHEET_ID = '1bK6IuVpAdLyYc9_NPbJxFMkvCZ0fNvadJABBUM-Rc0M';
const ADMIN_PW       = 'chilbi2025';
const SS             = SpreadsheetApp.openById(SHEET_ID);
const SS_KUERZEL     = SpreadsheetApp.openById(KUERZEL_SHEET_ID);
const SH_CONFIG      = 'Konfiguration';
const SH_SIGNUP      = 'Anmeldungen';
const SH_TAGE        = 'Tage';
const SH_GAST        = 'Gast';
const SH_EHEMALIGE   = 'Ehemalige';
const SH_FEUERWEHR   = 'Feuerwehr';
const SH_FEUERWEHREN = 'Feuerwehren';

function doGet(e) {
  return jsonResponse({ config: getConfig(), signups: getSignups(), tage: getTage() });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'signup')         return jsonResponse(saveSignup(p));
    if (p.action === 'unsignup')       return jsonResponse(removeSignup(p));
    if (p.action === 'editSignup')     return jsonResponse(editSignup(p));
    if (p.action === 'registerGuest')  return jsonResponse(registerGuest(p));
    if (p.action === 'updateGuestContact') return jsonResponse(updateGuestContact(p));
    if (p.action === 'updateGuest')        return jsonResponse(updateGuest(p));
    if (p.action === 'deleteGuest')        return jsonResponse(deleteGuest(p));
    if (p.action === 'ehemaligeSignup')    return jsonResponse(ehemaligeSignup(p));
    if (p.action === 'feuerwehrSignup')    return jsonResponse(feuerwehrSignup(p));
    if (p.action === 'saveFeuerwehren') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(saveFeuerwehren(p.rows));
    }
    if (p.action === 'gutscheinMail')      return jsonResponse(gutscheinMail(p));
    if (p.action === 'saveConfig') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(saveConfig(p.rows));
    }
    if (p.action === 'saveTage') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(saveTage(p.rows));
    }
    if (p.action === 'importSignups') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(importSignups(p.rows));
    }
    return jsonResponse({ ok: false, error: 'Unbekannte Aktion' });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function getConfig() {
  const sh = SS.getSheetByName(SH_CONFIG);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function getSignups() {
  const sh = SS.getSheetByName(SH_SIGNUP);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function getTage() {
  const sh = SS.getSheetByName(SH_TAGE);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function saveSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  sh.appendRow([p.tag, p.name, p.schicht, p.aufgabe, new Date()]);
  return { ok: true };
}

function removeSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(p.tag) && String(rows[i][1]) === String(p.name)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Eintrag nicht gefunden' };
}

function editSignup(p) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(p.oldTag) && String(rows[i][1]) === String(p.oldName)) {
      sh.getRange(i + 1, 1, 1, 5).setValues([[p.tag, p.name, p.schicht, p.aufgabe, rows[i][4]]]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Eintrag nicht gefunden' };
}

function registerGuest(p) {
  const shGast = SS_KUERZEL.getSheetByName(SH_GAST);
  if (!shGast) return { ok: false, error: 'Gast-Tab nicht gefunden' };

  // Prüfe ob bereits registriert
  const rows = shGast.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === String(p.vorname).toLowerCase() &&
        String(rows[i][2]).toLowerCase() === String(p.nachname).toLowerCase()) {
      return { ok: true, kuerzel: rows[i][0] };
    }
  }

  // Neuen Gast eintragen
  shGast.appendRow([p.kuerzel, p.vorname, p.nachname, '']);
  return { ok: true, kuerzel: p.kuerzel };
}

function updateGuestContact(p) {
  const shGast = SS_KUERZEL.getSheetByName(SH_GAST);
  if (!shGast) return { ok: false };
  const rows = shGast.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.kuerzel)) {
      if (p.email) shGast.getRange(i+1, 4).setValue(p.email);
      if (p.tel)   shGast.getRange(i+1, 5).setValue(p.tel);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Kürzel nicht gefunden' };
}

function updateGuest(p) {
  const sh = SS_KUERZEL.getSheetByName(SH_GAST);
  if (!sh) return { ok: false };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.oldKuerzel)) {
      sh.getRange(i+1, 1, 1, 5).setValues([[p.kuerzel, p.vorname, p.nachname, p.email||'', p.tel||'']]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Nicht gefunden' };
}

function deleteGuest(p) {
  const sh = SS_KUERZEL.getSheetByName(SH_GAST);
  if (!sh) return { ok: false };
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(p.kuerzel)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Nicht gefunden' };
}

function saveConfig(rows) {
  const sh = SS.getSheetByName(SH_CONFIG);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 10).clearContent();
  if (rows.length > 0) {
    const data = rows.map(r => [
      r.Tag, r.Datum || r.TagLabel, r.Von, r.Bis,
      r.Schicht, r.Aufgabe, r.MaxPersonen, r.Farbe, r.Informationen || '', r.Geschlossen || '0'
    ]);
    sh.getRange(2, 1, data.length, 10).setValues(data);
  }
  return { ok: true };
}

function saveTage(rows) {
  let sh = SS.getSheetByName(SH_TAGE);
  if (!sh) {
    sh = SS.insertSheet(SH_TAGE);
    sh.getRange(1, 1, 1, 2).setValues([['Datum', 'Typ']]);
  }
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (rows.length > 0) {
    const data = rows.map(r => [r.Datum, r.Typ]);
    sh.getRange(2, 1, data.length, 2).setValues(data);
  }
  return { ok: true };
}

// ============================================================
// importSignups: ersetzt ALLE Anmeldungen (replace-all, analog saveConfig)
// Erwartet p.rows = [{tag,name,schicht,aufgabe,timestamp}]
// Timestamp aus CSV wird uebernommen, sonst jetzt.
// ============================================================
function importSignups(rows) {
  const sh = SS.getSheetByName(SH_SIGNUP);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 5).clearContent();
  if (rows && rows.length > 0) {
    const data = rows.map(r => {
      let ts = r.timestamp;
      if (!ts) {
        ts = new Date();
      } else {
        const d = new Date(ts);
        ts = isNaN(d.getTime()) ? ts : d;   // gueltiges Datum als Date, sonst Rohwert
      }
      return [r.tag, r.name, r.schicht, r.aufgabe, ts];
    });
    sh.getRange(2, 1, data.length, 5).setValues(data);
  }
  return { ok: true, count: rows ? rows.length : 0 };
}

// ============================================================
// ehemaligeSignup: Zu-/Absage Ehemaligen-Treffen -> Reiter "Ehemalige"
// Kopf: Timestamp | Anmeldung | Name | Vorname | Mail | Tel
// Legt den Reiter inkl. Kopfzeile an, falls er noch nicht existiert.
// ============================================================
function ehemaligeSignup(p) {
  let sh = SS.getSheetByName(SH_EHEMALIGE);
  if (!sh) {
    sh = SS.insertSheet(SH_EHEMALIGE);
    sh.appendRow(['Timestamp', 'Anmeldung', 'Name', 'Vorname', 'Mail', 'Tel']);
  }
  const r = sh.getLastRow() + 1;
  // Mail + Tel als Text formatieren, damit fuehrendes "+" (z.B. +41 ...) nicht als Formel gilt
  sh.getRange(r, 3, 1, 4).setNumberFormat('@');
  sh.getRange(r, 1, 1, 6).setValues([[
    new Date(), p.anmeldung || '', p.name || '', p.vorname || '', p.mail || '', p.tel || ''
  ]]);
  return { ok: true };
}

// Feuerwehrkollegen-Anmeldung -> Reiter "Feuerwehr"
function feuerwehrSignup(p) {
  let sh = SS.getSheetByName(SH_FEUERWEHR);
  if (!sh) {
    sh = SS.insertSheet(SH_FEUERWEHR);
    sh.appendRow(['Timestamp', 'Feuerwehr', 'Anmeldung', 'Grad', 'Vorname', 'Name', 'Mail', 'Tel']);
  }
  const r = sh.getLastRow() + 1;
  // Mail + Tel als Text, damit fuehrendes "+" nicht als Formel gilt
  sh.getRange(r, 7, 1, 2).setNumberFormat('@');
  sh.getRange(r, 1, 1, 8).setValues([[
    new Date(), p.feuerwehr || '', p.anmeldung || '', p.grad || '', p.vorname || '', p.name || '', p.mail || '', p.tel || ''
  ]]);
  return { ok: true };
}

// Feuerwehren-Liste (Dropdown-Quelle) speichern: Spalte A, Kopf "Feuerwehr" in A1, replace-all
function saveFeuerwehren(rows) {
  let sh = SS.getSheetByName(SH_FEUERWEHREN);
  if (!sh) sh = SS.insertSheet(SH_FEUERWEHREN);
  var list = (rows || []).map(function (v) { return String(v == null ? '' : v).trim(); })
                         .filter(function (v) { return v !== ''; });
  sh.clearContents();
  sh.getRange(1, 1).setValue('Feuerwehr');
  if (list.length) sh.getRange(2, 1, list.length, 1).setValues(list.map(function (v) { return [v]; }));
  return { ok: true, count: list.length };
}

// Gutschein per Mail senden (PDF kommt fertig als Base64 vom Frontend; Versand via Alias chilbi@feuerwehrmeilen.ch)
function gutscheinMail(p) {
  try {
    if (!p.mail) return { ok: false, error: 'Keine Mailadresse' };
    if (!p.pdf)  return { ok: false, error: 'Kein PDF' };
    var bytes = Utilities.base64Decode(p.pdf);
    var fn = ('Gutschein_' + (p.vorname || '') + '_' + (p.name || '')).replace(/[^A-Za-z0-9_]/g, '_') + '.pdf';
    var blob = Utilities.newBlob(bytes, 'application/pdf', fn);
    var vn = p.vorname || '';
    var subj, text, html;
    if (p.variant === 'fw') {
      subj = 'Dein Gutschein für die Füürwehr Spätzli-Beiz an der Chilbi Herrliberg';
      text = 'Liebe/r ' + vn + ',\n\n'
        + 'Schön, dass Du dabei bist! Im Anhang findest Du Deinen persönlichen Gutschein für eine Portion Spätzli und ein Süssgetränk/Bier für 10 CHF (nur gültig in Kombination mit einem Feuerwehr-T-Shirt / Oberteil).\n\n'
        + 'Wann:\n'
        + 'Freitag, 14.08.26 · 16:00–22:00\n'
        + 'Samstag, 15.08.26 · 17:00–22:00\n'
        + 'Sonntag, 16.08.26 · 11:00–21:00\n'
        + 'Wo: Füürwehr Spätzli-Beiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)\n\n'
        + 'Bitte druck den Gutschein aus und gib ihn bei uns am Stand ab.\n\n'
        + 'Wir freuen uns auf Dich!\n\nKameradschaftliche Grüsse\nEuer Chilbi-OK';
      html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.55">'
        + '<p>Liebe/r ' + vn + ',</p>'
        + '<p>Schön, dass Du dabei bist! Im Anhang findest Du Deinen persönlichen Gutschein für eine Portion Spätzli und ein Süssgetränk/Bier für 10 CHF (nur gültig in Kombination mit einem Feuerwehr-T-Shirt / Oberteil).</p>'
        + '<p><b>Wann:</b><br>Freitag, 14.08.26 · 16:00–22:00<br>Samstag, 15.08.26 · 17:00–22:00<br>Sonntag, 16.08.26 · 11:00–21:00<br>'
        + '<b>Wo:</b> Füürwehr Spätzli-Beiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)</p>'
        + '<p>Bitte <b>druck den Gutschein aus und gib ihn bei uns am Stand ab.</b></p>'
        + '<p>Wir freuen uns auf Dich!</p>'
        + '<p>Kameradschaftliche Grüsse<br>Euer Chilbi-OK</p></div>';
    } else {
      subj = 'Dein Gutschein fürs Ehemaligen-Treffen an der Chilbi Herrliberg';
      text = 'Liebe/r ' + vn + ',\n\n'
        + 'Schön, dass Du am Ehemaligen-Treffen dabei bist! Im Anhang findest Du Deinen persönlichen Gutschein für ein Getränk nach Wahl.\n\n'
        + 'Wann: Sonntag, 16. August 2026, 14:00–16:00 Uhr\n'
        + 'Wo: Füürwehr Spätzlibeiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)\n\n'
        + 'Bitte druck den Gutschein aus und gib ihn bei uns am Stammtisch ab.\n\n'
        + 'Wir freuen uns auf Dich!\n\nKameradschaftliche Grüsse\nEuer Chilbi-OK';
      html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.55">'
        + '<p>Liebe/r ' + vn + ',</p>'
        + '<p>Schön, dass Du am Ehemaligen-Treffen dabei bist! Im Anhang findest Du Deinen persönlichen Gutschein für ein Getränk nach Wahl.</p>'
        + '<p><b>Wann:</b> Sonntag, 16. August 2026, 14:00–16:00 Uhr<br>'
        + '<b>Wo:</b> Füürwehr Spätzlibeiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)</p>'
        + '<p>Bitte <b>druck den Gutschein aus und gib ihn bei uns am Stammtisch ab.</b></p>'
        + '<p>Wir freuen uns auf Dich!</p>'
        + '<p>Kameradschaftliche Grüsse<br>Euer Chilbi-OK</p></div>';
    }
    GmailApp.sendEmail(p.mail, subj, text, {
      from: 'chilbi@feuerwehrmeilen.ch',
      name: 'Chilbi Herrliberg',
      replyTo: 'chilbi@feuerwehrmeilen.ch',
      htmlBody: html,
      attachments: [blob]
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
