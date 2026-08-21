// ============================================================
// Chilbi Herrliberg – Schichtplanung Backend
// Google Apps Script  |  Cl1.162
// Schema Konfiguration: ID|Datum|Von|Bis|Schicht|Aufgabe|Max Personen|Farbe|Informationen|Geschlossen
// Schema Anmeldungen:   ID|Name|Schicht|Aufgabe|Timestamp
// Schema Tage:          Datum|Typ
// Schema Gast:          Kürzel|Vorname|Name|Email|Tel
// ============================================================

const SHEET_ID       = '1XqTNfgONmHX9GvmfOVvb94BUo_oQ04Uy7R97FVfdWyo';
const KUERZEL_SHEET_ID = '1bK6IuVpAdLyYc9_NPbJxFMkvCZ0fNvadJABBUM-Rc0M';
const ADMIN_PW       = 'chilbi2025';
const ABRECHNUNG_SHEET_ID = '1IbXEr2UJLh6GOJsvFttVScuF3l4KpLR6hVUESqHlQ4s';
const ABRECHNUNG_PW  = 'Kohle_26';
const ABRECHNUNG_PW_VIEW = 'CHberg_26';
const SS             = SpreadsheetApp.openById(SHEET_ID);
const SS_KUERZEL     = SpreadsheetApp.openById(KUERZEL_SHEET_ID);
const SH_CONFIG      = 'Konfiguration';
const SH_SIGNUP      = 'Anmeldungen';
const SH_TAGE        = 'Tage';
const SH_GAST        = 'Gast';
const SH_EHEMALIGE   = 'Ehemalige';
const SH_FEUERWEHR   = 'Feuerwehr';
const SH_FEUERWEHREN = 'Feuerwehren';
const SH_STATUS      = 'Status';

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'abrLoad') {
    if (p.pw === ABRECHNUNG_PW) return jsonResponse(abrLoad(false));
    if (p.pw === ABRECHNUNG_PW_VIEW) return jsonResponse(abrLoad(true));
    return jsonResponse({ ok: false, error: 'Falsches Passwort' });
  }
  return jsonResponse({ config: getConfig(), signups: getSignups(), tage: getTage(), status: getStatus() });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'signup')         return jsonResponse(saveSignup(p));
    if (p.action === 'saveStatus')     return jsonResponse(saveStatus(p));
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
    if (p.action === 'kuerzelSave') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(kuerzelSave(p));
    }
    if (p.action === 'sendeSchichtMails') {
      if (p.password !== ADMIN_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(sendeSchichtMails(p));
    }
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
    if (p.action === 'abrSave') {
      if (p.password !== ABRECHNUNG_PW && p.pw !== ABRECHNUNG_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(abrSave(p));
    }
    if (p.action === 'rechnungMail') {
      if (p.password !== ABRECHNUNG_PW && p.pw !== ABRECHNUNG_PW) return jsonResponse({ ok: false, error: 'Falsches Passwort' });
      return jsonResponse(rechnungMail(p));
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

function getStatus() {
  var m = { mode:'live', vor:{}, nach:{} };
  try {
    var sh = SS.getSheetByName(SH_STATUS);
    if (!sh) return m;
    var rows = sh.getDataRange().getValues();
    var kv = {};
    for (var i=1; i<rows.length; i++){ if (rows[i][0] !== '') kv[String(rows[i][0])] = rows[i][1]; }
    m.mode = kv.mode || 'live';
    m.vor  = { titel:kv.vorTitel||'', datum:kv.vorDatum||'', ort:kv.vorOrt||'', info:kv.vorInfo||'', schluss:kv.vorSchluss||'' };
    m.nach = { eyebrow:kv.nachEyebrow||'', titel:kv.nachTitel||'', sub:kv.nachSub||'', text:kv.nachText||'', ausblick:kv.nachAusblick||'' };
  } catch(e) {}
  return m;
}

function saveStatus(p) {
  if (p.password !== ADMIN_PW) return { ok:false, error:'Falsches Passwort' };
  var o; try { o = JSON.parse(p.statusJson || '{}'); } catch(e){ return { ok:false, error:'Ungueltige Daten' }; }
  var v = o.vor || {}, n = o.nach || {};
  var pairs = [
    ['mode', o.mode || 'live'],
    ['vorTitel', v.titel||''], ['vorDatum', v.datum||''], ['vorOrt', v.ort||''], ['vorInfo', v.info||''], ['vorSchluss', v.schluss||''],
    ['nachEyebrow', n.eyebrow||''], ['nachTitel', n.titel||''], ['nachSub', n.sub||''], ['nachText', n.text||''], ['nachAusblick', n.ausblick||'']
  ];
  var sh = SS.getSheetByName(SH_STATUS);
  if (!sh) sh = SS.insertSheet(SH_STATUS);
  sh.clear();
  sh.getRange(1,1,1,2).setValues([['Key','Value']]);
  sh.getRange(2,1,pairs.length,2).setValues(pairs);
  return { ok:true };
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

// Schicht-Mails: Frontend liefert fertige Nachrichten (email, subject, html, text, ics); hier nur versenden
function sendeSchichtMails(p) {
  var msgs = p.messages || [];
  var sent = 0, fails = [];
  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    if (!m || !m.email) continue;
    try {
      var opts = { from: 'chilbi@feuerwehrmeilen.ch', name: 'Chilbi Herrliberg', htmlBody: m.html };
      var atts = [];
      if (m.ics) atts.push(Utilities.newBlob(m.ics, 'text/calendar', 'Chilbi_Schichten.ics'));
      if (m.pdf) atts.push(Utilities.newBlob(Utilities.base64Decode(m.pdf), 'application/pdf', m.pdfname || 'Dankes-Gutschein.pdf'));
      if (atts.length) opts.attachments = atts;
      GmailApp.sendEmail(m.email, m.subject, m.text, opts);
      sent++;
    } catch (e) { fails.push(m.email + ': ' + e.message); }
  }
  return { ok: true, sent: sent, fails: fails };
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
    if (p.variant === 'custom') {
      var was = p.was || 'Deinen Gutschein';
      subj = 'Dein Gutschein für die Chilbi Herrliberg';
      text = 'Hallo ' + vn + ',\n\n'
        + 'Im Anhang findest Du Deinen persönlichen Gutschein für: ' + was + '\n'
        + (p.zusatz ? p.zusatz + '\n' : '')
        + '\nWann:\n'
        + 'Freitag, 14.08.26 · 16:00–22:00\n'
        + 'Samstag, 15.08.26 · 17:00–22:00\n'
        + 'Sonntag, 16.08.26 · 11:00–21:00\n'
        + 'Wo: Füürwehr Spätzli-Beiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)\n\n'
        + 'Bitte druck den Gutschein aus und gib ihn bei uns am Stand ab.\n\n'
        + 'Wir freuen uns auf Dich!\n\nKameradschaftliche Grüsse\nEuer Chilbi-OK';
      html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.55">'
        + '<p>Hallo ' + vn + ',</p>'
        + '<p>Im Anhang findest Du Deinen persönlichen Gutschein für: <b>' + was + '</b>'
        + (p.zusatz ? '<br>' + p.zusatz : '') + '</p>'
        + '<p><b>Wann:</b><br>Freitag, 14.08.26 · 16:00–22:00<br>Samstag, 15.08.26 · 17:00–22:00<br>Sonntag, 16.08.26 · 11:00–21:00<br>'
        + '<b>Wo:</b> Füürwehr Spätzli-Beiz an der Chilbi Herrliberg (Primarschule Rebacker, Schulhausstrasse 23, 8704 Herrliberg)</p>'
        + '<p>Bitte <b>druck den Gutschein aus und gib ihn bei uns am Stand ab.</b></p>'
        + '<p>Wir freuen uns auf Dich!</p>'
        + '<p>Kameradschaftliche Grüsse<br>Euer Chilbi-OK</p></div>';
    } else if (p.variant === 'fw') {
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

// ===== Finanzen / Abrechnung (separates Sheet ABRECHNUNG_SHEET_ID) =====
function _abrNum(v){ var n = parseFloat(String(v==null?'':v).replace(',', '.')); return isNaN(n) ? 0 : n; }
function _abrReadTab(ss, name){
  var sh = ss.getSheetByName(name); if(!sh) return [];
  var vals = sh.getDataRange().getValues(); if(vals.length < 2) return [];
  var head = vals[0].map(function(h){ return String(h).trim(); });
  var out = [];
  for(var i=1;i<vals.length;i++){ var o={}; for(var j=0;j<head.length;j++){ o[head[j]] = vals[i][j]; } out.push(o); }
  return out;
}
function _abrWriteTab(ss, name, header, rows){
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if(rows.length) sh.getRange(2,1,rows.length,header.length).setValues(rows);
}
function abrLoad(readonly){
  var ss = SpreadsheetApp.openById(ABRECHNUNG_SHEET_ID);
  var meta = {};
  var mSheet = ss.getSheetByName('Meta');
  if(mSheet){ var mv = mSheet.getDataRange().getValues(); for(var i=1;i<mv.length;i++){ if(mv[i][0]) meta[String(mv[i][0]).trim()] = mv[i][1]; } }
  return { ok:true, readonly:!!readonly, meta:meta,
    einnahmen:_abrReadTab(ss,'Einnahmen'), ausgaben:_abrReadTab(ss,'Ausgaben'),
    stock:_abrReadTab(ss,'Stock'), kasse:_abrReadTab(ss,'Kasse'),
    rechnungen:_abrReadTab(ss,'Rechnungen') };
}
function abrSave(p){
  var ss = SpreadsheetApp.openById(ABRECHNUNG_SHEET_ID);
  var meta = p.meta || {};
  var mSheet = ss.getSheetByName('Meta') || ss.insertSheet('Meta');
  mSheet.clear();
  mSheet.getRange(1,1,21,2).setValues([
    ['Schlüssel','Wert'],
    ['Jahr', meta.jahr||''],
    ['Vorjahr', meta.vorjahr||''],
    ['VorjahrKasse', meta.vorjahrKasse||''],
    ['EinnahmenKategorien', (meta.einKat||[]).join(', ')],
    ['AusgabenKategorien', (meta.ausKat||[]).join(', ')],
    ['KostenKategorien', (meta.kostKat||[]).join(', ')],
    ['SollStock', meta.sollStock||''],
    ['SollKasse', meta.sollKasse||''],
    ['MuenzProRolle', meta.muenzProRolle||''],
    ['Spaetzli', meta.spaetzli||''],
    ['Helferstunden', meta.helferstunden||''],
    ['Twint', meta.twint||''],
    ['VorjahrJson', meta.vorjahrJson||''],
    ['RgAbsender', meta.rgAbsender||''],
    ['RgIntro', meta.rgIntro||''],
    ['RgVorPos', meta.rgVorPos||''],
    ['RgIban', meta.rgIban||''],
    ['RgGruss', meta.rgGruss||''],
    ['RgFooter', meta.rgFooter||''],
    ['RgNext', meta.rgNext||'']
  ]);
  _abrWriteTab(ss,'Einnahmen',['Nr','Kategorie','Beschreibung','Betrag','Vorjahr'],
    (p.einnahmen||[]).map(function(r,i){ return [i+1, r.kat||'', r.besch||'', _abrNum(r.betrag), _abrNum(r.vorjahr)]; }));
  _abrWriteTab(ss,'Ausgaben',['Nr','Zahlungsart','Beschreibung','Betrag','Vorjahr','Bemerkung','Kostenart','Vorschuss','Zurueckbezahlt'],
    (p.ausgaben||[]).map(function(r,i){ return [i+1, r.art||'', r.besch||'', _abrNum(r.betrag), _abrNum(r.vorjahr), r.bem||'', r.kostenart||'', r.vorschuss?1:'', r.rueck?1:'']; }));
  _abrWriteTab(ss,'Stock',['Wert','Einzeln','Rollen'],
    (p.stock||[]).map(function(r){ return [_abrNum(r.wert), _abrNum(r.einzeln), _abrNum(r.rollen)]; }));
  _abrWriteTab(ss,'Kasse',['Wert','Einzeln','Rollen'],
    (p.kasse||[]).map(function(r){ return [_abrNum(r.wert), _abrNum(r.einzeln), _abrNum(r.rollen)]; }));
  _abrWriteTab(ss,'Rechnungen',['Nr','Datum','Empfaenger','Email','Betreff','Total','JSON'],
    (p.rechnungen||[]).map(function(r){ return [r.nr||'', r.datum||'', r.empfaenger||'', r.email||'', r.betreff||'', _abrNum(r.total), r.json||'']; }));
  return { ok:true };
}
function rechnungMail(p){
  var opts = { from:'chilbi@feuerwehrmeilen.ch', name:'Chilbi Herrliberg', htmlBody:p.html };
  if(p.pdf) opts.attachments = [Utilities.newBlob(Utilities.base64Decode(p.pdf), 'application/pdf', p.pdfname||'Rechnung.pdf')];
  GmailApp.sendEmail(p.email, p.subject, p.text||'', opts);
  return { ok:true };
}

function kuerzelSave(p){
  _kzWrite(SS_KUERZEL,'Tabellenblatt1',['Kürzel','Vorname','Name','Email'],(p.feuerwehr||[]).map(function(r){ return [r.kuerzel||'', r.vorname||'', r.name||'', r.email||'']; }));
  _kzWrite(SS_KUERZEL,'Gast',['Kürzel','Vorname','Name','Email','Tel'],(p.gasthelfer||[]).map(function(r){ return [r.kuerzel||'', r.vorname||'', r.name||'', r.email||'', r.tel||'']; }));
  return { ok:true };
}
function _kzWrite(ss,name,headers,rows){
  var sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
}
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
