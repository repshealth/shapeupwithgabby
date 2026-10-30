/**
 * Shape Up waitlist -> "Gabby Waitlist Sign Ups" Google Sheet -> REPS dashboard (Gabby's Leads)
 * Deployed from the sheet's Extensions > Apps Script as a Web app (Execute as: Me, Who has access: Anyone).
 * Each sign-up is saved as a row, then sent to the dashboard. Column E shows whether that worked.
 * To send rows the dashboard hasn't had yet (e.g. older sign-ups), run syncToDashboard once from the editor.
 */
var SHEET_ID = '1o56HLsf8SZX9G_Yf8nTzg9Gor4Jifa64db4LRr2yAQA';
var CRM_URL = 'https://repshealthgrowth.com/api/forms';
var CRM_CLIENT = 'shape-up-with-gabby'; // Gabby's client ID in the dashboard
var HEADERS = ['Date', 'First name', 'Email', 'Source', 'Dashboard'];

function clean(v) {
  v = String(v || '').trim().slice(0, 200);
  // stop anything typed into the form being treated as a spreadsheet formula
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function getSheet() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  else if (!sheet.getRange(1, 5).getValue()) sheet.getRange(1, 5).setValue('Dashboard');
  return sheet;
}

/** Sends one sign-up to the dashboard and returns a short status for column E. */
function sendToDashboard(name, email, source) {
  try {
    var res = UrlFetchApp.fetch(CRM_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        client: CRM_CLIENT, source: 'prereg', form: 'New programmes waitlist',
        page: source, name: name, email: email
      }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) return 'Sent ' + Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM HH:mm');
    var msg = '';
    try { msg = JSON.parse(res.getContentText()).error || ''; } catch (e) { msg = res.getContentText().slice(0, 80); }
    return 'Error ' + code + (msg ? ': ' + msg : '');
  } catch (err) {
    return 'Error: ' + String(err).slice(0, 80);
  }
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.company) return ContentService.createTextOutput('ok'); // spam trap filled in, ignore
  if (!p.email || !/^\S+@\S+\.\S+$/.test(p.email)) return ContentService.createTextOutput('invalid');
  var name = clean(p.name), email = clean(p.email), source = clean(p.source);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var sheet, row;
  try {
    sheet = getSheet();
    sheet.appendRow([new Date(), name, email, source, '']);
    row = sheet.getLastRow();
  } finally {
    lock.releaseLock();
  }
  sheet.getRange(row, 5).setValue(sendToDashboard(name, email, source));
  return ContentService.createTextOutput('ok');
}

/** Run once from the editor: sends every row with an empty Dashboard column (or an earlier error). */
function syncToDashboard() {
  var sheet = getSheet();
  var last = sheet.getLastRow();
  if (last < 2) return;
  var rows = sheet.getRange(2, 1, last - 1, 5).getValues();
  rows.forEach(function (r, i) {
    var status = String(r[4] || '');
    if (status.indexOf('Sent') === 0 || !r[2]) return;
    sheet.getRange(i + 2, 5).setValue(sendToDashboard(String(r[1]), String(r[2]), String(r[3])));
  });
}
