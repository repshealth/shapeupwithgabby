/**
 * Shape Up waitlist -> "Gabby Waitlist Sign Ups" Google Sheet
 * Deployed from the sheet's Extensions > Apps Script as a Web app (Execute as: Me, Who has access: Anyone).
 */
var SHEET_ID = '1o56HLsf8SZX9G_Yf8nTzg9Gor4Jifa64db4LRr2yAQA';

function clean(v) {
  v = String(v || '').trim().slice(0, 200);
  // stop anything typed into the form being treated as a spreadsheet formula
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.company) return ContentService.createTextOutput('ok'); // spam trap filled in, ignore
  if (!p.email || !/^\S+@\S+\.\S+$/.test(p.email)) return ContentService.createTextOutput('invalid');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['Date', 'First name', 'Email', 'Source']);
    sheet.appendRow([new Date(), clean(p.name), clean(p.email), clean(p.source)]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
