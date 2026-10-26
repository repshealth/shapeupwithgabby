/**
 * Shape Up waitlist -> Google Sheet
 * Paste into the sheet's Extensions > Apps Script, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Send the Web app URL back so it can be added to the page.
 */
var SHEET_ID = '1dVEQ9SQKQihVnZRbPNeEK472xjjPVtan6FKXhf3Xxeg';

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
    sheet.appendRow([new Date(), clean(p.name), clean(p.email), clean(p.source)]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
