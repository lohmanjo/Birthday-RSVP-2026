/**
 * Birthday RSVP + potluck backend (JSON API).
 * Same fetch-based pattern as the MSMS page — no google.script.run.
 *
 * Setup:
 *   1. Make a Google Sheet.
 *   2. Extensions > Apps Script, paste this in as Code.gs.
 *   3. Deploy > New deployment > Web app.
 *        Execute as: Me   ·   Who has access: Anyone
 *   4. Copy the /exec URL into APPS_SCRIPT_URL at the top of the HTML page.
 */

var SHEET   = 'RSVPs';
var LOG      = 'Changelog';
var HEADERS = ['Timestamp','Name','Email','Attending','Headcount','Category','Bringing','Details','Dog','Dog Names'];

function doGet()  { return json_({ rsvps: readAll_() }); }

function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) {}
  switch (d.action) {
    case 'submitRSVP': return json_(submit_(d));
    case 'updateRSVP': return json_(update_(d));
    case 'lookupEmail': return json_(lookup_(d.email));
    default: return json_({ ok:false, error:'unknown action' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) {
    s.appendRow(HEADERS);
    s.getRange('1:1').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function logSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  if (s.getLastRow() === 0) {
    s.appendRow(['Timestamp','Name','Email','Field','From','To']);
    s.getRange('1:1').setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function readAll_() {
  var rows = sheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    out.push({
      timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'MMM d, h:mm a') : '',
      name: r[1] || '', email: String(r[2] || '').toLowerCase(), attending: r[3] || '',
      headcount: r[4] || 1, category: r[5] || '', bringing: r[6] || '', details: r[7] || '',
      dog: r[8] || '', dogNames: r[9] || ''
    });
  }
  return out;
}

function findRow_(email) {
  email = String(email || '').toLowerCase();
  if (!email) return -1;
  var rows = sheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2] || '').toLowerCase() === email) return i + 1;
  }
  return -1;
}

function submit_(d) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (findRow_(d.email) > 0) return update_(d);  // same email → update, don't duplicate
    sheet_().appendRow([
      new Date(), d.name || '', String(d.email || '').toLowerCase(), d.attending || '',
      d.headcount || 1, d.category || '', d.bringing || '', d.details || '', d.dog || '', d.dogNames || ''
    ]);
    return { ok: true, rsvps: readAll_() };
  } finally { lock.releaseLock(); }
}

function update_(d) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = sheet_(), row = findRow_(d.email);
    if (row < 0) return submit_(d);
    var cur = s.getRange(row, 1, 1, HEADERS.length).getValues()[0];

    var oldBring = cur[6] || '', newBring = d.bringing || '';
    if (oldBring !== newBring) {
      logSheet_().appendRow([new Date(), d.name || cur[1], d.email, 'Bringing',
        oldBring || '(nothing)', newBring || '(nothing)']);
    }
    s.getRange(row, 4).setValue(d.attending || cur[3]);
    s.getRange(row, 5).setValue(d.headcount || cur[4]);
    s.getRange(row, 6).setValue(d.category || '');
    s.getRange(row, 7).setValue(d.bringing || '');
    s.getRange(row, 8).setValue(d.details || '');
    s.getRange(row, 9).setValue(d.dog || '');
    s.getRange(row, 10).setValue(d.dogNames || '');
    return { ok: true, rsvps: readAll_() };
  } finally { lock.releaseLock(); }
}

function lookup_(email) {
  var row = findRow_(email);
  if (row < 0) return { found: false };
  var r = sheet_().getRange(row, 1, 1, HEADERS.length).getValues()[0];
  return {
    found: true, name: r[1], email: String(r[2]).toLowerCase(), attending: r[3],
    headcount: r[4] || 1, category: r[5], bringing: r[6], details: r[7], dog: r[8], dogNames: r[9]
  };
}
