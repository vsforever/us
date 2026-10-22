/**
 * RSVP backend for the Srishti & Vivek invite (Google Apps Script web app).
 *
 * Each RSVP is appended to the "RSVPs" tab of the Google Sheet this script is bound to,
 * then the whole sheet is emailed as one Excel (.xlsx) file to NOTIFY_EMAIL.
 *
 * Setup (either way works):
 *   A) Google Sheet → Extensions → Apps Script → paste this file (leave SPREADSHEET_ID empty).
 *   B) script.google.com → New project → paste this file → set SPREADSHEET_ID to the ID in the
 *      sheet URL: docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
 * Then Deploy → New deployment → type "Web app", Execute as "Me", Who has access "Anyone"
 * → copy the /exec URL into RSVP_ENDPOINT in index.html.
 */
const SPREADSHEET_ID = ''; // only needed for setup B (standalone script)
const NOTIFY_EMAIL = 'vivekindia136@gmail.com';
const SHEET_NAME = 'RSVPs';
const HEADERS = ['Submitted at', 'Name', 'Guests'];
const MAX_GUESTS = 20;
const MAX_NAME_LENGTH = 80;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const params = (e && e.parameter) || {};

    // Honeypot: real guests never see or fill the "website" field.
    if (params.website) return json_({ ok: true });

    const name = String(params.name || '').trim().replace(/\s+/g, ' ').slice(0, MAX_NAME_LENGTH);
    const guests = Number.parseInt(params.guests, 10);
    if (!name) return json_({ ok: false, error: 'name_required' });
    if (!(guests >= 1 && guests <= MAX_GUESTS)) return json_({ ok: false, error: 'invalid_guests' });

    const sheet = getSheet_();
    sheet.appendRow([new Date(), escapeFormula_(name), guests]);
    SpreadsheetApp.flush();

    emailWorkbook_(sheet, name, guests);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// Lets you open the /exec URL in a browser to confirm the deployment is live.
function doGet() {
  return json_({ ok: true, service: 'rsvp' });
}

function getSheet_() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd mmm yyyy, hh:mm');
  }
  return sheet;
}

// Stops names like "=HYPERLINK(...)" from being treated as spreadsheet formulas.
function escapeFormula_(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function emailWorkbook_(sheet, name, guests) {
  const ss = sheet.getParent();
  const rows = Math.max(sheet.getLastRow() - 1, 0);
  const totalGuests = rows
    ? sheet.getRange(2, 3, rows, 1).getValues().reduce((sum, [n]) => sum + (Number(n) || 0), 0)
    : 0;

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const xlsx = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName('Srishti-Vivek-RSVPs.xlsx');

  // DriveApp is referenced so Apps Script grants the Drive scope the export URL needs.
  const sheetLink = DriveApp.getFileById(ss.getId()).getUrl();

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'New RSVP: ' + name + ' (' + guests + ') — ' + rows + ' RSVPs, ' + totalGuests + ' guests',
    htmlBody:
      '<p><b>' + htmlEscape_(name) + '</b> RSVP\'d for <b>' + guests + '</b> ' + (guests === 1 ? 'guest' : 'guests') + '.</p>' +
      '<p>Total so far: <b>' + rows + '</b> RSVPs, <b>' + totalGuests + '</b> guests.</p>' +
      '<p>The full list is attached as an Excel sheet. Live sheet: <a href="' + sheetLink + '">open in Google Sheets</a>.</p>',
    attachments: [xlsx]
  });
}

function htmlEscape_(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
