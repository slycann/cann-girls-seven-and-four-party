/**
 * RSVP collector + sanitized feed for the party summary page.
 *
 * Setup: create a Google Sheet, then Extensions > Apps Script, paste this file,
 * and Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web app URL into SHEET_URL in index.html and FEED_URL in summary/index.html.
 *
 * The sheet keeps the full data (phone, address). doGet only ever returns the
 * sanitized fields below, because the web app URL is public.
 */

const SHEET_NAME = 'RSVPs';
const HEADERS = ['submitted', 'parent_name', 'phone', 'mailing_address', 'children', 'number_of_children', 'attending'];
const NO_ADDRESS = '(not given)';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Receives the RSVP form's JSON (sent as text/plain to avoid a CORS preflight).
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    getSheet().appendRow([
      new Date().toISOString(),
      String(d.parent_name || ''),
      String(d.phone || ''),
      String(d.mailing_address || NO_ADDRESS),
      JSON.stringify(Array.isArray(d.kids) ? d.kids : []),
      Number(d.number_of_children) || 0,
      d.attending === 'Yes' ? 'Yes' : 'No'
    ]);
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Public feed: abbreviated names, address yes/no, no phone, no street address, no ages.
function doGet() {
  const rows = getSheet().getDataRange().getValues().slice(1).filter(r => r[0] && r[1]);

  // A family that RSVPs again replaces its earlier answer (matched on phone, else name).
  const latest = new Map();
  rows.forEach(r => {
    const digits = String(r[2]).replace(/\D/g, '');
    const key = digits || String(r[1]).trim().toLowerCase();
    latest.set(key, r);
  });

  const rsvps = [...latest.values()].map(r => {
    const kids = parseKids(r[4]);
    return {
      submitted: new Date(r[0]).toISOString(),
      parent: abbreviate(r[1]),
      children: kids.map(k => firstName(k.name)),
      childCount: Number(r[5]) || kids.length,
      attending: r[6] === 'Yes',
      hasAddress: !!r[3] && r[3] !== NO_ADDRESS
    };
  });

  return json({ generatedAt: new Date().toISOString(), rsvps: rsvps });
}

function parseKids(cell) {
  try { return JSON.parse(cell) || []; } catch (_) { return []; }
}

function firstName(name) {
  return String(name || '').trim().split(/\s+/)[0] || '?';
}

function abbreviate(name) {
  const parts = String(name).trim().split(/\s+/);
  if (parts.length < 2) return parts[0];
  return parts[0] + ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.';
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
