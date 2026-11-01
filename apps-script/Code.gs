/**
 * @OnlyCurrentDoc
 * Limits the script to the Sheet it's attached to, instead of every Sheet in
 * the account, so Google treats it as a lower-risk permission when authorizing.
 */

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
const HEADERS = ['submitted', 'parent_name', 'phone', 'mailing_address', 'children', 'number_of_children', 'attending', 'notes'];
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

// "Niki or Jeff K-M (game time decision)" -> { name: "Niki or Jeff K-M", notes: ["game time decision"] }
function splitNotes(text) {
  const notes = [];
  const name = String(text || '')
    .replace(/\(([^)]*)\)/g, (_, note) => { if (note.trim()) notes.push(note.trim()); return ' '; })
    .replace(/\s+/g, ' ')
    .trim();
  return { name: name, notes: notes };
}

// Name-only parent and children, with any notes typed into the name fields pulled out.
function cleanRsvp(parentName, kids, existingNotes) {
  const parent = splitNotes(parentName);
  const notes = existingNotes ? [String(existingNotes).trim()] : [];
  notes.push(...parent.notes);
  const cleanKids = kids.map(k => {
    const kid = splitNotes(k.name);
    kid.notes.forEach(n => notes.push(kid.name + ': ' + n));
    return { name: kid.name, age: k.age };
  });
  return { parent: parent.name, kids: cleanKids, notes: [...new Set(notes)].join('; ') };
}

// Receives the RSVP form's JSON (sent as text/plain to avoid a CORS preflight).
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const clean = cleanRsvp(d.parent_name, Array.isArray(d.kids) ? d.kids : []);
    getSheet().appendRow([
      new Date().toISOString(),
      clean.parent,
      String(d.phone || ''),
      String(d.mailing_address || NO_ADDRESS),
      JSON.stringify(clean.kids),
      Number(d.number_of_children) || 0,
      d.attending === 'Yes' ? 'Yes' : 'No',
      clean.notes
    ]);
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Public feed: full names, notes, address yes/no. No phone, no street address, no ages.
// Rows are cleaned again here so ones typed into the Sheet by hand show up the same way.
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
    const clean = cleanRsvp(r[1], kids, r[7]);
    return {
      submitted: new Date(r[0]).toISOString(),
      parent: clean.parent,
      children: clean.kids.map(k => k.name),
      childCount: Number(r[5]) || kids.length,
      attending: r[6] === 'Yes',
      hasAddress: !!r[3] && r[3] !== NO_ADDRESS,
      notes: clean.notes
    };
  });

  return json({ generatedAt: new Date().toISOString(), rsvps: rsvps });
}

function parseKids(cell) {
  try { return JSON.parse(cell) || []; } catch (_) { return []; }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
