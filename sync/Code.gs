// Shared to-do list backend for the wedding planner site.
// Lives in a Google Sheet so Peter and Nicole see the same list on every device.
//
// Setup (one time, ~5 minutes):
//   1. Create a new Google Sheet (any name, e.g. "Wedding to-dos").
//   2. Extensions -> Apps Script. Delete the sample code, paste this whole file, save.
//   3. Change PASSCODE below to something short you'll both remember.
//   4. Deploy -> New deployment -> type "Web app".
//        Execute as: Me.   Who has access: Anyone.
//      Authorize when asked, then copy the Web app URL and send it to Red.
//   The sheet gets a "todos" tab automatically on first use.

var PASSCODE = 'change-me';
var SHEET = 'todos';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['id', 'text', 'done', 'added_by', 'added_at', 'updated_at']);
  }
  return sh;
}

function rows_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({ row: i + 1, id: String(r[0]), text: String(r[1]), done: r[2] === true || r[2] === 'TRUE',
               added_by: String(r[3] || ''), added_at: r[4], updated_at: r[5] });
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// GET -> the whole list
function doGet() {
  return json_({ ok: true, items: rows_().map(function (r) {
    return { id: r.id, text: r.text, done: r.done, added_by: r.added_by };
  }) });
}

// POST {code, action:'add', id, text, by, done} | {code, action:'toggle', id, done} | {code, action:'remove', id}
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (body.code !== PASSCODE) return json_({ ok: false, error: 'bad code' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    var now = new Date();
    if (body.action === 'add') {
      var text = String(body.text || '').trim();
      if (!text) return json_({ ok: false, error: 'empty' });
      var id = String(body.id || ('t' + now.getTime()));
      sh.appendRow([id, text, !!body.done, String(body.by || ''), now, now]);
      return json_({ ok: true, id: id });
    }
    var hit = rows_().filter(function (r) { return r.id === String(body.id); })[0];
    if (!hit) return json_({ ok: false, error: 'not found' });
    if (body.action === 'toggle') {
      sh.getRange(hit.row, 3).setValue(!!body.done);
      sh.getRange(hit.row, 6).setValue(now);
      return json_({ ok: true });
    }
    if (body.action === 'remove') {
      sh.deleteRow(hit.row);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}
