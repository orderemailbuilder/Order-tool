/**
 * Order Email Builder – Google Apps Script backend (replaces the old script)
 *
 * What it does
 *  - GET  ?pin=1234                      -> login + returns the rows this PIN may see
 *  - GET  ?pin=1234&action=delete&ts=... -> deletes one row (own row, last 7 days; admin: any row, last 7 days)
 *  - POST with edit=1 and ets=<row ts>    -> corrects one row in place (own row, last 7 days; admin: any row, last 7 days);
 *                                           ts, saved and pin stay as they were and the 'edited' column gets the word edited
 *  - POST (JSON body from the page)      -> saves a new row, or updates the earlier row of the same patient
 *                                           (same patient + same PIN within 24 hours, or the row chosen with "Update earlier order")
 *
 * Setup
 *  1. Open the Google Sheet -> Extensions -> Apps Script. Keep a copy of the OLD code somewhere first.
 *  2. Replace everything with this file. Set ADMIN_PIN below (the PIN that sees every staff member's rows).
 *  3. Deploy -> Manage deployments -> pencil icon -> Version: "New version" -> Deploy.
 *     (Editing the existing deployment keeps the same /exec link, so the page needs no change.)
 *     Execute as: Me.  Who has access: Anyone.
 */

var ADMIN_PIN = '0111';      // admin PIN (sees every staff member's rows). Keep the quotes so the leading 0 is not lost. Empty = nobody is admin.
var SHEET_NAME = '';         // '' = read EVERY tab that has the order columns (e.g. one tab per month) and save into the first tab;
                             // or type one tab name to use only that tab
var MONTHLY_TABS = false;    // true = new emails are saved in a tab named after the current month (created automatically);
                             // History still reads all tabs
var MONTH_FORMAT = 'MMMM yyyy';   // tab name used when MONTHLY_TABS is true, e.g. 'October 2026' (make it match your tab names)
var ALLOWED_PINS = [];       // [] = any 4-digit PIN may log in (new PIN = new staff). Or list them: ['0111','0222']
var VERSION = '2026-10-11-tabs5';
var DAY = 86400000;
var TABS_ = [];             // which tabs were read / skipped (sent to the page for its info line)

// canonical field -> accepted header names (compared ignoring case, spaces and symbols)
var FIELDS = [
  'ts', 'saved', 'date', 'time', 'status', 'name', 'cnic', 'pmobile', 'product', 'rate', 'qty',
  'doctor', 'address', 'city', 'notes', 'invoice', 'icekit', 'chg', 'pin', 'speciality', 'disease', 'edited'
];
var ALIASES = {
  ts: ['ts', 'timestamp', 'cts'],
  saved: ['saved', 'savedat', 'savedon'],
  date: ['date', 'orderdate'],
  time: ['time', 'ordertime'],
  status: ['status', 'patientstatus'],
  name: ['name', 'patientname', 'patient', 'customername', 'customer'],
  cnic: ['cnic', 'patientcnic'],
  pmobile: ['pmobile', 'mobile', 'patientmobile', 'phone', 'contact', 'mobileno', 'phoneno', 'contactno'],
  product: ['product', 'productname', 'medicine', 'medicinename', 'drug', 'item'],
  rate: ['rate', 'patientprice', 'price'],
  qty: ['qty', 'quantity'],
  doctor: ['doctor', 'prescriber', 'prescribername', 'doctorname'],
  address: ['address', 'deliveryaddress'],
  city: ['city'],
  notes: ['notes', 'additionalnotes', 'note'],
  invoice: ['invoice'],
  icekit: ['icekit'],
  chg: ['chg', 'change'],
  pin: ['pin', 'staffpin'],
  speciality: ['speciality', 'specialty', 'prescriberspeciality', 'doctorspeciality'],
  disease: ['disease', 'diagnosis'],
  edited: ['edited', 'editedat']
};
var TEXT_FIELDS = ['cnic', 'pmobile', 'date', 'time', 'saved', 'rate', 'qty', 'pin', 'notes', 'speciality', 'disease', 'edited'];

/* ---------------- entry points ---------------- */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var who = auth_(p.pin);
    if (who.error) return json_({ error: who.error });
    if (p.action === 'delete') return deleteRow_(who, p);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var t = readTable_();
      var rows = t.rows.filter(function (r) { return canSee_(who, r); });
      return json_({ ok: true, role: who.admin ? 'admin' : 'user', v: VERSION, tabs: TABS_, rows: rows.map(publicRow_) });
    } finally { lock.releaseLock(); }
  } catch (err) {
    return json_({ error: 'Script error: ' + err });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var rec = JSON.parse(e.postData.contents);
    var who = auth_(rec.pin);
    if (who.error) return json_({ error: who.error });
    lock.waitLock(20000);
    var res = saveRecord_(who, rec);
    if (res && res.error) return json_({ error: res.error });
    return json_({ ok: true });
  } catch (err) {
    return json_({ error: 'Script error: ' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------------- auth ---------------- */

function pn_(x) {
  var d = String(x == null ? '' : x).replace(/\D/g, '');
  return d ? ('0000' + d).slice(-Math.max(4, d.length)) : '';
}

function auth_(pin) {
  var p = String(pin == null ? '' : pin).replace(/\D/g, '');
  if (p.length !== 4) return { error: 'pin' };
  var norm = pn_(p);
  var admin = !!ADMIN_PIN && norm === pn_(ADMIN_PIN);
  if (!admin && ALLOWED_PINS.length) {
    var ok = ALLOWED_PINS.some(function (a) { return pn_(a) === norm; });
    if (!ok) {   // wrong PINs are counted; a listed PIN is never blocked by other people's wrong guesses
      var cache = CacheService.getScriptCache();
      var fails = (+cache.get('fails') || 0) + 1;
      cache.put('fails', String(fails), 900);
      return { error: fails > 30 ? 'locked' : 'pin' };
    }
  }
  return { pin: norm, admin: admin };
}

function canSee_(who, row) {
  return who.admin || pn_(row.pin) === who.pin;
}

/* ---------------- sheet access ---------------- */

// every tab that looks like an order tab (has name + product columns); with SHEET_NAME only that tab
function dataSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = [];
  TABS_ = [];
  ss.getSheets().forEach(function (sh) {
    var nm = sh.getName();
    if (SHEET_NAME && !MONTHLY_TABS && nm !== SHEET_NAME) { TABS_.push({ n: nm, r: 0, s: 'not the chosen tab' }); return; }
    if (sh.getLastRow() < 1) { TABS_.push({ n: nm, r: 0, s: 'empty' }); return; }
    var c = columns_(sh, true);
    var hits = ['name', 'cnic', 'pmobile', 'product'].filter(function (f) { return f in c.map; }).length;
    if (hits >= 2) out.push({ sh: sh, c: c }); else TABS_.push({ n: nm, r: 0, s: 'no order columns' });
  });
  return out;
}

// the tab where a NEW email is appended
function writeSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh;
  if (MONTHLY_TABS) {
    var nm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), MONTH_FORMAT);
    sh = ss.getSheetByName(nm);
    if (!sh) {
      sh = ss.insertSheet(nm);
      var d = dataSheets_();
      var head = d.length ? d[0].c.head : FIELDS;
      sh.getRange(1, 1, 1, head.length).setValues([head]);
    }
    return sh;
  }
  sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sh) throw new Error('Sheet tab not found: ' + SHEET_NAME);
  return sh;
}

function norm_(h) { return String(h == null ? '' : h).toLowerCase().replace(/[^a-z0-9]/g, ''); }

// header row -> { field: columnIndex(0-based) }; when writing, missing canonical columns are added at the end
function columns_(sh, readOnly) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getLastRow() ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var hasHeader = head.some(function (h) { return String(h).trim() !== ''; });
  if (!hasHeader) {
    if (readOnly) return { map: {}, width: 0, head: [] };
    sh.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS]);
    head = FIELDS.slice();
  }
  var map = {}, extra = [];
  head.forEach(function (h, i) {
    var n = norm_(h);
    for (var f in ALIASES) {
      if (!(f in map) && ALIASES[f].indexOf(n) >= 0) { map[f] = i; return; }
    }
  });
  if (!readOnly) {
    FIELDS.forEach(function (f) { if (!(f in map)) extra.push(f); });
    if (extra.length) {
      var start = head.length;
      sh.getRange(1, start + 1, 1, extra.length).setValues([extra]);
      extra.forEach(function (f, i) { map[f] = start + i; });
      head = head.concat(extra);
    }
  }
  return { map: map, width: head.length, head: head };
}

function str_(v, field) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return field === 'time' ? Utilities.formatDate(v, tz, 'h:mm a')
         : field === 'saved' ? Utilities.formatDate(v, tz, 'dd/MM/yyyy, HH:mm:ss')
         : Utilities.formatDate(v, tz, 'd/M/yyyy');
  }
  return v == null ? '' : String(v);
}

function parseSaved_(s) {   // "10/10/2026, 14:05:03" (en-GB)
  var m = String(s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0)).getTime() : 0;
}

function readTable_() {
  var rows = [];
  dataSheets_().forEach(function (d) {
    var sh = d.sh, c = d.c, last = sh.getLastRow(), cnt = 0;
    if (last < 2) { TABS_.push({ n: sh.getName(), r: 0, s: '' }); return; }
    var vals = sh.getRange(2, 1, last - 1, c.width).getValues();
    vals.forEach(function (v, i) {
      var r = { _row: i + 2, _sheet: sh.getName() };
      FIELDS.forEach(function (f) { r[f] = str_(v[c.map[f]], f); });
      var t = +r.ts;
      if (!t) t = parseSaved_(r.saved);
      r.ts = t || '';
      if (v[c.map.ts] instanceof Date) r.ts = v[c.map.ts].getTime();
      var empty = !r.name && !r.cnic && !r.pmobile && !r.product;
      if (!empty) { rows.push(r); cnt++; }
    });
    TABS_.push({ n: sh.getName(), r: cnt, s: '' });
  });
  return { rows: rows };
}

function publicRow_(r) {
  var o = {};
  FIELDS.forEach(function (f) { o[f] = r[f]; });
  o.ts = r.ts;
  o.cts = r.ts;
  return o;
}

/* ---------------- save / update ---------------- */

function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }
function pkey_(s) { return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim(); }
function last10_(s) { return digits_(s).slice(-10); }
function mobs_(s) {
  return String(s || '').split(/[\/,;|\n]+/).map(digits_).filter(function (x) { return x.length >= 10; }).map(last10_);
}
function mobOverlap_(a, b) {
  var B = mobs_(b);
  return mobs_(a).some(function (x) { return B.indexOf(x) >= 0; });
}
function sameWho_(r, rec) {
  var a = digits_(r.cnic), b = digits_(rec.cnic);
  if (a && b) return a === b;
  return pkey_(r.name) === pkey_(rec.name) && a === b && (!!b || mobOverlap_(r.pmobile, rec.pmobile));
}

function saveRecord_(who, rec) {
  var t = readTable_();
  var now = Date.now();
  var cts = +rec.cts || now;
  var target = null;

  if (String(rec.edit) === '1') return editRecord_(who, rec, t, now);

  if (String(rec.chg) === '1' && rec.rts) {
    t.rows.forEach(function (r) {
      if (String(r.ts) === String(rec.rts) && canSee_(who, r)) target = r;
    });
  }
  if (!target && String(rec.nw) !== '1') {
    t.rows.forEach(function (r) {
      var ts = +r.ts;
      if (ts && now - ts < DAY && canSee_(who, r) && sameWho_(r, rec)) {
        if (!target || ts >= +target.ts) target = r;
      }
    });
  }

  var values = {};
  FIELDS.forEach(function (f) { values[f] = rec[f] == null ? '' : String(rec[f]); });
  values.ts = String(cts);
  values.pin = who.pin;
  values.chg = String(rec.chg === true || String(rec.chg) === '1' ? '1' : '');
  values.invoice = String(rec.invoice || '');
  values.icekit = rec.icekit === true ? 'Yes' : rec.icekit === false ? 'No' : String(rec.icekit || '');

  var sh = target ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(target._sheet) : writeSheet_();
  var c = columns_(sh);
  var rowIdx = target ? target._row : sh.getLastRow() + 1;
  if (rowIdx < 2) rowIdx = 2;
  FIELDS.forEach(function (f) {
    var cell = sh.getRange(rowIdx, c.map[f] + 1);
    if (TEXT_FIELDS.indexOf(f) >= 0) cell.setNumberFormat('@');   // keeps 0300… and CNIC dashes as typed
    if (f === 'ts') cell.setNumberFormat('0');
    cell.setValue(f === 'ts' ? Number(values[f]) : values[f]);
  });
}

/* ---------------- edit (correct an existing row) ---------------- */

function editRecord_(who, rec, t, now) {
  var key = String(rec.ets || '');
  if (!key) return { error: 'Missing row id' };
  var hit = null;
  t.rows.forEach(function (r) { if (String(r.ts) === key) hit = r; });
  if (!hit) return { error: 'Row not found (it may have been deleted)' };
  if (!canSee_(who, hit)) return { error: 'You can only edit your own emails' };
  var ts = +hit.ts;
  if (!ts || now - ts >= 7 * DAY) return { error: 'Emails older than 7 days cannot be edited' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hit._sheet);
  var c = columns_(sh);
  var KEEP = ['ts', 'saved', 'pin', 'chg', 'edited'];   // these stay as they were
  FIELDS.forEach(function (f) {
    if (KEEP.indexOf(f) >= 0) return;
    var v = rec[f] == null ? '' : String(rec[f]);
    if (f === 'invoice') v = String(rec.invoice || '');
    if (f === 'icekit') v = rec.icekit === true ? 'Yes' : rec.icekit === false ? 'No' : String(rec.icekit || '');
    var cell = sh.getRange(hit._row, c.map[f] + 1);
    if (TEXT_FIELDS.indexOf(f) >= 0) cell.setNumberFormat('@');
    cell.setValue(v);
  });
  var ec = sh.getRange(hit._row, c.map.edited + 1);
  ec.setNumberFormat('@');
  ec.setValue('edited');
  return { ok: true };
}

/* ---------------- delete ---------------- */

function deleteRow_(who, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var key = String(p.ts || p.cts || '');
    if (!key) return json_({ error: 'Missing row id' });
    var t = readTable_();
    var hit = null;
    t.rows.forEach(function (r) { if (String(r.ts) === key) hit = r; });
    if (!hit) return json_({ error: 'Row not found (it may already be deleted)' });
    if (!canSee_(who, hit)) return json_({ error: 'You can only delete your own emails' });
    var ts = +hit.ts;
    if (!ts || Date.now() - ts >= 7 * DAY) return json_({ error: 'Emails older than 7 days cannot be deleted' });
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hit._sheet).deleteRow(hit._row);
    return json_({ ok: true, deleted: true });
  } finally { lock.releaseLock(); }
}

/* ---------------- helpers ---------------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
