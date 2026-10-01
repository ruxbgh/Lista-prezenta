/**
 * Prezență la curs — backend Google Apps Script
 *
 * Serverul NU primește și NU stochează nume sau ID-uri de studenți.
 * Stochează doar:
 *   - foaia "Coduri":   curs | cod personal anonim (6 caractere)
 *   - foaia "Prezenta": curs | cod | săptămâna | prezent (1/0) | data
 * Legătura cod ↔ nume există doar pe laptopul profesorului (pagina admin).
 */

var CODE_RE = /^[A-Z0-9]{6}$/;
var COURSE_RE = /^[a-z0-9-]{1,40}$/;
var MAX_WEEKS = 20;

function doGet() {
  return out_({ ok: true, service: 'prezenta' });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad_request' });
  }
  var props = PropertiesService.getScriptProperties();
  var action = body.action;

  try {
    // Acțiuni publice (studenți)
    if (action === 'status') return out_(status_(props, body));
    if (action === 'checkin') return out_(checkin_(props, body));

    // Prima conectare a profesorului: fixează cheia de admin
    if (action === 'init') return out_(init_(props, body));

    // Toate celelalte cer cheia de admin
    var token = props.getProperty('ADMIN_TOKEN');
    if (!token || body.token !== token) return out_({ ok: false, error: 'unauthorized' });

    if (action === 'verify') return out_({ ok: true });
    if (action === 'set_codes') return out_(setCodes_(body));
    if (action === 'start') return out_(start_(props, body));
    if (action === 'stop') return out_(stop_(props, body));
    if (action === 'session') {
      var cur = session_(props, course_(body));
      return out_(cur ? { ok: true, active: true, week: cur.week, code: cur.code, exp: cur.exp } : { ok: true, active: false });
    }
    if (action === 'attendance') return out_(attendance_(body));
    if (action === 'set_mark') return out_(setMark_(body));

    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: 'server_error', detail: String(err) });
  }
}

/* ---------- utilitare ---------- */

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
  }
  return sh;
}

// Coloana cu coduri e forțată ca text, ca Sheets să nu transforme un cod ca 23E456 în număr.
function codesSheet_() {
  var sh = sheet_('Coduri', ['curs', 'cod']);
  sh.getRange('B:B').setNumberFormat('@');
  return sh;
}
function marksSheet_() {
  var sh = sheet_('Prezenta', ['curs', 'cod', 'saptamana', 'prezent', 'data']);
  sh.getRange('B:B').setNumberFormat('@');
  return sh;
}

function course_(b) {
  var c = String(b.course || '');
  if (!COURSE_RE.test(c)) throw new Error('curs invalid');
  return c;
}

function week_(w) {
  var n = parseInt(w, 10);
  if (!(n >= 1 && n <= MAX_WEEKS)) throw new Error('saptamana invalida');
  return n;
}

function normCode_(c) {
  return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function session_(props, course) {
  var raw = props.getProperty('S_' + course);
  if (!raw) return null;
  var s = JSON.parse(raw);
  if (s.exp < Date.now()) return null;
  return s;
}

function validCodes_(course) {
  var data = codesSheet_().getDataRange().getValues();
  var set = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === course) set[String(data[i][1])] = true;
  }
  return set;
}

// ultima stare pentru fiecare (cod, săptămână) dintr-un curs
function latestMarks_(course) {
  var data = marksSheet_().getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== course) continue;
    map[data[i][1] + '|' + data[i][2]] = Number(data[i][3]) === 1 ? 1 : 0;
  }
  return map;
}

/* ---------- acțiuni ---------- */

function init_(props, b) {
  var t = String(b.token || '');
  if (!/^[a-f0-9]{64}$/.test(t)) return { ok: false, error: 'bad_token' };
  var existing = props.getProperty('ADMIN_TOKEN');
  if (!existing) {
    props.setProperty('ADMIN_TOKEN', t);
    return { ok: true, initialized: true };
  }
  return existing === t ? { ok: true } : { ok: false, error: 'unauthorized' };
}

function status_(props, b) {
  var course = course_(b);
  var s = session_(props, course);
  if (!s) return { ok: true, active: false };
  return { ok: true, active: true, week: s.week, exp: s.exp }; // codul sesiunii NU se trimite
}

function checkin_(props, b) {
  var course = course_(b);
  var code = normCode_(b.code);
  if (!CODE_RE.test(code)) return { ok: false, error: 'bad_code' };

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var s = session_(props, course);
    if (!s) return { ok: false, error: 'no_session' };
    if (String(b.sessionCode || '').trim() !== s.code) return { ok: false, error: 'wrong_session_code' };

    if (!validCodes_(course)[code]) return { ok: false, error: 'unknown_code' };

    var marks = latestMarks_(course);
    var already = marks[code + '|' + s.week] === 1;
    if (!already) {
      marksSheet_().appendRow([course, code, s.week, 1, new Date()]);
      marks[code + '|' + s.week] = 1;
    }
    var count = 0;
    for (var k in marks) {
      if (k.indexOf(code + '|') === 0 && marks[k] === 1) count++;
    }
    return { ok: true, already: already, week: s.week, count: count };
  } finally {
    lock.releaseLock();
  }
}

function setCodes_(b) {
  var course = course_(b);
  var codes = (b.codes || []).map(normCode_).filter(function (c) { return CODE_RE.test(c); });
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = codesSheet_();
    var data = sh.getDataRange().getValues();
    var keep = data.slice(1).filter(function (r) { return r[0] !== course; });
    codes.forEach(function (c) { keep.push([course, c]); });
    sh.getRange(2, 1, Math.max(sh.getLastRow(), 2), 2).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 2).setValues(keep);
    return { ok: true, count: codes.length };
  } finally {
    lock.releaseLock();
  }
}

function start_(props, b) {
  var course = course_(b);
  var week = week_(b.week);
  var minutes = Math.min(Math.max(parseInt(b.minutes, 10) || 8, 1), 60);
  var s = {
    week: week,
    code: String(Math.floor(1000 + Math.random() * 9000)),
    exp: Date.now() + minutes * 60 * 1000
  };
  props.setProperty('S_' + course, JSON.stringify(s));
  return { ok: true, week: s.week, code: s.code, exp: s.exp };
}

function stop_(props, b) {
  props.deleteProperty('S_' + course_(b));
  return { ok: true };
}

function attendance_(b) {
  var course = course_(b);
  var marks = latestMarks_(course);
  var list = [];
  for (var k in marks) {
    if (marks[k] !== 1) continue;
    var parts = k.split('|');
    list.push({ code: parts[0], week: Number(parts[1]) });
  }
  return { ok: true, marks: list };
}

function setMark_(b) {
  var course = course_(b);
  var code = normCode_(b.code);
  if (!CODE_RE.test(code)) return { ok: false, error: 'bad_code' };
  var week = week_(b.week);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    marksSheet_().appendRow([course, code, week, b.present ? 1 : 0, new Date()]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
