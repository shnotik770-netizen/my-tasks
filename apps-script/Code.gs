/**
 * משימות שלי — צד שרת (Google Apps Script)
 * ------------------------------------------------------------
 * הקובץ הזה מודבק בעורך ה-Apps Script של הגיליון "משימות שלי - נתונים".
 * קוד העריכה לא נשמר כאן (הקובץ ציבורי ב-GitHub) אלא ב:
 *   הגדרות הפרויקט ← מאפייני סקריפט ← EDIT_CODE
 * הוראות מלאות: README.md במאגר.
 */

// מבנה הגיליונות: שם ← עמודות. העמודה הראשונה היא המפתח.
var HEADERS = {
  Categories: ['id', 'name', 'parentId', 'color', 'order', 'active'],
  Tasks: ['id', 'title', 'categoryId', 'dueDate', 'recurrence', 'status', 'continuous', 'parentTaskId',
          'rabbiType', 'rabbiStatus', 'needsDecision', 'notes', 'createdAt', 'updatedAt', 'doneDate'],
  Log: ['id', 'taskId', 'date', 'result', 'note', 'createdAt'],
  Priorities: ['taskId', 'order', 'note'],
  RabbiNotes: ['id', 'taskId', 'author', 'date', 'action', 'text'],
  WorkHours: ['id', 'date', 'start', 'end', 'note'],
  DaysOff: ['date', 'reason'],
  Settings: ['key', 'value']
};
var EDITABLE = ['Categories', 'Tasks', 'Log', 'RabbiNotes', 'WorkHours', 'DaysOff'];

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ status: 'error', message: 'בקשה לא תקינה' }); }
  return handle_(body);
}

function handle_(p) {
  try {
    return out_(route_(p), p.callback);
  } catch (err) {
    return out_({ status: 'error', message: String((err && err.message) || err) }, p.callback);
  }
}

function out_(obj, callback) {
  var s = JSON.stringify(obj);
  if (callback && /^[\w.$]+$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  switch (p.action) {
    case 'ping':
      return { status: 'success' };
    case 'checkCode':
      return { status: 'success', ok: isEditor_(p.code) };
    case 'load':
      return { status: 'success', data: readAll_(), editor: p.code ? isEditor_(p.code) : false };
    case 'rabbiAction':
      return withLock_(function () { return rabbiAction_(p); });
    case 'batch':
      if (!isEditor_(p.code)) throw new Error('קוד עריכה שגוי — אין הרשאה לשנות');
      return withLock_(function () { return batch_(p.ops || []); });
    default:
      throw new Error('פעולה לא מוכרת: ' + p.action);
  }
}

function isEditor_(code) {
  var real = PropertiesService.getScriptProperties().getProperty('EDIT_CODE');
  if (!real) throw new Error('לא הוגדר EDIT_CODE במאפייני הסקריפט');
  return String(code || '') === String(real);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- פעולות עריכה (רק עם קוד) ----------
function batch_(ops) {
  ops.forEach(function (op) {
    if (op.op === 'setPriorities') return setPriorities_(op.list || []);
    if (EDITABLE.indexOf(op.sheet) < 0) throw new Error('גיליון לא מורשה: ' + op.sheet);
    if (op.op === 'upsert') upsert_(op.sheet, op.rec || {});
    else if (op.op === 'delete') deleteRow_(op.sheet, op.id);
    else throw new Error('פעולה לא מוכרת: ' + op.op);
  });
  return { status: 'success' };
}

// ---------- פעולות הרב סגל (פתוח לכולם) ----------
// act: approve | reject | seen | comment
function rabbiAction_(p) {
  var act = String(p.act || '');
  if (['approve', 'reject', 'seen', 'comment'].indexOf(act) < 0) throw new Error('פעולה לא מוכרת');
  var task = findById_('Tasks', p.taskId);
  if (!task) throw new Error('המשימה לא נמצאה');
  var now = new Date().toISOString();
  upsert_('RabbiNotes', {
    id: p.noteId || ('n' + Date.now()),
    taskId: task.id,
    author: String(p.author || '').slice(0, 60),
    date: String(p.date || ''),
    action: act,
    text: String(p.text || '').slice(0, 2000)
  });
  if (act !== 'comment') {
    var status = act === 'approve' ? 'approved' : act === 'reject' ? 'rejected' : 'seen';
    upsert_('Tasks', { id: task.id, rabbiStatus: status, updatedAt: now });
  }
  if (act === 'reject') {
    upsert_('Tasks', {
      id: p.followupId || ('t' + Date.now()),
      title: 'המשך: ' + task.title,
      categoryId: task.categoryId,
      dueDate: '',
      recurrence: 'none',
      status: 'open',
      continuous: false,
      parentTaskId: task.id,
      rabbiType: '',
      rabbiStatus: '',
      needsDecision: true,
      notes: p.text ? ('הרב סגל לא אישר: ' + p.text) : 'הרב סגל לא אישר',
      createdAt: now,
      updatedAt: now,
      doneDate: ''
    });
  }
  return { status: 'success' };
}

// ---------- עזרי גיליון ----------
function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function headers_(s) {
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String);
}

function norm_(v, tz) {
  if (v instanceof Date) {
    if (v.getFullYear() < 1901) return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  if (v === 'TRUE') return true;
  if (v === 'FALSE') return false;
  return v;
}

function readSheet_(name) {
  var s = sheet_(name);
  var lr = s.getLastRow();
  if (lr < 2) return [];
  var h = headers_(s);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return s.getRange(2, 1, lr - 1, h.length).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      var o = {};
      h.forEach(function (k, i) { o[k] = norm_(r[i], tz); });
      return o;
    });
}

function readAll_() {
  var d = {};
  Object.keys(HEADERS).forEach(function (n) { if (n !== 'Settings') d[n] = readSheet_(n); });
  return d;
}

function findRow_(s, id) {
  var lr = s.getLastRow();
  if (lr < 2) return -1;
  var keys = s.getRange(2, 1, lr - 1, 1).getDisplayValues();
  for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === String(id)) return i + 2;
  return -1;
}

function findById_(name, id) {
  var list = readSheet_(name);
  for (var i = 0; i < list.length; i++) if (String(list[i][HEADERS[name][0]]) === String(id)) return list[i];
  return null;
}

function cell_(v) {
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  if (v === null || v === undefined) return '';
  return String(v);
}

function upsert_(name, rec) {
  var s = sheet_(name);
  var h = headers_(s);
  var key = rec[h[0]];
  if (key === undefined || key === '') throw new Error('חסר מפתח ל-' + name);
  var row = findRow_(s, key);
  var existing = row > 0 ? s.getRange(row, 1, 1, h.length).getValues()[0] : h.map(function () { return ''; });
  var values = h.map(function (k, i) { return rec.hasOwnProperty(k) ? cell_(rec[k]) : existing[i]; });
  if (row < 0) row = s.getLastRow() + 1;
  var rng = s.getRange(row, 1, 1, h.length);
  rng.setNumberFormat('@');
  rng.setValues([values]);
}

function deleteRow_(name, id) {
  var s = sheet_(name);
  var row = findRow_(s, id);
  if (row > 0) s.deleteRow(row);
}

function setPriorities_(list) {
  var s = sheet_('Priorities');
  var lr = s.getLastRow();
  if (lr > 1) s.getRange(2, 1, lr - 1, s.getLastColumn()).clearContent();
  if (!list.length) return;
  var rows = list.map(function (p, i) { return [cell_(p.taskId), String(i + 1), cell_(p.note)]; });
  var rng = s.getRange(2, 1, rows.length, 3);
  rng.setNumberFormat('@');
  rng.setValues(rows);
}

/** להרצה ידנית פעם אחת (לא חובה): יוצר גיליונות חסרים */
function setup() {
  Object.keys(HEADERS).forEach(function (n) { sheet_(n); });
}
