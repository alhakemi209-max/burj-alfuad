/**
 * برج الفؤاد الطبي — مطابقة محفظة جيب مع حسابات المحصلين
 * إعداد: أصيل أمين
 */

var TABLES = {
  collectors: { title: 'المحصلون', fields: [
    ['id','المعرف'],['name','الاسم'],['codes','أرقام النقاط'],['aliases','أسماء أخرى'],['ts','وقت الإضافة']] },
  stmt: { title: 'كشف جيب', fields: [
    ['id','الرقم المرجعي'],['date','التاريخ'],['time','الوقت'],['cur','العملة'],['amount','المبلغ'],['kind','النوع'],
    ['code','رقم النقطة'],['pointName','اسم النقطة'],['opNo','رقم العملية'],['payer','الدافع'],['op','نوع العملية'],
    ['desc','البيان'],['bal','الرصيد'],['file','الملف'],['coll','المحصل']] },
  vouchers: { title: 'سندات السيستم', fields: [
    ['id','المعرف'],['date','التاريخ'],['time','الوقت'],['coll','المحصل'],['cur','العملة'],['method','طريقة الدفع'],
    ['amount','المبلغ'],['no','رقم السند'],['note','البيان'],['ts','وقت الإدخال']] },
  entries: { title: 'إجماليات المحصلين', fields: [
    ['id','المعرف'],['date','التاريخ'],['coll','المحصل'],['cur','العملة'],['cash','نقدي'],['jaib','جيب'],['note','ملاحظة'],['ts','وقت الإدخال']] }
};
var NUMERIC = { amount:1, bal:1, ts:1, cash:1, jaib:1 };
var LISTS = { codes:1, aliases:1 };
var USERS_SHEET = 'المستخدمون', SETTINGS_SHEET = 'الإعدادات';
var SESSION_SECONDS = 21600;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('برج الفؤاد الطبي')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// ---------- sheets ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dbe7fb');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  }
  return sh;
}

function formats_(t, count) {
  var row = t.fields.map(function (f) { return NUMERIC[f[0]] ? '#,##0.00' : '@'; }), out = [];
  for (var i = 0; i < count; i++) out.push(row);
  return out;
}

function table_(col) {
  var t = TABLES[col];
  if (!t) throw new Error('جدول غير معروف');
  return { t: t, sh: sheet_(t.title, t.fields.map(function (f) { return f[1]; })) };
}

function cellOut_(key, v) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, key === 'time' ? 'HH:mm' : 'yyyy-MM-dd');
  }
  if (LISTS[key]) return String(v || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (NUMERIC[key]) return v === '' || v === null ? 0 : Number(v) || 0;
  return v === null || v === undefined ? '' : String(v);
}

function cellIn_(key, v) {
  if (LISTS[key]) return (v || []).join(', ');
  if (NUMERIC[key]) return v === '' || v === null || v === undefined ? '' : (Number(v) || 0);
  return v === null || v === undefined ? '' : String(v);
}

function readTable_(col) {
  var x = table_(col), n = x.t.fields.length, last = x.sh.getLastRow();
  if (last < 2) return [];
  var vals = x.sh.getRange(2, 1, last - 1, n).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === '' || vals[i][0] === null) continue;
    var o = {};
    for (var j = 0; j < n; j++) o[x.t.fields[j][0]] = cellOut_(x.t.fields[j][0], vals[i][j]);
    out.push(o);
  }
  return out;
}

function writeRows_(col, rows) {
  var x = table_(col), n = x.t.fields.length, last = x.sh.getLastRow();
  var ids = last >= 2 ? x.sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var at = {};
  for (var i = 0; i < ids.length; i++) at[String(ids[i][0])] = i + 2;
  var fresh = [];
  rows.forEach(function (r) {
    if (!r || !r.id) return;
    var line = x.t.fields.map(function (f) { return cellIn_(f[0], r[f[0]]); });
    var row = at[String(r.id)];
    if (row) {
      x.sh.getRange(row, 1, 1, n).setNumberFormats(formats_(x.t, 1)).setValues([line]);
    } else {
      at[String(r.id)] = -1;
      fresh.push(line);
    }
  });
  if (fresh.length) {
    var start = x.sh.getLastRow() + 1;
    x.sh.getRange(start, 1, fresh.length, n).setNumberFormats(formats_(x.t, fresh.length)).setValues(fresh);
  }
}

function deleteRow_(col, id) {
  var x = table_(col), last = x.sh.getLastRow();
  if (last < 2) return;
  var ids = x.sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) x.sh.deleteRow(i + 2);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- users & sessions ----------
function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function hash_(pw, salt) {
  var x = salt + ':' + pw;
  for (var i = 0; i < 300; i++) x = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, x + salt, Utilities.Charset.UTF_8));
  return x;
}

function usersSheet_() {
  var sh = sheet_(USERS_SHEET, ['اسم المستخدم', 'الملح', 'البصمة', 'الصلاحية']);
  return sh;
}

function users_() {
  var sh = usersSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { u: String(r[0]), salt: String(r[1]), hash: String(r[2]), role: String(r[3]) || 'user' }; });
}

function findUser_(name) {
  var n = String(name || '').trim().toLowerCase();
  return users_().filter(function (u) { return u.u.toLowerCase() === n; })[0] || null;
}

function addUserRow_(name, pw, role) {
  var salt = Utilities.getUuid().replace(/-/g, '');
  usersSheet_().appendRow([String(name).trim(), salt, hash_(pw, salt), role]);
}

function newSession_(u) {
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + token, JSON.stringify({ u: u.u, role: u.role }), SESSION_SECONDS);
  return { token: token, user: { u: u.u, role: u.role } };
}

function session_(token) {
  var raw = token && CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  var s = JSON.parse(raw);
  if (!findUser_(s.u)) throw new Error('SESSION_EXPIRED');
  return s;
}

function admin_(token) {
  var s = session_(token);
  if (s.role !== 'admin') throw new Error('هذه العملية للمدير فقط');
  return s;
}

function authStatus() {
  return { setup: users_().length === 0 };
}

function setupAdmin(name, pw) {
  return withLock_(function () {
    if (users_().length) throw new Error('تم إنشاء حساب المدير مسبقاً');
    name = String(name || '').trim();
    if (name.length < 2) throw new Error('اسم المستخدم قصير');
    if (String(pw || '').length < 4) throw new Error('الرمز 4 خانات على الأقل');
    addUserRow_(name, pw, 'admin');
    return newSession_(findUser_(name));
  });
}

function login(name, pw) {
  var cache = CacheService.getScriptCache(), key = 'f_' + String(name || '').trim().toLowerCase();
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) throw new Error('محاولات كثيرة خاطئة، حاول بعد 10 دقائق');
  var u = findUser_(name);
  if (!u || hash_(pw, u.salt) !== u.hash) {
    cache.put(key, String(fails + 1), 600);
    throw new Error('اسم المستخدم أو الرمز غير صحيح');
  }
  cache.remove(key);
  return newSession_(u);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return true;
}

function listUsers(token) {
  admin_(token);
  return users_().map(function (u) { return { u: u.u, role: u.role }; });
}

function addUser(token, name, pw, role) {
  admin_(token);
  return withLock_(function () {
    name = String(name || '').trim();
    if (name.length < 2 || String(pw || '').length < 4) throw new Error('الاسم حرفان والرمز 4 خانات على الأقل');
    if (findUser_(name)) throw new Error('اسم المستخدم موجود');
    addUserRow_(name, pw, role === 'admin' ? 'admin' : 'user');
    return listUsers(token);
  });
}

function deleteUser(token, name) {
  var s = admin_(token);
  return withLock_(function () {
    if (String(name).toLowerCase() === s.u.toLowerCase()) throw new Error('لا يمكن حذف حسابك');
    var sh = usersSheet_(), last = sh.getLastRow();
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).toLowerCase() === String(name).toLowerCase()) sh.deleteRow(i + 2);
    }
    return listUsers(token);
  });
}

// ---------- data ----------
function settings_() {
  var sh = sheet_(SETTINGS_SHEET, ['الإعداد', 'القيمة']), last = sh.getLastRow(), o = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { if (r[0] !== '') o[String(r[0])] = Number(r[1]); });
  return o;
}

function loadAll(token) {
  session_(token);
  return {
    collectors: readTable_('collectors'),
    stmt: readTable_('stmt'),
    vouchers: readTable_('vouchers'),
    entries: readTable_('entries'),
    settings: settings_()
  };
}

function putRows(token, col, rows) {
  session_(token);
  if (!Array.isArray(rows)) throw new Error('بيانات غير صالحة');
  return withLock_(function () { writeRows_(col, rows); return rows.length; });
}

function deleteRow(token, col, id) {
  session_(token);
  return withLock_(function () { deleteRow_(col, id); return true; });
}

function saveSettings(token, obj) {
  session_(token);
  return withLock_(function () {
    var sh = sheet_(SETTINGS_SHEET, ['الإعداد', 'القيمة']), last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
    var rows = Object.keys(obj || {}).map(function (k) { return [k, String(Number(obj[k]) || 0)]; });
    if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
    return true;
  });
}

// ---------- reading collector files (image, PDF, Excel, Word) ----------
function extractFile(token, name, mime, b64) {
  session_(token);
  if (typeof Drive === 'undefined') throw new Error('DRIVE_DISABLED');
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name);
  var isSheet = /sheet|excel|csv|spreadsheet/i.test(mime || '') || /\.(xlsx|xls|csv|ods)$/i.test(name || '');
  var target = isSheet ? MimeType.GOOGLE_SHEETS : MimeType.GOOGLE_DOCS;
  var file;
  if (Drive.Files.create) {
    file = Drive.Files.create({ name: 'burj_tmp_' + name, mimeType: target }, blob, { ocrLanguage: 'ar' });
  } else {
    file = Drive.Files.insert({ title: 'burj_tmp_' + name, mimeType: target }, blob, { convert: true, ocr: !isSheet, ocrLanguage: 'ar' });
  }
  try {
    var tables = [], text = '';
    if (isSheet) {
      SpreadsheetApp.openById(file.id).getSheets().forEach(function (sh) {
        if (sh.getLastRow() > 0) tables.push(sh.getDataRange().getDisplayValues());
      });
    } else {
      var body = DocumentApp.openById(file.id).getBody();
      body.getTables().forEach(function (t) {
        var rows = [];
        for (var r = 0; r < t.getNumRows(); r++) {
          var row = t.getRow(r), cells = [];
          for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText());
          rows.push(cells);
        }
        tables.push(rows);
      });
      text = body.getText();
    }
    return { tables: tables, text: text };
  } finally {
    try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e) {}
  }
}
