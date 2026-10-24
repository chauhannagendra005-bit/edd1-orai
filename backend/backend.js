/* EDD-I Orai Field Entry backend - loaded at run time by the one-time Loader.gs in the Google Sheet.
   Do not paste this file into Apps Script; paste Loader.gs instead. Generated from Code.gs. */
/**
 * EDD-I Orai — Field Entry backend (Google Apps Script)
 * Paste this whole file into the Sheet:  Extensions ▸ Apps Script ▸ Code.gs
 * (Running 'setup' once is optional now - the sheets and the admin user are created automatically on first use.)
 * Then: Deploy ▸ New deployment ▸ Web app
 *       Execute as: Me   ·   Who has access: Anyone
 * Copy the Web-app URL (ends with /exec) into the dashboard's Field Entry login screen.
 * Roles: field (enters data), admin (also sees everyone's entries and manages users), developer (like admin, and the only one who can manage developer accounts).
 * Version 4: also saves Survey / Disconnection / Recheck / Phone Calling data in their own tabs (created automatically).
 * Version 5: also moves survey / disconnection / recheck / call entries that an OLD script version saved in the Entries tab into their own tabs (automatically).
 * Version 6: can run behind the one-time LOADER script (Loader.gs) - then it updates itself from GitHub and never needs pasting again.
 * Version 10: logins survive for 30 days (kept in Script Properties, not only the 6-hour cache) - refreshing the app no longer logs out.
 * Version 9: one active session per user (admin / developer exempt) - a second login logs the other device out, and every such event goes to the 'Unauthorized Logins' tab.
 * Version 5: opening the /exec URL in a browser shows the version and WHICH spreadsheet the entries go to (and how many rows each tab has).
 * After ANY change to this code: Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: New version ▸ Deploy (URL stays the same).
 */

var ENTRY_HEADERS = ['Entry ID','Server time','Visit date','Username','Entered by','Account ID','Consumer name','Report status',
  'Mobile (records)','Mobile (updated)','Meter No. (records)','Meter No. (site)','Meter type','Sub-division',
  'Substation (records)','Feeder (records)','Substation (site)','Feeder (site)','Feeder changed','Pole / DT','Tariff',
  'Sanctioned load','Connected load found (kW)','Reading kWh','Reading kVAh','MD (kW)','MD (kVA)','Meter status',
  'Abnormalities','Abnormality count','Remarks','Latitude','Longitude','GPS accuracy (m)','Device time','Sync mode'];
/* field-work sheets - same columns as the SDO field-work workbook */
var FW_TYPES = {
  survey: { sheet: 'Survey Data', id: 'Survey ID', headers: ['User ID','User Name','Account ID','Consumer Name','Father/Husband','Address','Supply Type','Load','SDO Code','Village','Meter Condition','AC Installed','Mobile Number','Created At','Local Record ID','Survey Type','Survey Status','Connection Type','Survey ID','House Locked','Mobile Refused','Meter Installed','Meter Number','Master Reading','Outstanding','Connected Load','Nature of Supply','House Condition','Consumer Payment Response','Consumer Payment Date','Latitude','Longitude','GPS Accuracy','Remarks','Last Updated By User ID','Last Updated By Name','Last Updated At','Feeded Village Name','Connection Status','Last Pay Date','Last Pay Amount','Updated At','Substation','Feeder','DT','Theft Possibility','Site Reading','Site Max Demand'] },
  disc:   { sheet: 'Disconnection Data', id: 'Disconnection ID', headers: ['Disconnection ID','User ID','User Name','Account ID','Consumer Name','Father/Husband','Address','Supply Type','Load','SDO Code','Disconnection Status','Payment Mode','Committed Payment Date','Meter Status','Site Reading','Mobile Number','House Condition','Remarks','Latitude','Longitude','GPS Accuracy','Created At','Total Outstanding','Last Updated By User ID','Last Updated By Name','Last Updated At','Feeded Village Name','Connection Status','Last Pay Date','Last Pay Amount','Updated At','Video File','Video Key','Video Uploaded At'] },
  rechk:  { sheet: 'Recheck Data', id: 'Recheck ID', headers: ['Recheck ID','User ID','User Name','Account ID','Consumer Name','Father/Husband','Address','Supply Type','Load','SDO Code','Present Status','Payment Mode','Payment Date','Meter Status','Site Reading','Mobile Number','Remarks','Latitude','Longitude','GPS Accuracy','Created At','Total Outstanding','Last Updated By User ID','Last Updated By Name','Last Updated At','Feeded Village Name','Connection Status','Last Pay Date','Last Pay Amount','Updated At','Video File','Video Key','Video Uploaded At'] },
  call:   { sheet: 'Phone Calling Data', id: 'Calling ID', headers: ['Calling ID','User ID','User Name','Account ID','Consumer Name','Father/Husband','Address','Supply Type','Load','SDO Code','Mobile Number','Call Response','Committed Payment Date','Remarks','Created At','Total Outstanding','Last Updated By User ID','Last Updated By Name','Last Updated At','Feeded Village Name','Connection Status','Last Pay Date','Last Pay Amount','Updated At'] }
};
var FW_TEXT = ['Account ID','Mobile Number','Meter Number','SDO Code','User ID','Survey ID','Disconnection ID','Recheck ID','Calling ID'];
var USER_HEADERS = ['Username','Password hash','Role','Full name','Active','Created','Last login','Dashboard access'];
var SESSION_HOURS = 720;   // 30 days - a login stays until log-out, another device logging in, or the user being made inactive
var VERSION = 10;
var LOG_HEADERS = ['Time','Username','Name','Role','Event','This device','This device ID','Other device','Other device ID','Other session started'];
var SHEET_ID = '1QTJkyI14ApxcpD-D1rc7y3eHnjUngjvvJ4Xr92n2r_8';   // "EDD-I Orai Field Entries" - used if the script is not opened from the Sheet

function ss_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  try { return SpreadsheetApp.openById(SHEET_ID); }
  catch (e) {
    if (/permission/i.test(String(e.message || e)))
      throw new Error('PERMISSION MISSING - the web-app link runs a deployment made before the Google Sheets permission was allowed. ' +
        'Fix once in Apps Script: (1) paste Loader.gs version 2 and Save, (2) run updateNow and Allow BOTH permissions, ' +
        '(3) Deploy > Manage deployments > pencil on THIS deployment > Version: New version > Deploy.');
    throw e;
  }
}
function sheet_(name) { ensure_(); return ss_().getSheetByName(name); }
var ENSURED_ = false;
/* creates the Entries / Users tabs, text formats and the default admin if anything is missing */
function ensure_() {
  if (ENSURED_) return; ENSURED_ = true;
  var ss = ss_();
  var en = ss.getSheetByName('Entries');
  if (!en) {
    var first = ss.getSheets()[0];
    var firstHdr = first.getLastRow() ? String(first.getRange(1, 1).getValue()) : '';
    en = (firstHdr === 'Entry ID' || first.getLastRow() === 0) ? first.setName('Entries') : ss.insertSheet('Entries');
  }
  if (en.getLastRow() === 0) en.appendRow(ENTRY_HEADERS);
  en.getRange('A:V').setNumberFormat('@'); en.getRange('AB:AJ').setNumberFormat('@');
  var us = ss.getSheetByName('Users');
  if (!us) us = ss.insertSheet('Users');
  us.getRange('A:E').setNumberFormat('@');           // usernames and password codes must stay text
  if (us.getLastRow() === 0) us.appendRow(USER_HEADERS);
  else if (String(us.getRange(1, 8).getValue()) !== 'Dashboard access') us.getRange(1, 8).setValue('Dashboard access');   // v8: module rights for the dashboard
  if (!findUser_('admin') && countUsers_() === 0) addUser_('admin', 'Admin@123', 'admin', 'Administrator');
}
/* version + the spreadsheet this script really writes to (so a wrong / old deployment is easy to spot) */
function info_() {
  ensure_();
  var moved = 0; try { moved = migrateLegacy_(); } catch (e) {}
  var ss = ss_(), tabs = {};
  ss.getSheets().forEach(function (sh) { tabs[sh.getName()] = Math.max(0, sh.getLastRow() - 1); });
  var bound = false; try { bound = !!SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  return { ok: true, version: VERSION, users: countUsers_(), types: Object.keys(FW_TYPES),
           sheet: ss.getName(), sheetId: ss.getId(), sheetUrl: ss.getUrl(), bound: bound, rows: tabs, moved: moved, loader: (typeof LOADER_VERSION !== 'undefined' ? LOADER_VERSION : 0) };
}
/* entries sent while the sheet ran an old script land in 'Entries' with Sync mode "legacy:<type>:<json row>" - copy them to their own tab once */
function migrateLegacy_() {
  var en = ss_().getSheetByName('Entries'); if (!en || en.getLastRow() < 2) return 0;
  var n = en.getLastRow() - 1, mc = ENTRY_HEADERS.length;
  var modes = en.getRange(2, mc, n, 1).getValues(), todo = [];
  modes.forEach(function (m, k) { if (String(m[0]).indexOf('legacy:') === 0) todo.push(k); });
  if (!todo.length) return 0;
  var info = en.getRange(2, 1, n, 5).getValues(), moved = 0;
  todo.forEach(function (k) {
    var m = /^legacy:(\w+):([\s\S]*)$/.exec(String(modes[k][0])); if (!m || !FW_TYPES[m[1]]) return;
    var row; try { row = JSON.parse(m[2]); } catch (e) { return; }
    var r = info[k], created = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss') : String(r[1]);
    submitTyped_({ username: String(r[3]), name: String(r[4]) }, [{ id: String(r[0]), type: m[1], row: row, created: created }]);
    en.getRange(k + 2, mc).setValue('moved to ' + FW_TYPES[m[1]].sheet); moved++;
  });
  return moved;
}
function countUsers_() { var us = ss_().getSheetByName('Users'); return us ? Math.max(0, us.getLastRow() - 1) : 0; }

/* ---------- one-time setup ---------- */
function be_setup() {
  ENSURED_ = false; ensure_();
  var ss = ss_(), en = ss.getSheetByName('Entries'), us = ss.getSheetByName('Users');
  en.getRange(1, 1, 1, ENTRY_HEADERS.length).setValues([ENTRY_HEADERS]).setFontWeight('bold').setBackground('#e0e7ff'); en.setFrozenRows(1);
  us.getRange(1, 1, 1, USER_HEADERS.length).setFontWeight('bold').setBackground('#fef3c7'); us.setFrozenRows(1);
  Logger.log('Setup done. Default login (if no users existed): admin / Admin@123 - change it from the dashboard.');
}

/* ---------- web endpoints ---------- */
function be_doGet(e) {
  var p = e && e.parameter || {};
  // JSONP mode (used by the dashboard when a direct POST is blocked): ?cb=fn&p={json}
  if (p.cb && p.p) {
    var cb = String(p.cb).replace(/[^\w$.]/g, '');
    var out;
    try { out = handle_(JSON.parse(p.p)); } catch (err) { out = { ok: false, error: String(err.message || err), auth: err.auth === true, kicked: err.kicked === true }; }
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  var out2; try { out2 = info_(); } catch (err) { out2 = { ok: false, version: VERSION, error: String(err.message || err) }; }
  out2.service = 'EDD-I Orai field entry'; out2.time = new Date();
  return json_(out2);
}

function be_doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  try { return json_(handle_(req)); }
  catch (err) { return json_({ ok: false, error: String(err.message || err), auth: err.auth === true, kicked: err.kicked === true }); }
}
function handle_(req) {
  switch (req.action) {
    case 'ping':     return info_();
    case 'login':    return login_(req);
    case 'logout':   return logout_(req);
    case 'check':    auth_(req); return { ok: true };
    case 'submit':   return submit_(auth_(req), req);
    case 'list':     return list_(auth_(req), req);
    case 'users':    return users_(auth_(req, 'admin'));
    case 'saveUser': return saveUser_(auth_(req, 'admin'), req);
    case 'deleteUser': return deleteUser_(auth_(req, 'admin'), req);
    case 'changePassword': return changePassword_(auth_(req), req);
    default:         return { ok: false, error: 'Unknown action' };
  }
}

/* ---------- auth ---------- */
function login_(req) {
  ensure_();
  var u = findUser_(req.username);
  if (!u || u.active !== 'Y' || u.hash !== String(req.hash)) return { ok: false, error: 'Wrong username or password' };
  var token = Utilities.getUuid() + '~' + encodeURIComponent(u.username);
  var sess = { username: u.username, role: u.role, name: u.name };
  var cache = CacheService.getScriptCache();
  cache.put('t_' + token, JSON.stringify(sess), Math.min(SESSION_HOURS * 3600, 21600));
  var warn = '';
  if (!isAdmin_(u.role)) {
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      var props = PropertiesService.getScriptProperties(), key = 'sess_' + u.username;
      var act = null; try { act = JSON.parse(props.getProperty(key) || 'null'); } catch (e) {}
      var me = { t: token, dev: String(req.dev || ''), ua: String(req.ua || '').slice(0, 120), at: now_(), ms: Date.now() };
      var live = act && act.t && (act.ms ? Date.now() - act.ms < SESSION_HOURS * 3600e3 : !!cache.get('t_' + act.t));
      if (live && !(me.dev && act.dev === me.dev)) {
        logUnauth_(u, 'Second login while the user was active on another device - the other device was logged out', me, act);
        warn = 'This user ID was active on another device. That device has been logged out and this login has been recorded.';
      }
      var old = (act && act.old) || []; if (act && act.t) old.unshift({ t: act.t, dev: act.dev, ua: act.ua, at: act.at });
      me.old = old.slice(0, 5);
      props.setProperty(key, JSON.stringify(me));
    } finally { lock.releaseLock(); }
  } else {
    var pa = PropertiesService.getScriptProperties(), ka = 'sessa_' + u.username, list = [];
    try { list = JSON.parse(pa.getProperty(ka) || '[]'); } catch (e) {}
    list = list.filter(function (x) { return Date.now() - x.ms < SESSION_HOURS * 3600e3; });
    list.unshift({ t: token, ms: Date.now() }); pa.setProperty(ka, JSON.stringify(list.slice(0, 10)));
  }
  u.sheet.getRange(u.row, 7).setValue(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'));
  return { ok: true, token: token, username: u.username, role: u.role, name: u.name, dash: u.dash, hours: SESSION_HOURS, single: !isAdmin_(u.role), warn: warn };
}
function isAdmin_(r) { return r === 'admin' || r === 'developer'; }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'); }
function logout_(req) {
  var c = CacheService.getScriptCache(), raw = req.token && c.get('t_' + req.token);
  var s = raw ? JSON.parse(raw) : restore_(req.token);
  c.remove('t_' + req.token);
  if (s) { var props = PropertiesService.getScriptProperties();
    try { var act = JSON.parse(props.getProperty('sess_' + s.username) || 'null'); if (act && act.t === req.token) { act.t = ''; props.setProperty('sess_' + s.username, JSON.stringify(act)); } } catch (e) {}
    try { var l = JSON.parse(props.getProperty('sessa_' + s.username) || '[]'); var l2 = l.filter(function (x) { return x.t !== req.token; }); if (l2.length !== l.length) props.setProperty('sessa_' + s.username, JSON.stringify(l2)); } catch (e) {} }
  return { ok: true };
}
/* the 6-hour cache forgot the token: rebuild the session from Script Properties (logins last SESSION_HOURS) */
function restore_(token) {
  var m = /~([^~]+)$/.exec(String(token || '')); if (!m) return null;
  var un = decodeURIComponent(m[1]), props = PropertiesService.getScriptProperties(), ok = false, max = SESSION_HOURS * 3600e3;
  try { var act = JSON.parse(props.getProperty('sess_' + un) || 'null');
    if (act && act.t === token && Date.now() - (act.ms || 0) < max) ok = true;
    else if (act && (act.old || []).some(function (o) { return o.t === token; })) ok = true;      // an older device - auth_ then logs it out
  } catch (e) {}
  try { if (!ok && JSON.parse(props.getProperty('sessa_' + un) || '[]').some(function (x) { return x.t === token && Date.now() - x.ms < max; })) ok = true; } catch (e) {}
  if (!ok) return null;
  var u = findUser_(un); if (!u || u.active !== 'Y') return null;
  var s = { username: u.username, role: u.role, name: u.name };
  CacheService.getScriptCache().put('t_' + token, JSON.stringify(s), 21600);
  return s;
}
/* the 'Unauthorized Logins' tab - created with its headers the first time something is logged */
function logUnauth_(u, event, me, other) {
  var ss = ss_(), sh = ss.getSheetByName('Unauthorized Logins');
  if (!sh) { sh = ss.insertSheet('Unauthorized Logins'); sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#fee2e2'); sh.setFrozenRows(1); sh.getRange('A:J').setNumberFormat('@'); }
  sh.appendRow([now_(), u.username, u.name || '', u.role || '', event, me.ua || '', me.dev || '', (other && other.ua) || '', (other && other.dev) || '', (other && other.at) || '']);
}
function auth_(req, needRole) {
  var raw = req.token && CacheService.getScriptCache().get('t_' + req.token);
  var s = raw ? JSON.parse(raw) : restore_(req.token);
  if (!s) { var e = new Error('Session expired — please log in again'); e.auth = true; throw e; }
  if (!isAdmin_(s.role)) {
    var act = null; try { act = JSON.parse(PropertiesService.getScriptProperties().getProperty('sess_' + s.username) || 'null'); } catch (x) {}
    if (act && act.t && act.t !== req.token) {
      var c = CacheService.getScriptCache();
      if (!c.get('kl_' + req.token)) {
        c.put('kl_' + req.token, '1', 21600);
        var mine = (act.old || []).filter(function (o) { return o.t === req.token; })[0] || {};
        try { logUnauth_(findUser_(s.username) || { username: s.username, name: s.name, role: s.role }, 'Logged-out device tried to continue - stopped', { dev: String(req.dev || mine.dev || ''), ua: String(req.ua || mine.ua || '') }, act); } catch (x) {}
      }
      var k = new Error("Don't try to Unauthorized login. This user ID is now logged in on another device, so this device has been logged out.");
      k.auth = true; k.kicked = true; throw k;
    }
    if (act && act.t === req.token && act.ms && Date.now() - act.ms > 86400e3) { act.ms = Date.now(); PropertiesService.getScriptProperties().setProperty('sess_' + s.username, JSON.stringify(act)); }   // still in use: keep it alive
  }
  if (needRole === 'admin' && s.role !== 'admin' && s.role !== 'developer') throw new Error('Admin only');
  if (needRole === 'developer' && s.role !== 'developer') throw new Error('Developer only');
  return s;
}

/* ---------- entries ---------- */
function submit_(s, req) {
  var all = req.entries || [];
  var items = all.filter(function (x) { return !x.type || x.type === 'insp'; });
  var typed = all.filter(function (x) { return x.type && x.type !== 'insp'; });
  var savedT = typed.length ? submitTyped_(s, typed) : [];
  if (!items.length) return { ok: true, saved: savedT, sheet: ss_().getName() };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_('Entries');
    var n = sh.getLastRow(), seen = {};
    if (n > 1) {
      var from = Math.max(2, n - 3000);
      sh.getRange(from, 1, n - from + 1, 1).getValues().forEach(function (r) { seen[r[0]] = 1; });
    }
    var rows = [], saved = [];
    items.forEach(function (x) {
      if (!x.id || seen[x.id]) { saved.push(x.id); return; }
      seen[x.id] = 1; saved.push(x.id);
      var o = owner_(s, x);
      rows.push([x.id, x.created || Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), x.visitDate || '', o.u, o.n, x.acct, x.name, x.status,
        x.mobRec, x.mobNew, x.mtrRec, x.mtrSite, x.meterType, x.sdo, x.ssRec, x.fdRec, x.ssSite, x.fdSite,
        x.fdChanged ? 'YES' : 'NO', x.pole, x.tariff, x.load, num_(x.connLoad), num_(x.kwh), num_(x.kvah),
        num_(x.mdKw), num_(x.mdKva), x.meterStatus, (x.abn || []).join('; '), (x.abn || []).length, x.remarks,
        x.lat || '', x.lng || '', x.acc || '', x.deviceTime || '', x.mode || 'online'].map(function (v) { return v == null ? '' : v; }));
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, ENTRY_HEADERS.length).setValues(rows);
    return { ok: true, saved: saved.concat(savedT), added: rows.length, sheet: ss_().getName() };
  } finally { lock.releaseLock(); }
}
/* the sheet tab for a field-work type, created with its headers if missing */
function fwSheet_(t) {
  var T = FW_TYPES[t]; if (!T) throw new Error('Unknown type ' + t);
  ensure_();
  var ss = ss_(), sh = ss.getSheetByName(T.sheet);
  if (!sh) {
    sh = ss.insertSheet(T.sheet);
    sh.getRange(1, 1, 1, T.headers.length).setValues([T.headers]).setFontWeight('bold').setBackground('#dcfce7'); sh.setFrozenRows(1);
    FW_TEXT.forEach(function (h) { var c = T.headers.indexOf(h); if (c >= 0) sh.getRange(1, c + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}
/* who made an entry: the logged-in user, or - when an admin copies entries saved on GitHub - the original field user */
function owner_(s, x) {
  if (x && x.by && x.by.u && (s.role === 'admin' || s.role === 'developer')) return { u: String(x.by.u), n: String(x.by.n || x.by.u) };
  return { u: s.username, n: s.name || s.username };
}
function submitTyped_(s, items) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var saved = [], byType = {};
    items.forEach(function (x) { (byType[x.type] = byType[x.type] || []).push(x); });
    Object.keys(byType).forEach(function (t) {
      var T = FW_TYPES[t]; if (!T) throw new Error('Unknown entry type ' + t + ' - update the dashboard app');
      var sh = fwSheet_(t), H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      var idc = H.indexOf(T.id), n = sh.getLastRow(), seen = {};
      if (idc >= 0 && n > 1) { var from = Math.max(2, n - 5000); sh.getRange(from, idc + 1, n - from + 1, 1).getValues().forEach(function (r) { seen[r[0]] = 1; }); }
      var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), rows = [];
      byType[t].forEach(function (x) {
        if (!x.id || seen[x.id]) { saved.push(x.id); return; }
        seen[x.id] = 1; saved.push(x.id);
        var r = x.row || {}, o = owner_(s, x);
        r[T.id] = x.id; r['User ID'] = o.u; r['User Name'] = o.n;
        r['Created At'] = x.created || now; r['Updated At'] = now;
        if (H.indexOf('Last Updated By User ID') >= 0) { r['Last Updated By User ID'] = o.u; r['Last Updated By Name'] = o.n; r['Last Updated At'] = x.created || now; }
        rows.push(H.map(function (h) { var v = r[h]; return v == null ? '' : v; }));
      });
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, H.length).setValues(rows);
    });
    return saved;
  } finally { lock.releaseLock(); }
}
function list_(s, req) {
  if (req.type && req.type !== 'insp') {
    var tsh = fwSheet_(req.type), tn = tsh.getLastRow(), tc = tsh.getLastColumn();
    var H = tsh.getRange(1, 1, 1, tc).getDisplayValues()[0];
    if (tn < 2) return { ok: true, rows: [], headers: H };
    var tl = Math.min(Number(req.limit) || 1000, 5000), tf = Math.max(2, tn - 8000);
    var tv = tsh.getRange(tf, 1, tn - tf + 1, tc).getDisplayValues(), uc = H.indexOf('User ID');
    var tall = (s.role === 'admin' || s.role === 'developer') && req.scope === 'all', tout = [];
    for (var k = tv.length - 1; k >= 0 && tout.length < tl; k--) if (tall || tv[k][uc] === s.username) tout.push(tv[k]);
    return { ok: true, rows: tout, headers: H };
  }
  var sh = sheet_('Entries');
  var n = sh.getLastRow(); if (n < 2) return { ok: true, rows: [], headers: ENTRY_HEADERS };
  var lim = Math.min(Number(req.limit) || 500, 3000), from = Math.max(2, n - 6000);
  var vals = sh.getRange(from, 1, n - from + 1, ENTRY_HEADERS.length).getDisplayValues();
  var all = (s.role === 'admin' || s.role === 'developer') && req.scope === 'all';
  var out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < lim; i--) {
    if (all || vals[i][3] === s.username) out.push(vals[i]);
  }
  return { ok: true, rows: out, headers: ENTRY_HEADERS };
}

/* ---------- users ---------- */
function users_() {
  var sh = sheet_('Users');
  var v = sh.getDataRange().getDisplayValues().slice(1);
  return { ok: true, users: v.filter(function (r) { return String(r[0]).trim(); }).map(function (r) { return { username: r[0], role: r[2], name: r[3], active: r[4], created: r[5], lastLogin: r[6], dash: r[7] || '' }; }) };
}
function saveUser_(s, req) {
  var un = String(req.username || '').trim().toLowerCase().replace(/\s+/g, '.');
  if (!/^[a-z0-9._@-]{3,40}$/.test(un)) throw new Error('User name must be 3-40 characters: letters, digits, dot, dash, underscore or @ (no spaces).');
  var role = req.role === 'developer' ? 'developer' : req.role === 'admin' ? 'admin' : 'field';
  var anyDev = users_().users.some(function (x) { return x.role === 'developer'; });
  if (role === 'developer' && s.role !== 'developer' && anyDev) throw new Error('Only a developer can create or change developer accounts.');
  var active = req.active === 'N' ? 'N' : 'Y';
  var name = String(req.name || '').trim();
  var hash = String(req.hash || '');
  var dash = req.dash == null ? null : String(req.dash).replace(/[^\w,*+-]/g, '').slice(0, 400);   // dashboard modules: '' none, '*' all, list of module ids, '+upd' = may update / publish
  if (hash && !/^[0-9a-f]{64}$/.test(hash)) throw new Error('Password could not be encoded - reload the dashboard and try again.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var u = findUser_(un);
    if (!u) {
      if (!hash) throw new Error('Enter a password for the new user (at least 6 characters).');
      var sh = sheet_('Users'), row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 5).setNumberFormat('@');
      sh.getRange(row, 1, 1, 8).setValues([[un, hash, role, name || un, active, Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'), '', dash || '']]);
    } else {
      if (u.role === 'developer' && s.role !== 'developer') throw new Error('Only a developer can change a developer account.');
      if (un === s.username && (active === 'N' || role === 'field' || (s.role === 'developer' && role !== 'developer'))) throw new Error('You cannot disable or demote your own account.');
      if (hash) u.sheet.getRange(u.row, 2).setNumberFormat('@').setValue(hash);
      u.sheet.getRange(u.row, 3, 1, 3).setValues([[role, name || u.name, active]]);
      if (dash != null) u.sheet.getRange(u.row, 8).setValue(dash);
    }
  } finally { lock.releaseLock(); }
  var out = users_(); out.saved = un; out.created = !u; return out;
}
function deleteUser_(s, req) {
  var un = String(req.username || '').trim().toLowerCase();
  if (un === s.username) throw new Error('You cannot delete your own account.');
  var u = findUser_(un); if (!u) throw new Error('User not found: ' + un);
  if (u.role === 'developer' && s.role !== 'developer') throw new Error('Only a developer can delete a developer account.');
  u.sheet.deleteRow(u.row);
  return users_();
}
function changePassword_(s, req) {
  var u = findUser_(s.username);
  if (!u || u.hash !== String(req.oldHash)) throw new Error('Current password is wrong');
  if (!/^[0-9a-f]{64}$/.test(String(req.newHash || ''))) throw new Error('Password could not be encoded.');
  u.sheet.getRange(u.row, 2).setNumberFormat('@').setValue(req.newHash);
  return { ok: true };
}
function findUser_(username) {
  var sh = sheet_('Users'); if (!sh) return null;
  var un = String(username || '').trim().toLowerCase();
  var v = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === un)
    return { sheet: sh, row: i + 1, username: String(v[i][0]), hash: String(v[i][1]).trim(), role: String(v[i][2]).trim(), name: String(v[i][3]), active: String(v[i][4]).trim().toUpperCase(), dash: String(v[i][7] || '') };
  return null;
}
function addUser_(un, pass, role, name) {
  var sh = ss_().getSheetByName('Users'), row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 5).setNumberFormat('@');
  sh.getRange(row, 1, 1, 7).setValues([[un, hash_(un, pass), role, name, 'Y', Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm'), '']]);
}
/* same hash the dashboard computes: SHA-256 hex of "username:password" (username lower-case) */
function hash_(un, pass) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(un).toLowerCase() + ':' + pass, Utilities.Charset.UTF_8);
  return b.map(function (x) { var h = (x < 0 ? x + 256 : x).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}
function num_(v) { return v === '' || v == null || isNaN(Number(v)) ? '' : Number(v); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
