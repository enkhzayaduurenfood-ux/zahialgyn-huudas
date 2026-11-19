/**
 * Захиалгын хуудасны сайт — Пи Пи Пи Эм Жи Эл ХХК
 * Өгөгдөл: Google Sheets, зураг: Google Drive
 *
 * Анх суулгахдаа: setup() функцийг нэг удаа ажиллуулна.
 * Каталог оруулах: Drive-д catalog_images.zip байршуулаад importCatalog() ажиллуулна
 *                  (6 минутаас удвал дахин ажиллуулахад үргэлжилнэ).
 */

// Бүх өгөгдөл, зураг хадгалагдах Drive хавтас
var ROOT_FOLDER_ID = '1r48kZSQtz6oxRdaja0ugDRHhIubkBAQ3';
// Каталогийн zip-ийг эндээс татна (Drive-д байхгүй бол)
var CATALOG_ZIP_URL = 'https://raw.githubusercontent.com/enkhzayaduurenfood-ux/zahialgyn-huudas/main/catalog_images.zip';

var SHEET_ORDERS = 'Захиалга';
var SHEET_CATALOG = 'Каталог';
var SHEET_LOG = 'Түүх';
var SHEET_SETTINGS = 'Тохиргоо';

var ORDER_HEADERS = ['ID', 'Захиалгын №', 'Дараалал', 'Он', 'Захиалагч', 'Хүлээж авсан', 'Хүлээлгэн өгөх',
  'Загварын код', 'Загварын тайлбар', 'Алхам', 'Төлөв', 'Нийт тоо', 'Үүсгэсэн', 'Шинэчилсэн', 'Илгээсэн',
  'Захиалагчийн токен', 'Үйлдвэрийн токен', 'Устгасан', 'Өгөгдөл (JSON)'];
var CATALOG_HEADERS = ['Ангилал', 'Код', 'Тодорхойлолт', 'CLO', 'Үлгэр', 'Үнэлгээ', 'Загвар', 'Энгэр зураг', 'Ар зураг', 'Шинэчилсэн'];
var LOG_HEADERS = ['Огноо', 'Хэн', 'Захиалгын №', 'Хуудас ID', 'Үйлдэл', 'Тэмдэглэл'];

// Захиалагч засаж болох талбарууд (ZH хэсэг)
var CUSTOMER_KEYS = ['customer', 'recvDate', 'dueDate', 'modelDesc', 'photos', 'colors', 'sizes', 'extraSizes',
  'khav', 'soilt', 'logo', 'tsats', 'halaas', 'dotor', 'people', 'custNote', 'zhSigns'];
var ORDER_LEVEL_KEYS = ['customer', 'recvDate', 'dueDate'];
var STEPS = ['Захиалга хүлээн авах', 'Захиалгын мэдээлэл нягтлах', 'Захиалгын хуудас бөглөх', 'Захиалга баталгаажуулах', 'Эх загвар, скиз баталгаажуулах'];

/* ------------------------------------------------------------------ */
/* Суулгах                                                             */
/* ------------------------------------------------------------------ */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SS_ID');
  var ss;
  if (ssId) { ss = SpreadsheetApp.openById(ssId); }
  else {
    ss = SpreadsheetApp.create('Захиалгын хуудас — өгөгдлийн сан');
    props.setProperty('SS_ID', ss.getId());
  }
  ensureSheet_(ss, SHEET_ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, SHEET_CATALOG, CATALOG_HEADERS);
  ensureSheet_(ss, SHEET_LOG, LOG_HEADERS);
  var st = ensureSheet_(ss, SHEET_SETTINGS, ['Түлхүүр', 'Утга']);
  if (st.getLastRow() < 2) {
    st.appendRow(['year', new Date().getFullYear()]);
    st.appendRow(['seq', 0]);
  }
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Хуудас1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  var folder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  props.setProperty('FOLDER_ID', folder.getId());
  try { DriveApp.getFileById(ss.getId()).moveTo(folder); } catch (e) { Logger.log('Хүснэгтийг хавтас руу зөөж чадсангүй: ' + e); }
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { Logger.log('Хуваалцах тохиргоо: ' + e); }

  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', token_(10));
  Logger.log('Өгөгдлийн сан: ' + ss.getUrl());
  Logger.log('Зургийн хавтас: ' + folder.getUrl());
  Logger.log('УДИРДЛАГЫН ТҮЛХҮҮР: ' + props.getProperty('ADMIN_KEY'));
  Logger.log('Дараа нь: Deploy → New deployment → Web app. Удирдлагын линк = <web app URL>?k=' + props.getProperty('ADMIN_KEY'));
}

/** Drive-д байршуулсан catalog_images.zip-ээс каталог оруулах (дахин ажиллуулбал үргэлжилнэ) */
function importCatalog() {
  var start = Date.now();
  var props = PropertiesService.getScriptProperties();
  var folder = DriveApp.getFolderById(props.getProperty('FOLDER_ID'));
  var sub = getSubFolder_(folder, 'Каталог');
  var files = folder.getFilesByName('catalog_images.zip');
  var zip;
  if (files.hasNext()) zip = files.next().getBlob();
  else if (CATALOG_ZIP_URL) zip = UrlFetchApp.fetch(CATALOG_ZIP_URL).getBlob();
  else throw new Error('catalog_images.zip олдсонгүй.');
  var blobs = Utilities.unzip(zip.setContentType('application/zip'));
  var catBlob = null, imgs = {};
  blobs.forEach(function (b) { if (b.getName() === 'catalog.json') catBlob = b; else imgs[b.getName()] = b; });
  var items = JSON.parse(catBlob.getDataAsString('UTF-8'));

  var done = JSON.parse(props.getProperty('IMPORT_DONE') || '{}');
  var existing = {};
  var it = sub.getFiles();
  while (it.hasNext()) { var f = it.next(); existing[f.getName()] = f.getId(); }
  var names = Object.keys(imgs);
  for (var i = 0; i < names.length; i++) {
    var n = names[i];
    if (existing[n]) { done[n] = existing[n]; continue; }
    if (Date.now() - start > 300000) {
      props.setProperty('IMPORT_DONE', JSON.stringify(done));
      Logger.log('Хугацаа дуусах гэж байна. ' + Object.keys(done).length + '/' + names.length + ' зураг. importCatalog()-ийг ДАХИН ажиллуулна уу.');
      return;
    }
    done[n] = sub.createFile(imgs[n].setContentType('image/jpeg')).getId();
  }
  props.setProperty('IMPORT_DONE', JSON.stringify(done));

  var sh = ss_().getSheetByName(SHEET_CATALOG);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CATALOG_HEADERS.length).clearContent();
  var now = new Date();
  var rows = items.map(function (x) {
    var fr = (x.front || []).map(function (n) { return done[n]; }).filter(String).join(',');
    var bk = (x.back || []).map(function (n) { return done[n]; }).filter(String).join(',');
    return [x.cat, x.code, x.desc || '', x['CLO'] || '', x['Үлгэр'] || '', x['Үнэлгээ'] || '', x['Загвар'] || '', fr, bk, now];
  });
  if (rows.length) sh.getRange(2, 1, rows.length, CATALOG_HEADERS.length).setValues(rows);
  props.deleteProperty('IMPORT_DONE');
  Logger.log('Каталог бэлэн: ' + rows.length + ' загвар, ' + names.length + ' зураг.');
}

/* ------------------------------------------------------------------ */
/* Вэб                                                                  */
/* ------------------------------------------------------------------ */
// Vercel дээрх сайт энэ API руу POST хүсэлт илгээнэ
var API_FNS = { apiList: apiList, apiGetSheet: apiGetSheet, apiGetOrder: apiGetOrder, apiCreateOrder: apiCreateOrder,
  apiAddSheet: apiAddSheet, apiSave: apiSave, apiSendToFactory: apiSendToFactory, apiDeleteSheet: apiDeleteSheet,
  apiSetSeq: apiSetSeq, apiUpload: apiUpload, apiCatalog: apiCatalog, apiSaveCatalog: apiSaveCatalog,
  apiDeleteCatalog: apiDeleteCatalog, apiLog: apiLog };

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var args = req.args || [];
    var res;
    if (req.fn === 'boot') res = bootData_(args[0] || {});
    else if (API_FNS[req.fn]) res = API_FNS[req.fn].apply(null, args);
    else throw new Error('Үл мэдэгдэх үйлдэл: ' + req.fn);
    out = { ok: true, data: res };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'zahialgyn-huudas' })).setMimeType(ContentService.MimeType.JSON);
}

function bootData_(p) {
  var base = '';
  if (p.v === 'f' && p.t) {
    var r = findBy_('factory', p.t);
    if (!r) return { mode: 'error', error: 'Захиалгын хуудас олдсонгүй.' };
    return { mode: 'factory', base: base, sheet: r.obj, catalog: catalogFor_(r.obj.data.modelCode, r.obj.data.modelCat), title: r.obj.orderNo + ' — ' + (r.obj.data.modelCode || '') };
  }
  if (p.v === 'c' && p.t) {
    var list = sheetsByCustToken_(p.t);
    if (!list.length) return { mode: 'error', error: 'Захиалга олдсонгүй.' };
    return { mode: 'customer', base: base, token: p.t, sheets: list, title: 'Захиалга ' + list[0].orderNo };
  }
  if (p.k && p.k === adminKey_()) {
    return { mode: 'admin', base: base, key: p.k, list: listSheets_(), catalog: allCatalog_(), settings: settings_(), steps: STEPS, title: 'Захиалгын хуудас — удирдлага' };
  }
  return { mode: 'error', error: 'Линк буруу байна. Удирдлагаас зөв линк авна уу.' };
}


/* ------------------------------------------------------------------ */
/* API (google.script.run)                                             */
/* ------------------------------------------------------------------ */
function apiList(auth) { requireAdmin_(auth); return { list: listSheets_(), settings: settings_() }; }

function apiGetSheet(auth, id) {
  var r = findBy_('id', id);
  if (!r) throw new Error('Олдсонгүй');
  if (!isAdmin_(auth) && !(auth && auth.t && auth.t === r.obj.custToken)) throw new Error('Эрх байхгүй');
  return r.obj;
}

function apiGetOrder(auth, orderNo) {
  var rows = sheetsByOrder_(orderNo);
  if (!rows.length) throw new Error('Олдсонгүй');
  if (!isAdmin_(auth) && !(auth && auth.t && auth.t === rows[0].custToken)) throw new Error('Эрх байхгүй');
  if (!isAdmin_(auth)) rows.forEach(function (o) { o.factoryToken = ''; });
  return rows;
}

/** Шинэ захиалга (эхний загварын хуудастай) */
function apiCreateOrder(auth, init) {
  requireAdmin_(auth);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var recv = init.recvDate || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    var year = Number(recv.slice(0, 4));
    var st = settings_();
    var seq = (Number(st.year) === year) ? Number(st.seq) + 1 : 1;
    setSetting_('year', year); setSetting_('seq', seq);
    var orderNo = pad3_(seq) + '-' + recv.slice(2, 4) + recv.slice(5, 7) + recv.slice(8, 10);
    var data = { customer: init.customer || '', recvDate: recv, dueDate: init.dueDate || '', step: 1 };
    var obj = newRow_(orderNo, seq, year, token_(16), data);
    log_(auth.who, orderNo, obj.id, 'Шинэ захиалга', '');
    return obj;
  } finally { lock.releaseLock(); }
}

/** Захиалгад шинэ загварын хуудас нэмэх */
function apiAddSheet(auth, orderNo, copyFromId) {
  var rows = sheetsByOrder_(orderNo);
  if (!rows.length) throw new Error('Захиалга олдсонгүй');
  var first = rows[0];
  var isCust = auth && auth.t && auth.t === first.custToken;
  if (!isAdmin_(auth) && !isCust) throw new Error('Эрх байхгүй');
  var data = { customer: first.data.customer, recvDate: first.data.recvDate, dueDate: first.data.dueDate, step: first.data.step || 1 };
  if (copyFromId) {
    var src = rows.filter(function (r) { return r.id === copyFromId; })[0];
    if (src) { data = JSON.parse(JSON.stringify(src.data)); data.revisions = []; data.photos = data.photos || []; }
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var obj = newRow_(orderNo, first.seq, first.year, first.custToken, data);
    log_(auth.who, orderNo, obj.id, 'Загварын хуудас нэмсэн', '');
    return obj;
  } finally { lock.releaseLock(); }
}

/** Хадгалах: зөвхөн өөрчлөгдсөн талбаруудыг (patch) нийлүүлнэ */
function apiSave(auth, id, patch, revNote) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var r = findBy_('id', id);
    if (!r) throw new Error('Олдсонгүй');
    var admin = isAdmin_(auth);
    var cust = auth && auth.t && auth.t === r.obj.custToken;
    if (!admin && !cust) throw new Error('Эрх байхгүй');
    if (!admin) {
      if (r.obj.sent) throw new Error('Үйлдвэрт илгээгдсэн тул засах боломжгүй. Борлуулалтын менежерт хандана уу.');
      Object.keys(patch).forEach(function (k) { if (CUSTOMER_KEYS.indexOf(k) < 0) delete patch[k]; });
    }
    var d = r.obj.data;
    Object.keys(patch).forEach(function (k) { d[k] = patch[k]; });
    if (r.obj.sent && admin) {
      d.revisions = d.revisions || [];
      d.revisions.push({ at: nowStr_(), who: auth.who || '', note: revNote || '' });
    }
    writeRow_(r.row, r.obj);
    // Захиалгын түвшний талбарыг бусад хуудсанд ч шинэчилнэ
    var orderPatch = {};
    ORDER_LEVEL_KEYS.forEach(function (k) { if (k in patch) orderPatch[k] = patch[k]; });
    if (Object.keys(orderPatch).length) {
      sheetsByOrder_(r.obj.orderNo, true).forEach(function (o) {
        if (o.obj.id === id) return;
        Object.keys(orderPatch).forEach(function (k) { o.obj.data[k] = orderPatch[k]; });
        writeRow_(o.row, o.obj);
      });
    }
    log_(auth.who || (cust ? 'Захиалагч' : ''), r.obj.orderNo, id, r.obj.sent ? 'Илгээсний дараа зассан' : 'Хадгалсан', (revNote || '') + ' [' + Object.keys(patch).join(', ') + ']');
    return findBy_('id', id).obj;
  } finally { lock.releaseLock(); }
}

function apiSendToFactory(auth, id) {
  requireAdmin_(auth);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var r = findBy_('id', id);
    if (!r) throw new Error('Олдсонгүй');
    r.obj.sent = nowStr_();
    if (!r.obj.factoryToken) r.obj.factoryToken = token_(16);
    writeRow_(r.row, r.obj);
    log_(auth.who, r.obj.orderNo, id, 'Үйлдвэрт илгээсэн', '');
    return r.obj;
  } finally { lock.releaseLock(); }
}

function apiDeleteSheet(auth, id) {
  requireAdmin_(auth);
  var r = findBy_('id', id);
  if (!r) return true;
  r.obj.deleted = nowStr_();
  writeRow_(r.row, r.obj);
  log_(auth.who, r.obj.orderNo, id, 'Устгасан', '');
  return true;
}

function apiSetSeq(auth, year, seq) {
  requireAdmin_(auth);
  setSetting_('year', Number(year)); setSetting_('seq', Number(seq));
  return settings_();
}

/** Зураг байршуулах: base64 JPEG → Drive file id */
function apiUpload(auth, b64, name) {
  var ok = isAdmin_(auth) || (auth && auth.t && sheetsByCustToken_(auth.t).length);
  if (!ok) throw new Error('Эрх байхгүй');
  var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  var sub = getSubFolder_(folder, 'Захиалга');
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', (name || 'zurag') + '.jpg');
  var f = sub.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  return f.getId();
}

/* ---- Каталог ---- */
function apiCatalog(auth) { requireAdmin_(auth); return allCatalog_(); }

function apiSaveCatalog(auth, item, origKey) {
  requireAdmin_(auth);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName(SHEET_CATALOG);
    var vals = sh.getDataRange().getValues();
    var row = -1;
    var key = origKey || (item.cat + '|' + item.code);
    for (var i = 1; i < vals.length; i++) { if (vals[i][0] + '|' + vals[i][1] === key) { row = i + 1; break; } }
    var rec = [item.cat, item.code, item.desc || '', item.clo || '', item.ulger || '', item.unelgee || '', item.zagvar || '',
      (item.front || []).join(','), (item.back || []).join(','), new Date()];
    if (row < 0) sh.appendRow(rec); else sh.getRange(row, 1, 1, rec.length).setValues([rec]);
    log_(auth.who, '', '', row < 0 ? 'Каталогт нэмсэн' : 'Каталог зассан', item.cat + ' ' + item.code);
    return allCatalog_();
  } finally { lock.releaseLock(); }
}

function apiDeleteCatalog(auth, key) {
  requireAdmin_(auth);
  var sh = ss_().getSheetByName(SHEET_CATALOG);
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) { if (vals[i][0] + '|' + vals[i][1] === key) { sh.deleteRow(i + 1); break; } }
  log_(auth.who, '', '', 'Каталогоос устгасан', key);
  return allCatalog_();
}

function apiLog(auth, orderNo) {
  requireAdmin_(auth);
  var vals = ss_().getSheetByName(SHEET_LOG).getDataRange().getDisplayValues().slice(1);
  return vals.filter(function (v) { return !orderNo || v[2] === orderNo; }).slice(-200).reverse();
}

/* ------------------------------------------------------------------ */
/* Дотоод                                                              */
/* ------------------------------------------------------------------ */
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID')); }
function tz_() { return Session.getScriptTimeZone() || 'Asia/Ulaanbaatar'; }
function nowStr_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function pad3_(n) { return ('00' + n).slice(-Math.max(3, String(n).length)); }
function token_(n) { var c = 'abcdefghijkmnpqrstuvwxyz23456789', s = ''; for (var i = 0; i < n; i++) s += c.charAt(Math.floor(Math.random() * c.length)); return s; }
function adminKey_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'); }
function isAdmin_(auth) { return !!(auth && auth.k && auth.k === adminKey_()); }
function requireAdmin_(auth) { if (!isAdmin_(auth)) throw new Error('Удирдлагын эрх шаардлагатай'); }

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getSubFolder_(folder, name) {
  var it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}

function settings_() {
  var vals = ss_().getSheetByName(SHEET_SETTINGS).getDataRange().getValues();
  var o = {}; vals.slice(1).forEach(function (v) { o[v[0]] = v[1]; });
  return o;
}
function setSetting_(k, v) {
  var sh = ss_().getSheetByName(SHEET_SETTINGS);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) { if (vals[i][0] === k) { sh.getRange(i + 1, 2).setValue(v); return; } }
  sh.appendRow([k, v]);
}

function log_(who, orderNo, id, action, note) {
  try { ss_().getSheetByName(SHEET_LOG).appendRow([nowStr_(), who || '', orderNo || '', id || '', action, note || '']); } catch (e) { }
}

function sizeTotal_(d) {
  var t = 0;
  ['M', 'F'].forEach(function (g) { var s = (d.sizes || {})[g] || {}; Object.keys(s).forEach(function (k) { t += Number(s[k]) || 0; }); });
  (d.extraSizes || []).forEach(function (x) { t += Number(x.qty) || 0; });
  return t;
}

function rowToObj_(v) {
  var data = {};
  try { data = JSON.parse(v[18] || '{}'); } catch (e) { }
  return {
    id: String(v[0]), orderNo: String(v[1]), seq: v[2], year: v[3], created: fmt_(v[12]), updated: fmt_(v[13]),
    sent: fmt_(v[14]), custToken: String(v[15] || ''), factoryToken: String(v[16] || ''), deleted: fmt_(v[17]), data: data
  };
}
function fmt_(x) { if (!x) return ''; if (x instanceof Date) return Utilities.formatDate(x, tz_(), 'yyyy-MM-dd HH:mm'); return String(x); }

function objToRow_(o) {
  var d = o.data || {};
  return [o.id, o.orderNo, o.seq, o.year, d.customer || '', d.recvDate || '', d.dueDate || '', d.modelCode || '', d.modelDesc || '',
    STEPS[(d.step || 1) - 1] || '', o.sent ? 'Үйлдвэрт илгээсэн' : 'Боловсруулж байна', sizeTotal_(d), o.created, o.updated,
    o.sent || '', o.custToken, o.factoryToken || '', o.deleted || '', JSON.stringify(d)];
}

function newRow_(orderNo, seq, year, custToken, data) {
  var sh = ss_().getSheetByName(SHEET_ORDERS);
  var now = nowStr_();
  var obj = { id: token_(10), orderNo: orderNo, seq: seq, year: year, created: now, updated: now, sent: '', custToken: custToken, factoryToken: '', deleted: '', data: data };
  var r = objToRow_(obj);
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, r.length).setNumberFormat('@').setValues([r]);
  return obj;
}

function writeRow_(row, obj) {
  obj.updated = nowStr_();
  var sh = ss_().getSheetByName(SHEET_ORDERS);
  var r = objToRow_(obj);
  sh.getRange(row, 1, 1, r.length).setNumberFormat('@').setValues([r]);
}

function allRows_() {
  var sh = ss_().getSheetByName(SHEET_ORDERS);
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, ORDER_HEADERS.length).getDisplayValues().map(function (v, i) { return { row: i + 2, v: v }; });
}

function findBy_(kind, val) {
  var col = kind === 'id' ? 0 : kind === 'factory' ? 16 : 15;
  var rows = allRows_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].v[col] === val && !rows[i].v[17]) return { row: rows[i].row, obj: rowToObj_(rows[i].v) };
  }
  return null;
}

function sheetsByOrder_(orderNo, withRow) {
  return allRows_().filter(function (r) { return r.v[1] === orderNo && !r.v[17]; })
    .map(function (r) { return withRow ? { row: r.row, obj: rowToObj_(r.v) } : rowToObj_(r.v); });
}
function sheetsByCustToken_(t) {
  return allRows_().filter(function (r) { return r.v[15] === t && !r.v[17]; }).map(function (r) {
    var o = rowToObj_(r.v); o.factoryToken = ''; return o;
  });
}

function listSheets_() {
  return allRows_().filter(function (r) { return !r.v[17]; }).map(function (r) {
    var v = r.v, d = {};
    try { d = JSON.parse(v[18] || '{}'); } catch (e) { }
    var thumb = (d.photos && d.photos[0]) || '';
    return { id: v[0], orderNo: v[1], seq: Number(v[2]), year: Number(v[3]), customer: d.customer || '', recvDate: d.recvDate || '', dueDate: d.dueDate || '', modelCode: d.modelCode || '', modelCat: d.modelCat || '',
      modelDesc: v[8], step: d.step || 1, sent: v[14], total: Number(v[11]) || 0, updated: v[13], custToken: v[15], factoryToken: v[16], thumb: thumb };
  }).reverse();
}

function allCatalog_() {
  var sh = ss_().getSheetByName(SHEET_CATALOG);
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, CATALOG_HEADERS.length).getDisplayValues().map(function (v) {
    return { cat: v[0], code: v[1], desc: v[2], clo: v[3], ulger: v[4], unelgee: v[5], zagvar: v[6],
      front: v[7] ? v[7].split(',') : [], back: v[8] ? v[8].split(',') : [] };
  }).filter(function (x) { return x.code; });
}

function catalogFor_(code, cat) {
  if (!code) return null;
  var all = allCatalog_();
  return all.filter(function (x) { return x.code === code && (!cat || x.cat === cat); })[0] ||
    all.filter(function (x) { return x.code === code; })[0] || null;
}
