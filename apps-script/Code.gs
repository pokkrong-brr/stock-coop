/**
 * ระบบสต็อกสินค้า สหกรณ์สระว่ายน้ำโรงเรียน v5 (Google Sheets + Apps Script)
 *
 * ชีต: สินค้า, ความเคลื่อนไหว, ตั้งค่า, เจ้าหน้าที่, ผู้จำหน่าย, ใบสั่งซื้อ, ปิดยอด, บันทึกการแก้ไข, รูปสินค้า
 *
 * ออกแบบให้เร็ว:
 * - อ่าน/เขียนชีตแบบรวดเดียว (batch) ไม่เขียนทีละช่อง
 * - หน้าเว็บโหลดเฉพาะความเคลื่อนไหว 90 วันล่าสุด และขอเฉพาะแถวใหม่ (delta) ทุกครั้งที่ซิงก์
 * - ข้อมูลประกอบ (ผู้จำหน่าย ใบสั่งซื้อ ปิดยอด รูป) ส่งเฉพาะเมื่อมีการเปลี่ยน (เลขเวอร์ชัน)
 * - ทุกคำสั่งบันทึกมีรหัส opId กันบันทึกซ้ำเมื่อเน็ตหลุดแล้วส่งใหม่
 * - LockService กันหลายเครื่องตัดสต็อกพร้อมกัน
 *
 * ความปลอดภัย:
 * - เจ้าหน้าที่ต้องเข้าสู่ระบบ ทุกคำสั่งตรวจ token ฝั่งเซิร์ฟเวอร์
 * - รหัสผ่านเก็บเป็น SHA-256 + salt ไม่เก็บตัวจริง
 * - ผิด 5 ครั้งล็อก 5 นาที, เซสชันหมดอายุเมื่อไม่ใช้งาน 6 ชั่วโมง
 * - ผู้ดูแลตั้งรหัสใหม่ให้ใครก็ได้ โดยพิมพ์ในคอลัมน์ "ตั้งรหัสใหม่" ของชีตเจ้าหน้าที่
 *
 * งานอัตโนมัติ (ติดตั้งด้วย setup หรือปุ่มในหน้าตั้งค่า):
 * - 08:00 อีเมลแจ้งสินค้าใกล้หมด · 18:00 อีเมลสรุปยอดขาย · อาทิตย์ 02:00 สำรองชีตลง Google Drive
 */
var SH_P = 'สินค้า', SH_M = 'ความเคลื่อนไหว', SH_S = 'ตั้งค่า', SH_U = 'เจ้าหน้าที่', SH_SUP = 'ผู้จำหน่าย',
  SH_PO = 'ใบสั่งซื้อ', SH_CL = 'ปิดยอด', SH_LOG = 'บันทึกการแก้ไข', SH_IMG = 'รูปสินค้า';
var P_HEAD = ['id', 'รหัส', 'ชื่อสินค้า', 'หมวด', 'หน่วย', 'ราคาทุน', 'ราคาขาย', 'คงเหลือ', 'จุดสั่งซื้อ', 'ใช้งาน', 'id ผู้จำหน่าย'];
var M_HEAD = ['id', 'วันเวลา', 'ประเภท', 'id สินค้า', 'รหัส', 'ชื่อสินค้า', 'จำนวน', 'ราคาขาย', 'ราคาทุน', 'มูลค่าขาย', 'บิล', 'หมายเหตุ', 'ผู้บันทึก', 'ชำระโดย'];
var S_HEAD = ['ชื่อ', 'ค่า'];
var U_HEAD = ['ชื่อผู้ใช้', 'ชื่อที่แสดง', 'บทบาท', 'ใช้งาน', 'ตั้งรหัสใหม่', 'salt', 'hash', 'เข้าใช้ล่าสุด', 'สิทธิ์'];
/** สิทธิ์ที่ผู้ดูแลกำหนดให้เจ้าหน้าที่แต่ละคนได้ (ผู้ดูแลมีทุกสิทธิ์เสมอ) */
var PERMS = { sell: 'ขายสินค้า', refund: 'คืนสินค้า/ยกเลิกบิลของตัวเองวันนี้', refundAll: 'คืนสินค้า/ยกเลิกบิลใดก็ได้', recv: 'รับสินค้าเข้า',
  adj: 'ตัดออก / ตรวจนับสต็อก', cash: 'ปิดยอดเงินสด', po: 'ดูและรับของตามใบสั่งซื้อ', poManage: 'สร้างใบสั่งซื้อ / จัดการผู้จำหน่าย',
  prod: 'เพิ่ม/แก้สินค้าและราคา', report: 'ดูรายงานยอดขายและกำไร', hist: 'ดูประวัติความเคลื่อนไหว' };
var DEFAULT_PERMS = ['sell', 'refund', 'recv', 'adj', 'cash', 'po', 'report', 'hist'];
var SUP_HEAD = ['id', 'ชื่อร้าน/ผู้จำหน่าย', 'ผู้ติดต่อ', 'โทร', 'LINE/อีเมล', 'หมายเหตุ', 'ใช้งาน'];
var PO_HEAD = ['เลขที่', 'วันที่', 'id ผู้จำหน่าย', 'ผู้จำหน่าย', 'สถานะ', 'รายการ (JSON)', 'ยอดรวม', 'ผู้สร้าง', 'วันที่รับ', 'ผู้รับ', 'หมายเหตุ'];
var CL_HEAD = ['วันที่', 'เวลาปิด', 'ผู้ปิดยอด', 'เงินทอนตั้งต้น', 'ขายเงินสด', 'คืนเงินสด', 'ควรมีในลิ้นชัก', 'นับได้', 'ผลต่าง', 'ขายโอน (สุทธิ)', 'หมายเหตุ'];
var LOG_HEAD = ['เวลา', 'ผู้ทำ', 'รายการ', 'รายละเอียด'];
var IMG_HEAD = ['id สินค้า', 'รูป (data URL)'];
var ROLE_ADMIN = 'ผู้ดูแล', ROLE_STAFF = 'เจ้าหน้าที่';
var SESSION_SEC = 21600;
var TYPE_TH = { sale: 'ขาย', in: 'รับเข้า', adj: 'ตัดออก', count: 'ตรวจนับ', ret: 'คืนสินค้า' };
var TYPE_EN = { 'ขาย': 'sale', 'รับเข้า': 'in', 'ตัดออก': 'adj', 'ตรวจนับ': 'count', 'คืนสินค้า': 'ret' };
var PAY_TRANSFER = 'โอน';
var DEFAULT_DAYS = 90;
var SKEYS = { shop: 'ชื่อร้าน', cats: 'หมวดหมู่', emails: 'อีเมลแจ้งเตือน', notifyLow: 'แจ้งสินค้าใกล้หมดตอนเช้า',
  notifySummary: 'ส่งสรุปยอดขายตอนเย็น', backup: 'สำรองข้อมูลทุกสัปดาห์', float: 'เงินทอนตั้งต้น', lastBackup: 'สำรองล่าสุด' };
var AUX = ['suppliers', 'pos', 'closings', 'images'];

/* ---------- หน้าเว็บ ---------- */
/* ---------- API สำหรับหน้าเว็บที่อยู่นอก Google (เช่น GitHub Pages) ----------
   หน้าเว็บส่ง POST {fn, args} มาที่ลิงก์ /exec แล้วได้ {ok, r} หรือ {ok:false, e} กลับไป
   เรียกได้เฉพาะฟังก์ชันในรายการนี้ และทุกฟังก์ชัน (ยกเว้น login) ตรวจ token และสิทธิ์เองอยู่แล้ว */
var API_FNS = ['login', 'logout', 'changePassword', 'listUsers', 'saveUser', 'getData', 'getUpdates', 'loadOlder',
  'sell', 'refund', 'moveStock', 'saveProduct', 'setProductActive', 'saveImage', 'saveSettings', 'saveSupplier',
  'savePO', 'cancelPO', 'receivePO', 'closeCash', 'adminInfo', 'installTriggers', 'sendTestEmail', 'backupNow', 'healthCheck'];
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e) {
  var b = null;
  try {
    b = JSON.parse(e && e.postData ? e.postData.contents : '{}');
    if (API_FNS.indexOf(b.fn) < 0) throw E_('ไม่รู้จักคำสั่ง ' + b.fn);
    var fn = globalThis[b.fn];
    return json_({ ok: true, r: fn.apply(null, Array.isArray(b.args) ? b.args : []) });
  } catch (err) {
    var msg = String(err && err.message || err);
    if (msg.indexOf('ERR:') < 0) { // ข้อผิดพลาดที่ไม่คาดคิด → เก็บไว้ให้ผู้ดูแลเห็นในบันทึกการแก้ไข (ไม่เกินนาทีละครั้ง)
      console.error(msg);
      try { if (!cache_().get('elog')) { cache_().put('elog', '1', 60); log_('ระบบ', 'ข้อผิดพลาด', (b && b.fn || '?') + ': ' + msg.slice(0, 300)); } } catch (e2) {}
    }
    return json_({ ok: false, e: msg });
  }
}
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'ping') return json_({ ok: true, r: 'สต็อกสหกรณ์ พร้อมใช้งาน' });
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('สต็อกสหกรณ์สระว่ายน้ำ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('สต็อกสหกรณ์')
    .addItem('ตั้งค่าเริ่มต้น / อัปเกรด (สร้างชีต + งานอัตโนมัติ)', 'setup')
    .addItem('ส่งอีเมลแจ้งใกล้หมดตอนนี้', 'morningJob')
    .addItem('ส่งอีเมลสรุปยอดวันนี้ตอนนี้', 'eveningJob')
    .addItem('สำรองข้อมูลตอนนี้', 'backupJob')
    .addToUi();
}

/* ---------- ตัวช่วย ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Bangkok'; }
function fmt_(d, f) { return Utilities.formatDate(d, tz_(), f); }
function dkey_(d) { return fmt_(d instanceof Date ? d : new Date(d), 'yyyy-MM-dd'); }
function uid_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
/** ข้อความจากผู้ใช้ที่ขึ้นต้นด้วย = + - @ จะถูกบันทึกเป็นข้อความธรรมดา ไม่ให้ Sheets ตีความเป็นสูตร */
function safe_(v) {
  if (typeof v !== 'string') return v;
  v = v.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, 1000);
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}
function safeRow_(r) { return r.map(safe_); }
function num_(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function bool_(v, dflt) { if (v === '' || v === null || v === undefined) return !!dflt; return v === true || String(v).toUpperCase() === 'TRUE' || v === 1 || v === 'เปิด'; }
function isActive_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === 1; }
function money_(n) { return '฿' + num_(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
/** ข้อผิดพลาดที่ผู้ใช้ต้องรู้ (หน้าเว็บจะไม่ส่งซ้ำ) */
function E_(msg) { return new Error('ERR:' + msg); }
var HDR_OK_ = {};
function sheet_(name, head) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (!HDR_OK_[name]) { // อัปเกรดหัวตารางเมื่อมีคอลัมน์ใหม่
    var cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
    if (cur.join('|') !== head.join('|')) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  }
  HDR_OK_[name] = true;
  return sh;
}
function rows_(name, head) {
  var sh = sheet_(name, head), n = sh.getLastRow() - 1;
  return { sh: sh, vals: n > 0 ? sh.getRange(2, 1, n, head.length).getValues() : [] };
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('ระบบกำลังบันทึกจากเครื่องอื่น ลองใหม่อีกครั้ง');
  try { var r = fn(); SpreadsheetApp.flush(); return r; } finally { lock.releaseLock(); }
}
function props_() { return PropertiesService.getScriptProperties(); }
function bump_(k) { props_().setProperty('v_' + k, uid_()); }
function log_(who, what, detail) {
  var sh = sheet_(SH_LOG, LOG_HEAD);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([safeRow_([new Date(), who, what, detail || ''])]);
}

/* ---------- สร้างชีต / อัปเกรด ---------- */
function setup() {
  var p = sheet_(SH_P, P_HEAD), m = sheet_(SH_M, M_HEAD), s = sheet_(SH_S, S_HEAD);
  [[SH_SUP, SUP_HEAD], [SH_PO, PO_HEAD], [SH_CL, CL_HEAD], [SH_LOG, LOG_HEAD], [SH_IMG, IMG_HEAD]].forEach(function (x) { sheet_(x[0], x[1]); });
  m.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var cur = readSettingsRaw_();
  var defs = [['shop', 'สหกรณ์สระว่ายน้ำโรงเรียน'], ['cats', 'อุปกรณ์ว่ายน้ำ, ชุดว่ายน้ำ, ของใช้, เครื่องดื่ม, ขนม'], ['emails', ''],
    ['notifyLow', true], ['notifySummary', true], ['backup', true], ['float', 500]];
  defs.forEach(function (d) { if (cur[SKEYS[d[0]]] === undefined) s.appendRow([SKEYS[d[0]], d[1]]); });
  ensureAdmin_();
  if (p.getLastRow() < 2) {
    var rows = [
      ['A001', 'แว่นตาว่ายน้ำ เด็ก', 'อุปกรณ์ว่ายน้ำ', 'อัน', 90, 150, 20, 5],
      ['A002', 'แว่นตาว่ายน้ำ ผู้ใหญ่', 'อุปกรณ์ว่ายน้ำ', 'อัน', 120, 199, 15, 5],
      ['A003', 'หมวกว่ายน้ำ ซิลิโคน', 'อุปกรณ์ว่ายน้ำ', 'ใบ', 60, 100, 25, 5],
      ['A004', 'ที่อุดหู', 'อุปกรณ์ว่ายน้ำ', 'คู่', 15, 30, 30, 10],
      ['S001', 'ชุดว่ายน้ำ นักเรียนหญิง', 'ชุดว่ายน้ำ', 'ชุด', 220, 350, 12, 3],
      ['S002', 'กางเกงว่ายน้ำ นักเรียนชาย', 'ชุดว่ายน้ำ', 'ตัว', 110, 180, 15, 3],
      ['U001', 'ผ้าขนหนู', 'ของใช้', 'ผืน', 70, 120, 10, 3],
      ['D001', 'น้ำดื่ม 600 มล.', 'เครื่องดื่ม', 'ขวด', 5, 7, 48, 12],
      ['K001', 'ขนมปังแซนด์วิช', 'ขนม', 'ชิ้น', 12, 15, 4, 6]
    ].map(function (r) { return [uid_()].concat(r).concat([true, '']); });
    p.getRange(2, 1, rows.length, P_HEAD.length).setValues(rows);
  }
  try { installTriggers_(); } catch (e) { /* ผู้ใช้อาจยังไม่อนุญาตสิทธิ์ ตั้งเวลาได้ภายหลังในหน้าตั้งค่า */ }
}

/* ---------- สินค้า ---------- */
function loadP_() {
  var R = rows_(SH_P, P_HEAD), fixIds = false, byId = {};
  R.vals.forEach(function (r, i) {
    if (!String(r[2]).trim()) return;
    if (!r[0]) { r[0] = uid_(); fixIds = true; } // แถวที่พิมพ์เพิ่มเองในชีต
    byId[String(r[0])] = i;
  });
  if (fixIds) R.sh.getRange(2, 1, R.vals.length, 1).setValues(R.vals.map(function (r) { return [r[0]]; }));
  return { sh: R.sh, vals: R.vals, byId: byId };
}
function pObj_(r) {
  return { id: String(r[0]), code: String(r[1]), name: String(r[2]), cat: String(r[3]), unit: String(r[4] || 'ชิ้น'),
    cost: num_(r[5]), price: num_(r[6]), stock: num_(r[7]), min: num_(r[8]), active: bool_(r[9], true), sup: String(r[10] || '') };
}
function productsOut_(P) { return P.vals.filter(function (r) { return r[0] && String(r[2]).trim(); }).map(pObj_); }
/** เขียนคอลัมน์ ทุน/ราคาขาย/คงเหลือ กลับทีเดียว */
function writeNums_(P) {
  if (P.vals.length) P.sh.getRange(2, 6, P.vals.length, 3).setValues(P.vals.map(function (r) { return [r[5], r[6], r[7]]; }));
}

/* ---------- ความเคลื่อนไหว ---------- */
function mSheet_() { return sheet_(SH_M, M_HEAD); }
function rowToMove_(r) {
  var type = TYPE_EN[r[2]] || r[2], q = num_(r[6]);
  return {
    id: String(r[0]), t: r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(), type: type,
    pid: String(r[3]), code: String(r[4]), pname: String(r[5]),
    qty: type === 'sale' ? -q : q, // ในชีตการขายเป็นค่าลบ, คืนสินค้าเป็นค่าบวก (ของกลับเข้าสต็อก)
    price: num_(r[7]), cost: num_(r[8]), bill: String(r[10]), note: String(r[11]), who: String(r[12]), pay: String(r[13] || '')
  };
}
function readMoves_(from, to) {
  if (to < from) return [];
  return mSheet_().getRange(from, 1, to - from + 1, M_HEAD.length).getValues().filter(function (r) { return r[0]; }).map(rowToMove_);
}
/** หาแถวแรกที่วันเวลา >= since (ข้อมูลเรียงตามเวลา ใช้ binary search) */
function firstRowSince_(since) {
  var sh = mSheet_(), last = sh.getLastRow();
  if (last < 2) return 2;
  var col = sh.getRange(2, 2, last - 1, 1).getValues(), lo = 0, hi = col.length;
  while (lo < hi) {
    var mid = (lo + hi) >> 1, v = col[mid][0], t = v instanceof Date ? v.getTime() : new Date(v).getTime();
    if (t < since) lo = mid + 1; else hi = mid;
  }
  return lo + 2;
}
function movesSince_(since) { return readMoves_(firstRowSince_(since), mSheet_().getLastRow()); }
function appendMoves_(rows) {
  if (!rows.length) return;
  var sh = mSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, M_HEAD.length).setValues(rows.map(safeRow_));
}
/** opId นี้ถูกบันทึกไปแล้วหรือยัง (ดู 500 แถวล่าสุด) */
function done_(opId) {
  var sh = mSheet_(), last = sh.getLastRow();
  if (!opId || last < 2) return false;
  var from = Math.max(2, last - 499), ids = sh.getRange(from, 1, last - from + 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]).indexOf(opId + '-') === 0) return true;
  return false;
}
function moveRow_(id, type, r, qty, x) {
  x = x || {};
  var price = x.price != null ? x.price : num_(r[6]), cost = x.cost != null ? x.cost : num_(r[5]);
  var val = type === 'sale' ? qty * price : type === 'ret' ? -qty * price : '';
  return [id, new Date(), TYPE_TH[type], r[0], r[1], r[2], type === 'sale' ? -qty : qty, price, cost, val,
    x.bill || '', x.note || '', x.who || '', x.pay || ''];
}

/* ---------- ตั้งค่า ---------- */
function readSettingsRaw_() {
  var R = rows_(SH_S, S_HEAD), o = {};
  R.vals.forEach(function (r) { if (r[0] !== '') o[String(r[0])] = r[1]; });
  return o;
}
function readSettings_() {
  var o = readSettingsRaw_();
  return {
    shop: String(o[SKEYS.shop] || 'สหกรณ์สระว่ายน้ำโรงเรียน'),
    cats: String(o[SKEYS.cats] || 'ทั่วไป').split(',').map(function (s) { return s.trim(); }).filter(String),
    emails: String(o[SKEYS.emails] || ''),
    notifyLow: bool_(o[SKEYS.notifyLow], true), notifySummary: bool_(o[SKEYS.notifySummary], true), backup: bool_(o[SKEYS.backup], true),
    float: o[SKEYS.float] === undefined || o[SKEYS.float] === '' ? 500 : num_(o[SKEYS.float]),
    lastBackup: o[SKEYS.lastBackup] instanceof Date ? o[SKEYS.lastBackup].getTime() : (o[SKEYS.lastBackup] ? String(o[SKEYS.lastBackup]) : '')
  };
}
function setSetting_(key, val) {
  var R = rows_(SH_S, S_HEAD), label = SKEYS[key];
  for (var i = 0; i < R.vals.length; i++) if (String(R.vals[i][0]) === label) { R.sh.getRange(i + 2, 2).setValue(safe_(val)); return; }
  R.sh.appendRow([label, safe_(val)]);
}
function emailList_() {
  return readSettings_().emails.split(/[,;\s]+/).filter(function (e) { return /@/.test(e); });
}

/* ---------- เจ้าหน้าที่และการเข้าสู่ระบบ ---------- */
function cache_() { return CacheService.getScriptCache(); }
function hash_(salt, pw) { // แบบเดิม (v1) ใช้ตรวจบัญชีเก่าแล้วอัปเกรดเป็น v2 ตอนเข้าสู่ระบบ
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8));
}
var HASH_ROUNDS = 5000;
var SHA_K_ = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
  0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
  0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
  0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
/** SHA-256 ของข้อความ ASCII → hex (เขียนด้วย JS ล้วน เร็วกว่าการเรียก Utilities ซ้ำหลายพันครั้ง) */
function sha256hex_(str) {
  var K = SHA_K_, H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  var l = str.length, nw = (((l + 8) >> 6) + 1) * 16, m = new Array(nw), w = new Array(64), i, j;
  for (i = 0; i < nw; i++) m[i] = 0;
  for (i = 0; i < l; i++) m[i >> 2] |= (str.charCodeAt(i) & 255) << (24 - (i % 4) * 8);
  m[l >> 2] |= 0x80 << (24 - (l % 4) * 8);
  m[nw - 1] = l * 8;
  for (i = 0; i < nw; i += 16) {
    var a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
    for (j = 0; j < 64; j++) {
      if (j < 16) w[j] = m[i + j];
      else {
        var x = w[j - 15], y = w[j - 2];
        w[j] = (((x >>> 7 | x << 25) ^ (x >>> 18 | x << 14) ^ (x >>> 3)) + w[j - 7] + ((y >>> 17 | y << 15) ^ (y >>> 19 | y << 13) ^ (y >>> 10)) + w[j - 16]) | 0;
      }
      var t1 = (h + ((e >>> 6 | e << 26) ^ (e >>> 11 | e << 21) ^ (e >>> 25 | e << 7)) + ((e & f) ^ (~e & g)) + K[j] + w[j]) | 0;
      var t2 = (((a >>> 2 | a << 30) ^ (a >>> 13 | a << 19) ^ (a >>> 22 | a << 10)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
    H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
  }
  return H.map(function (v) { return ('00000000' + (v >>> 0).toString(16)).slice(-8); }).join('');
}
/** v3: SHA-256 วนซ้ำ 5,000 รอบ (รอบแรกใช้ Utilities รองรับภาษาไทย) */
function hash3_(salt, pw) {
  var h = hash_(salt, pw);
  for (var i = 0; i < HASH_ROUNDS; i++) h = sha256hex_(h + salt);
  return 'v3$' + h;
}
/** v2: SHA-256 วนซ้ำ 1,500 รอบ ทำให้เดารหัสจากค่าในชีตช้าลงมาก */
function hash2_(salt, pw) {
  var h = hash_(salt, pw);
  for (var i = 0; i < 1500; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return 'v2$' + h;
}
function same_(a, b) { a = String(a); b = String(b); var d = a.length ^ b.length; for (var i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0); return d === 0; }
function checkPw_(r, pw) { var h = String(r[6] || ''); return h.indexOf('v3$') === 0 ? same_(hash3_(r[5], pw), h) : h.indexOf('v2$') === 0 ? same_(hash2_(r[5], pw), h) : !!h && same_(hash_(r[5], pw), h); }
var WEAK_PW = ['admin1234', '12345678', '123456', 'password', 'qwerty123', '11111111', '00000000'];
function pwRule_(pw) {
  pw = String(pw || '');
  if (pw.length < 8) throw E_('รหัสผ่านต้องยาวอย่างน้อย 8 ตัว');
  if (WEAK_PW.indexOf(pw.toLowerCase()) >= 0 || /^(.)\1+$/.test(pw)) throw E_('รหัสผ่านนี้เดาง่ายเกินไป');
}
var SESSION_MAX_MS = 12 * 3600 * 1000; // เข้าสู่ระบบได้นานสุด 12 ชั่วโมง แล้วต้องเข้าใหม่
/** ถ้ายังไม่มีบัญชีเลย สร้างผู้ดูแลเริ่มต้น admin / admin1234 */
function ensureAdmin_() {
  var u = sheet_(SH_U, U_HEAD);
  if (u.getLastRow() < 2) {
    u.getRange(2, 1, 1, 5).setValues([['admin', 'ผู้ดูแลระบบ', ROLE_ADMIN, true, 'admin1234']]);
    readUsers_();
  }
}
/** อ่านบัญชีทั้งหมด; แถวที่มีค่าใน "ตั้งรหัสใหม่" จะถูกเข้ารหัสแล้วล้างช่องนั้นทิ้ง */
function readUsers_() {
  var R = rows_(SH_U, U_HEAD), dirty = false;
  R.vals.forEach(function (r) {
    if (String(r[4]) !== '') { r[5] = uid_() + uid_(); r[6] = hash3_(r[5], String(r[4])); r[4] = ''; dirty = true; }
  });
  if (dirty) R.sh.getRange(2, 1, R.vals.length, U_HEAD.length).setValues(R.vals);
  return R;
}
function findUser_(U, user) {
  user = String(user || '').trim().toLowerCase();
  for (var i = 0; i < U.vals.length; i++) if (String(U.vals[i][0]).trim().toLowerCase() === user) return i;
  return -1;
}
/** ตรวจ token; role = ROLE_ADMIN เมื่อคำสั่งนั้นสำหรับผู้ดูแลเท่านั้น */
var CUR_ = null; // ผู้ใช้ของคำสั่งนี้ (ส่งสิทธิ์ล่าสุดกลับไปให้หน้าเว็บ)
function permsOf_(r) {
  if (r[2] === ROLE_ADMIN) return Object.keys(PERMS);
  var v = String(r[8] || '').trim();
  if (!v) return DEFAULT_PERMS.slice();
  return v.split(/[,\s]+/).filter(function (k) { return PERMS[k]; });
}
function can_(s, perm) { return s.role === ROLE_ADMIN || (s.perms || []).indexOf(perm) >= 0; }
/** ตรวจ token; need = ROLE_ADMIN (เฉพาะผู้ดูแล) หรือชื่อสิทธิ์ เช่น 'sell' */
function auth_(token, need) {
  var raw = token && cache_().get('s_' + token);
  if (!raw) throw E_('AUTH:กรุณาเข้าสู่ระบบอีกครั้ง');
  var s = JSON.parse(raw);
  if (s.t0 && Date.now() - s.t0 > SESSION_MAX_MS) { cache_().remove('s_' + token); throw E_('AUTH:ครบเวลาใช้งาน 12 ชั่วโมง กรุณาเข้าสู่ระบบใหม่'); }
  if (s.mc && need !== '_pw') throw E_('AUTH:ต้องเปลี่ยนรหัสผ่านเริ่มต้นก่อนใช้งาน');
  if (cache_().get('off_' + s.u)) { cache_().remove('s_' + token); throw E_('AUTH:บัญชีนี้ถูกปิดหรือเปลี่ยนรหัสผ่าน กรุณาเข้าสู่ระบบใหม่'); }
  var pc = cache_().get('perm_' + s.u); // ผู้ดูแลเพิ่งแก้สิทธิ์ → ใช้ทันทีไม่ต้องเข้าระบบใหม่
  if (pc) { var o = JSON.parse(pc); if (o.v !== s.pv) { s.role = o.role; s.perms = o.perms; s.pv = o.v; } }
  if (!s.perms) s.perms = s.role === ROLE_ADMIN ? Object.keys(PERMS) : DEFAULT_PERMS.slice();
  if (need === ROLE_ADMIN && s.role !== ROLE_ADMIN) throw E_('เฉพาะผู้ดูแลเท่านั้น');
  if (need && need !== ROLE_ADMIN && need !== '_pw' && !can_(s, need)) throw E_('บัญชีนี้ไม่มีสิทธิ์ ' + PERMS[need] + ' ติดต่อผู้ดูแล');
  cache_().put('s_' + token, JSON.stringify(s), SESSION_SEC);
  CUR_ = s;
  return s;
}
function meOut_() { return CUR_ ? { role: CUR_.role, perms: CUR_.role === ROLE_ADMIN ? Object.keys(PERMS) : CUR_.perms } : null; }
function login(user, pw) {
  user = String(user || '').trim().toLowerCase();
  var fk = 'fail_' + user, fails = Number(cache_().get(fk) || 0);
  if (fails >= 5) throw E_('ใส่รหัสผิดหลายครั้ง บัญชีนี้ถูกล็อก 15 นาที');
  ensureAdmin_();
  var U = readUsers_(), i = findUser_(U, user), r = i >= 0 ? U.vals[i] : null;
  var ok = r ? checkPw_(r, String(pw)) : (hash3_('x', String(pw)), false); // ใช้เวลาเท่ากันแม้ไม่มีผู้ใช้นี้
  if (!ok) {
    cache_().put(fk, String(fails + 1), 900);
    if (fails + 1 === 5) log_('ระบบ', 'ล็อกบัญชีชั่วคราว', user + ' ใส่รหัสผิด 5 ครั้ง');
    throw E_('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  if (String(r[6]).indexOf('v3$') !== 0) { r[6] = hash3_(r[5], String(pw)); U.sh.getRange(i + 2, 7).setValue(r[6]); } // อัปเกรดเป็น v3
  if (!isActive_(r[3])) throw E_('บัญชีนี้ถูกปิดการใช้งาน ติดต่อผู้ดูแล');
  cache_().remove(fk); cache_().remove('off_' + user);
  var token = Utilities.getUuid(), s = { u: user, name: String(r[1] || r[0]), role: r[2] === ROLE_ADMIN ? ROLE_ADMIN : ROLE_STAFF, perms: permsOf_(r), t0: Date.now() };
  if (WEAK_PW.indexOf(String(pw).toLowerCase()) >= 0 || String(pw).length < 8) s.mc = true; // รหัสเริ่มต้น/อ่อน → บังคับเปลี่ยนก่อน
  cache_().put('s_' + token, JSON.stringify(s), SESSION_SEC);
  U.sh.getRange(i + 2, 8).setValue(new Date());
  return { token: token, user: s.u, name: s.name, role: s.role, perms: s.perms, mustChange: !!s.mc };
}
function logout(token) { if (token) cache_().remove('s_' + token); return true; }
function changePassword(token, oldPw, newPw) {
  var s = auth_(token, '_pw');
  pwRule_(newPw);
  if (String(newPw) === String(oldPw)) throw E_('รหัสผ่านใหม่ต้องไม่ซ้ำรหัสเดิม');
  return withLock_(function () {
    var U = readUsers_(), i = findUser_(U, s.u), r = U.vals[i];
    if (i < 0 || !checkPw_(r, String(oldPw))) throw E_('รหัสผ่านเดิมไม่ถูกต้อง');
    r[5] = uid_() + uid_(); r[6] = hash3_(r[5], String(newPw));
    U.sh.getRange(i + 2, 6, 1, 2).setValues([[r[5], r[6]]]);
    log_(s.name, 'เปลี่ยนรหัสผ่านตัวเอง', s.u);
    if (s.mc) { s.mc = false; cache_().put('s_' + token, JSON.stringify(s), SESSION_SEC); }
    return true;
  });
}
function usersOut_(U) {
  return U.vals.filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    return { user: String(r[0]).trim().toLowerCase(), name: String(r[1]), role: r[2] === ROLE_ADMIN ? ROLE_ADMIN : ROLE_STAFF,
      active: isActive_(r[3]), last: r[7] instanceof Date ? r[7].getTime() : null, perms: permsOf_(r) };
  });
}
function permNames_(L) { return L.length ? L.map(function (k) { return PERMS[k]; }).join(', ') : '(ไม่มี)'; }
function listUsers(token) { auth_(token, ROLE_ADMIN); return usersOut_(readUsers_()); }
function saveUser(token, d) {
  var s = auth_(token, ROLE_ADMIN);
  return withLock_(function () {
    var U = readUsers_(), user = String(d.user || '').trim().toLowerCase(), i = findUser_(U, user);
    if (!/^[a-z0-9._-]{3,30}$/.test(user)) throw E_('ชื่อผู้ใช้ใช้ a-z 0-9 . _ - ยาว 3-30 ตัว');
    if (d.password) pwRule_(d.password);
    var role = d.role === ROLE_ADMIN ? ROLE_ADMIN : ROLE_STAFF, active = d.active !== false;
    var perms = d.perms ? [].concat(d.perms).filter(function (k) { return PERMS[k]; }) : null, pstr = perms ? (perms.join(',') || 'none') : '';
    if (user === s.u && (role !== ROLE_ADMIN || !active)) throw E_('ลดสิทธิ์หรือปิดบัญชีตัวเองไม่ได้');
    if (i < 0) {
      if (!d.password) throw E_('กรุณาตั้งรหัสผ่านสำหรับบัญชีใหม่');
      var salt = uid_() + uid_();
      U.sh.getRange(U.sh.getLastRow() + 1, 1, 1, U_HEAD.length).setValues([safeRow_([user, String(d.name || user), role, active, '', salt, hash3_(salt, String(d.password)), '', pstr])]);
      log_(s.name, 'เพิ่มบัญชี', user + ' (' + role + ')' + (perms && role !== ROLE_ADMIN ? ' · สิทธิ์: ' + permNames_(perms) : ''));
    } else {
      var r = U.vals[i], before = r[2] + '/' + (isActive_(r[3]) ? 'ใช้งาน' : 'ปิด'), oldP = permsOf_(r).join(',');
      r[1] = d.name || r[1]; r[2] = role; r[3] = active;
      if (d.password) { r[5] = uid_() + uid_(); r[6] = hash3_(r[5], String(d.password)); }
      if (perms) r[8] = pstr;
      U.sh.getRange(i + 2, 2, 1, 6).setValues([safeRow_([r[1], r[2], r[3], '', r[5], r[6]])]);
      U.sh.getRange(i + 2, 9).setValue(r[8] || '');
      var newP = permsOf_(r);
      cache_().put('perm_' + user, JSON.stringify({ v: uid_(), role: role, perms: newP }), SESSION_SEC);
      if (!active || d.password) cache_().put('off_' + user, '1', SESSION_SEC); // เตะออกจากทุกเครื่อง
      log_(s.name, 'แก้บัญชี', user + ': ' + before + ' → ' + role + '/' + (active ? 'ใช้งาน' : 'ปิด') + (d.password ? ' · รีเซ็ตรหัสผ่าน' : '')
        + (role !== ROLE_ADMIN && newP.join(',') !== oldP ? ' · สิทธิ์: ' + permNames_(newP) : ''));
    }
    return usersOut_(readUsers_());
  });
}

/* ---------- ข้อมูลประกอบ (aux) ---------- */
function suppliersOut_() {
  return rows_(SH_SUP, SUP_HEAD).vals.filter(function (r) { return r[0] && String(r[1]).trim(); }).map(function (r) {
    return { id: String(r[0]), name: String(r[1]), contact: String(r[2]), phone: String(r[3]), line: String(r[4]), note: String(r[5]), active: bool_(r[6], true) };
  });
}
function posOut_() {
  var v = rows_(SH_PO, PO_HEAD).vals.filter(function (r) { return r[0]; });
  return v.slice(-150).map(function (r) {
    var lines = []; try { lines = JSON.parse(r[5] || '[]'); } catch (e) {}
    return { no: String(r[0]), t: r[1] instanceof Date ? r[1].getTime() : 0, sup: String(r[2]), supName: String(r[3]), status: String(r[4]),
      lines: lines, total: num_(r[6]), by: String(r[7]), recvAt: r[8] instanceof Date ? r[8].getTime() : null, recvBy: String(r[9]), note: String(r[10]) };
  });
}
function closingsOut_() {
  return rows_(SH_CL, CL_HEAD).vals.filter(function (r) { return r[0]; }).slice(-120).map(function (r) {
    return { date: r[0] instanceof Date ? dkey_(r[0]) : String(r[0]), t: r[1] instanceof Date ? r[1].getTime() : 0, by: String(r[2]), float: num_(r[3]),
      cash: num_(r[4]), refund: num_(r[5]), expected: num_(r[6]), counted: num_(r[7]), diff: num_(r[8]), transfer: num_(r[9]), note: String(r[10]) };
  });
}
function imagesOut_() {
  var o = {};
  rows_(SH_IMG, IMG_HEAD).vals.forEach(function (r) { if (r[0] && r[1]) o[String(r[0])] = String(r[1]); });
  return o;
}
var AUX_FN = { suppliers: suppliersOut_, pos: posOut_, closings: closingsOut_, images: imagesOut_ };

/* ---------- ส่งข้อมูลให้หน้าเว็บ ---------- */
function syncRow_(sync) { return typeof sync === 'number' ? sync : (sync && sync.row) || 0; }
function publicSettings_(s) { return { emails: s.emails, notifyLow: s.notifyLow, notifySummary: s.notifySummary, backup: s.backup, float: s.float, lastBackup: s.lastBackup }; }
function auxOut_(vers) {
  var aux = {}, nv = {};
  AUX.forEach(function (k) {
    var v = props_().getProperty('v_' + k) || '0';
    if (!vers || vers[k] !== v) { aux[k] = AUX_FN[k](); nv[k] = v; }
  });
  return { aux: aux, vers: nv };
}
function getData(token, days) { auth_(token); return data_(days); }
function data_(days) {
  var since = Date.now() - (days || DEFAULT_DAYS) * 864e5;
  var s = readSettings_(), last = mSheet_().getLastRow(), first = firstRowSince_(since), A = auxOut_(null);
  return { me: meOut_(), full: true, shop: s.shop, cats: s.cats, settings: publicSettings_(s), url: ss_().getUrl(), products: productsOut_(loadP_()),
    moves: readMoves_(first, last), firstRow: first, lastRow: Math.max(1, last), since: since, aux: A.aux, vers: A.vers };
}
function getUpdates(token, sync) { auth_(token); return updates_(sync); }
function updates_(sync) {
  var fromRow = syncRow_(sync), last = mSheet_().getLastRow();
  if (!fromRow || last < fromRow) return data_(); // มีคนลบแถวในชีต → โหลดใหม่ทั้งหมด
  var s = readSettings_(), A = auxOut_(sync && sync.vers ? sync.vers : {});
  return { me: meOut_(), shop: s.shop, cats: s.cats, settings: publicSettings_(s), products: productsOut_(loadP_()),
    moves: readMoves_(fromRow + 1, last), lastRow: Math.max(1, last), aux: A.aux, vers: A.vers };
}
function loadOlder(token, since, beforeRow) {
  auth_(token);
  var first = firstRowSince_(since);
  return { moves: first < beforeRow ? readMoves_(first, beforeRow - 1) : [], firstRow: Math.min(first, beforeRow), since: since };
}

/* ---------- ขาย / คืนสินค้า ---------- */
/** ขาย: cart = [{id, qty}], opt = {pay} */
function sell(token, opId, cart, opt, bill, sync) {
  var s = auth_(token, 'sell');
  opt = opt && typeof opt === 'object' ? opt : {};
  var pay = opt.pay === PAY_TRANSFER ? PAY_TRANSFER : 'เงินสด';
  return withLock_(function () {
    if (done_(opId)) return updates_(sync);
    var P = loadP_(), rows = [];
    cart.forEach(function (it, k) {
      var i = P.byId[it.id], q = Math.floor(num_(it.qty));
      if (i === undefined) throw E_('ไม่พบสินค้า (อาจถูกลบไปแล้ว)');
      if (q <= 0) return;
      var r = P.vals[i];
      if (q > num_(r[7])) throw E_('สต็อก “' + r[2] + '” เหลือเพียง ' + num_(r[7]) + ' ไม่พอขาย ' + q);
      r[7] = num_(r[7]) - q;
      rows.push(moveRow_(opId + '-' + k, 'sale', r, q, { bill: bill, who: s.name, pay: pay }));
    });
    if (!rows.length) throw E_('ไม่มีสินค้าในตะกร้า');
    writeNums_(P);
    appendMoves_(rows);
    return updates_(sync);
  });
}
/** สินค้าในบิลที่ยังคืนได้: {pid: {sold, returned, price, cost, name, code}} */
function billState_(bill) {
  var ms = movesSince_(Date.now() - 400 * 864e5).filter(function (m) { return m.bill === bill; });
  if (!ms.length) throw E_('ไม่พบบิล ' + bill);
  var st = {}, head = null;
  ms.forEach(function (m) {
    var o = st[m.pid] || (st[m.pid] = { sold: 0, returned: 0, price: m.price, cost: m.cost, name: m.pname, code: m.code });
    if (m.type === 'sale') { o.sold += m.qty; o.price = m.price; o.cost = m.cost; head = head || m; }
    if (m.type === 'ret') o.returned += m.qty;
  });
  if (!head) throw E_('บิลนี้ไม่ใช่บิลขาย');
  return { st: st, head: head };
}
/** คืนสินค้า/ยกเลิกบิล: items = [{pid, qty}] หรือ null = คืนทั้งหมดที่เหลือ */
function refund(token, opId, bill, items, reason, sync) {
  var s = auth_(token);
  return withLock_(function () {
    if (done_(opId)) return updates_(sync);
    var B = billState_(bill), h = B.head;
    if (!can_(s, 'refundAll')) {
      if (!can_(s, 'refund')) throw E_('บัญชีนี้ไม่มีสิทธิ์คืนสินค้า ติดต่อผู้ดูแล');
      if (h.who !== s.name || dkey_(h.t) !== dkey_(new Date())) throw E_('คืนได้เฉพาะบิลของตัวเองในวันนี้ บิลอื่นให้ผู้ดูแลทำ');
    }
    if (!String(reason || '').trim()) throw E_('กรุณาระบุเหตุผล');
    var list = items || Object.keys(B.st).map(function (pid) { return { pid: pid, qty: B.st[pid].sold - B.st[pid].returned }; });
    var P = loadP_(), rows = [], total = 0;
    list.forEach(function (it, k) {
      var o = B.st[it.pid], q = Math.floor(num_(it.qty));
      if (!o || q <= 0) return;
      if (q > o.sold - o.returned) throw E_('“' + o.name + '” คืนได้ไม่เกิน ' + (o.sold - o.returned));
      var i = P.byId[it.pid], r = i === undefined ? [it.pid, o.code, o.name, '', '', o.cost, o.price, 0] : P.vals[i];
      if (i !== undefined) r[7] = num_(r[7]) + q;
      rows.push(moveRow_(opId + '-' + k, 'ret', r, q, { price: o.price, cost: o.cost, bill: bill, note: reason, who: s.name, pay: h.pay }));
      total += q * o.price;
    });
    if (!rows.length) throw E_('ไม่มีรายการให้คืน');
    writeNums_(P);
    appendMoves_(rows);
    log_(s.name, items ? 'คืนสินค้า' : 'ยกเลิกบิล', bill + ' · ' + money_(total) + ' · ' + reason);
    return updates_(sync);
  });
}

/** รับเข้า (in) / ตัดออก (adj) / ตรวจนับ (count) */
function moveStock(token, opId, pid, type, qty, cost, note, who, sync) {
  var s = auth_(token);
  return withLock_(function () {
    if (done_(opId)) return updates_(sync);
    if (['in', 'adj', 'count'].indexOf(type) < 0) throw E_('ประเภทไม่ถูกต้อง');
    if (!can_(s, type === 'in' ? 'recv' : 'adj')) throw E_('บัญชีนี้ไม่มีสิทธิ์ ' + PERMS[type === 'in' ? 'recv' : 'adj'] + ' ติดต่อผู้ดูแล');
    var P = loadP_(), i = P.byId[pid];
    if (i === undefined) throw E_('ไม่พบสินค้า');
    var r = P.vals[i], stock = num_(r[7]);
    qty = Math.floor(num_(qty));
    if (qty < 0 || (type !== 'count' && qty === 0)) throw E_('จำนวนไม่ถูกต้อง');
    if (type === 'adj' && qty > stock) throw E_('ตัดออกได้ไม่เกิน ' + stock);
    var delta = type === 'in' ? qty : type === 'adj' ? -qty : qty - stock;
    if (type === 'in' && cost !== '' && cost != null) r[5] = avgCost_(stock, num_(r[5]), qty, num_(cost));
    r[7] = stock + delta;
    writeNums_(P);
    appendMoves_([moveRow_(opId + '-0', type, r, delta, { note: note, who: s.name })]);
    return updates_(sync);
  });
}
function avgCost_(stock, cost, qty, c) { return stock > 0 ? Math.round((stock * cost + qty * c) / (stock + qty) * 100) / 100 : c; }

/* ---------- จัดการสินค้า (ผู้ดูแล) ---------- */
function saveProduct(token, opId, d, sync) {
  var s = auth_(token, 'prod');
  return withLock_(function () {
    if (done_(opId)) return updates_(sync);
    var P = loadP_();
    d.code = String(d.code || '').trim(); d.name = String(d.name || '').trim();
    if (!d.name) throw E_('กรุณากรอกชื่อสินค้า');
    var dup = P.vals.some(function (r) { return d.code && String(r[1]) === d.code && String(r[0]) !== d.id; });
    if (dup) throw E_('รหัสสินค้า ' + d.code + ' มีอยู่แล้ว');
    var i = d.id ? P.byId[d.id] : undefined;
    if (i !== undefined) {
      var r = P.vals[i], labels = ['รหัส', 'ชื่อ', 'หมวด', 'หน่วย', 'ทุน', 'ราคาขาย'], nv = [d.code, d.name, d.cat, d.unit, num_(d.cost), num_(d.price)], diffs = [];
      nv.forEach(function (v, k) { if (String(r[k + 1]) !== String(v)) diffs.push(labels[k] + ' ' + r[k + 1] + ' → ' + v); });
      if (num_(r[8]) !== num_(d.min)) diffs.push('จุดสั่งซื้อ ' + r[8] + ' → ' + num_(d.min));
      if (String(r[10] || '') !== String(d.sup || '')) diffs.push('ผู้จำหน่าย');
      P.sh.getRange(i + 2, 2, 1, 6).setValues([safeRow_(nv)]);
      P.sh.getRange(i + 2, 9).setValue(num_(d.min));
      P.sh.getRange(i + 2, 11).setValue(d.sup || '');
      if (diffs.length) log_(s.name, 'แก้สินค้า', d.name + ': ' + diffs.join(' · '));
    } else {
      var stock = Math.floor(num_(d.stock));
      var nr = [d.id || uid_(), d.code, d.name, d.cat, d.unit, num_(d.cost), num_(d.price), stock, num_(d.min), true, d.sup || ''];
      P.sh.getRange(P.sh.getLastRow() + 1, 1, 1, nr.length).setValues([safeRow_(nr)]);
      if (stock) appendMoves_([moveRow_(opId + '-0', 'count', nr, stock, { note: 'ยอดยกมา', who: s.name })]);
      log_(s.name, 'เพิ่มสินค้า', d.name + ' · ราคา ' + money_(d.price) + ' · ยอดยกมา ' + stock);
    }
    return updates_(sync);
  });
}
function setProductActive(token, opId, pid, active, sync) {
  var s = auth_(token, 'prod');
  return withLock_(function () {
    var P = loadP_(), i = P.byId[pid];
    if (i === undefined) throw E_('ไม่พบสินค้า');
    P.sh.getRange(i + 2, 10).setValue(!!active);
    log_(s.name, active ? 'เปิดใช้งานสินค้า' : 'ปิดใช้งานสินค้า', P.vals[i][2]);
    return updates_(sync);
  });
}
function saveImage(token, opId, pid, data, sync) {
  var s = auth_(token, 'prod');
  if (data && !/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(String(data))) throw E_('ไฟล์รูปไม่ถูกต้อง');
  if (data && String(data).length > 48000) throw E_('รูปใหญ่เกินไป');
  return withLock_(function () {
    var R = rows_(SH_IMG, IMG_HEAD), idx = -1;
    R.vals.forEach(function (r, i) { if (String(r[0]) === pid) idx = i; });
    if (idx >= 0) R.sh.getRange(idx + 2, 2).setValue(data || '');
    else if (data) R.sh.appendRow([pid, data]);
    bump_('images');
    log_(s.name, data ? 'เปลี่ยนรูปสินค้า' : 'ลบรูปสินค้า', pid);
    return updates_(sync);
  });
}
function saveSettings(token, opId, patch, sync) {
  var s = auth_(token, ROLE_ADMIN);
  return withLock_(function () {
    var allowed = ['shop', 'cats', 'emails', 'notifyLow', 'notifySummary', 'backup', 'float'], changed = [];
    Object.keys(patch || {}).forEach(function (k) {
      if (allowed.indexOf(k) < 0) return;
      var v = patch[k];
      if (k === 'cats') v = [].concat(v).join(', ');
      if (k === 'float') v = num_(v);
      setSetting_(k, v); changed.push(SKEYS[k] + ' = ' + v);
    });
    if (changed.length) log_(s.name, 'แก้ตั้งค่า', changed.join(' · '));
    return updates_(sync);
  });
}

/* ---------- ผู้จำหน่าย / ใบสั่งซื้อ ---------- */
function saveSupplier(token, opId, d, sync) {
  var s = auth_(token, 'poManage');
  if (!String(d.name || '').trim()) throw E_('กรุณากรอกชื่อผู้จำหน่าย');
  return withLock_(function () {
    var R = rows_(SH_SUP, SUP_HEAD), idx = -1, row = safeRow_([String(d.id || uid_()), String(d.name), String(d.contact || ''), String(d.phone || ''), String(d.line || ''), String(d.note || ''), d.active !== false]);
    R.vals.forEach(function (r, i) { if (d.id && String(r[0]) === d.id) idx = i; });
    if (idx >= 0) R.sh.getRange(idx + 2, 1, 1, row.length).setValues([row]); else R.sh.appendRow(row);
    bump_('suppliers');
    log_(s.name, idx >= 0 ? 'แก้ผู้จำหน่าย' : 'เพิ่มผู้จำหน่าย', d.name);
    return updates_(sync);
  });
}
function findPO_(R, no) { for (var i = 0; i < R.vals.length; i++) if (String(R.vals[i][0]) === no) return i; return -1; }
/** po = {no?, sup, lines:[{pid, qty, cost}], note} → สถานะ "สั่งแล้ว" */
function savePO(token, opId, po, sync) {
  var s = auth_(token, 'poManage');
  return withLock_(function () {
    var R = rows_(SH_PO, PO_HEAD), P = loadP_(), sups = suppliersOut_(), sup = sups.filter(function (x) { return x.id === po.sup; })[0];
    var lines = (po.lines || []).map(function (l) {
      var i = P.byId[l.pid]; if (i === undefined || num_(l.qty) <= 0) return null;
      return { pid: l.pid, name: P.vals[i][2], unit: P.vals[i][4], qty: Math.floor(num_(l.qty)), cost: num_(l.cost) };
    }).filter(Boolean);
    if (!lines.length) throw E_('ใบสั่งซื้อต้องมีอย่างน้อย 1 รายการ');
    var total = lines.reduce(function (a, l) { return a + l.qty * l.cost; }, 0);
    var idx = po.no ? findPO_(R, po.no) : -1;
    if (idx >= 0 && String(R.vals[idx][4]) !== 'สั่งแล้ว') throw E_('ใบสั่งซื้อนี้รับของหรือยกเลิกแล้ว แก้ไม่ได้');
    var no = idx >= 0 ? po.no : 'PO' + fmt_(new Date(), 'yyMMdd') + '-' + ('00' + (R.vals.filter(function (r) { return String(r[0]).indexOf('PO' + fmt_(new Date(), 'yyMMdd')) === 0; }).length + 1)).slice(-3);
    var row = [no, idx >= 0 ? R.vals[idx][1] : new Date(), po.sup || '', sup ? sup.name : '', 'สั่งแล้ว', JSON.stringify(lines), total, idx >= 0 ? R.vals[idx][7] : s.name, '', '', String(po.note || '')].map(safe_);
    if (idx >= 0) R.sh.getRange(idx + 2, 1, 1, row.length).setValues([row]); else R.sh.appendRow(row);
    bump_('pos');
    log_(s.name, idx >= 0 ? 'แก้ใบสั่งซื้อ' : 'สร้างใบสั่งซื้อ', no + ' · ' + (sup ? sup.name : '-') + ' · ' + money_(total));
    return updates_(sync);
  });
}
function cancelPO(token, opId, no, sync) {
  var s = auth_(token, 'poManage');
  return withLock_(function () {
    var R = rows_(SH_PO, PO_HEAD), i = findPO_(R, no);
    if (i < 0) throw E_('ไม่พบใบสั่งซื้อ');
    if (String(R.vals[i][4]) !== 'สั่งแล้ว') throw E_('ใบสั่งซื้อนี้ปิดไปแล้ว');
    R.sh.getRange(i + 2, 5).setValue('ยกเลิก');
    bump_('pos'); log_(s.name, 'ยกเลิกใบสั่งซื้อ', no);
    return updates_(sync);
  });
}
/** รับของตามใบสั่งซื้อ: lines = [{pid, qty, cost}] จำนวนที่รับจริง */
function receivePO(token, opId, no, lines, sync) {
  var s = auth_(token, 'po');
  return withLock_(function () {
    if (done_(opId)) return updates_(sync);
    var R = rows_(SH_PO, PO_HEAD), pi = findPO_(R, no);
    if (pi < 0) throw E_('ไม่พบใบสั่งซื้อ');
    if (String(R.vals[pi][4]) !== 'สั่งแล้ว') throw E_('ใบสั่งซื้อนี้รับของหรือยกเลิกไปแล้ว');
    var P = loadP_(), rows = [];
    lines.forEach(function (l, k) {
      var i = P.byId[l.pid], q = Math.floor(num_(l.qty));
      if (i === undefined || q <= 0) return;
      var r = P.vals[i], stock = num_(r[7]);
      if (l.cost !== '' && l.cost != null) r[5] = avgCost_(stock, num_(r[5]), q, num_(l.cost));
      r[7] = stock + q;
      rows.push(moveRow_(opId + '-' + k, 'in', r, q, { note: 'รับตามใบสั่งซื้อ ' + no, who: s.name }));
    });
    if (!rows.length) throw E_('ยังไม่ได้ใส่จำนวนที่รับ');
    writeNums_(P);
    appendMoves_(rows);
    R.sh.getRange(pi + 2, 5, 1, 1).setValue('รับแล้ว');
    R.sh.getRange(pi + 2, 9, 1, 2).setValues([[new Date(), s.name]]);
    bump_('pos'); log_(s.name, 'รับของตามใบสั่งซื้อ', no + ' · ' + rows.length + ' รายการ');
    return updates_(sync);
  });
}

/* ---------- ปิดยอดเงินสดประจำวัน ---------- */
function daySummary_(key) {
  var start = new Date(key + 'T00:00:00').getTime() - 864e5;
  var ms = movesSince_(start).filter(function (m) { return dkey_(m.t) === key; });
  var o = { cash: 0, refund: 0, transfer: 0, sales: 0, cost: 0, bills: {}, items: {}, ret: 0 };
  ms.forEach(function (m) {
    var v = m.qty * m.price;
    if (m.type === 'sale') {
      o.sales += v; o.cost += m.qty * m.cost; o.bills[m.bill] = 1;
      if (m.pay === PAY_TRANSFER) o.transfer += v; else o.cash += v;
      var it = o.items[m.pid] || (o.items[m.pid] = { name: m.pname, q: 0, v: 0 }); it.q += m.qty; it.v += v;
    }
    if (m.type === 'ret') {
      o.sales -= v; o.cost -= m.qty * m.cost; o.ret += v;
      if (m.pay === PAY_TRANSFER) o.transfer -= v; else o.refund += v;
      var it2 = o.items[m.pid] || (o.items[m.pid] = { name: m.pname, q: 0, v: 0 }); it2.q -= m.qty; it2.v -= v;
    }
  });
  o.nbills = Object.keys(o.bills).length;
  return o;
}
function closeCash(token, opId, key, counted, float, note, sync) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(key))) throw E_('วันที่ไม่ถูกต้อง');
  var s = auth_(token, 'cash');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(key))) throw E_('วันที่ไม่ถูกต้อง');
  return withLock_(function () {
    var R = rows_(SH_CL, CL_HEAD), idx = -1, d = daySummary_(key);
    R.vals.forEach(function (r, i) { if ((r[0] instanceof Date ? dkey_(r[0]) : String(r[0])) === key) idx = i; });
    if (idx >= 0 && s.role !== ROLE_ADMIN) throw E_('วันนี้ปิดยอดไปแล้ว ให้ผู้ดูแลแก้ไข');
    var expected = num_(float) + d.cash - d.refund, diff = num_(counted) - expected;
    var row = ["'" + key, new Date(), s.name, num_(float), d.cash, d.refund, expected, num_(counted), diff, d.transfer, safe_(String(note || ''))];
    if (idx >= 0) R.sh.getRange(idx + 2, 1, 1, row.length).setValues([row]); else R.sh.appendRow(row);
    bump_('closings');
    log_(s.name, idx >= 0 ? 'แก้การปิดยอด' : 'ปิดยอดเงินสด', key + ' · ควรมี ' + money_(expected) + ' · นับได้ ' + money_(counted) + ' · ผลต่าง ' + money_(diff));
    return updates_(sync);
  });
}

/* ---------- บันทึกการแก้ไข / ข้อมูลผู้ดูแล ---------- */
function adminInfo(token) {
  auth_(token, ROLE_ADMIN);
  var L = rows_(SH_LOG, LOG_HEAD).vals.slice(-300).reverse().map(function (r) {
    return { t: r[0] instanceof Date ? r[0].getTime() : 0, who: String(r[1]), what: String(r[2]), detail: String(r[3]) };
  });
  var trig = [];
  try { trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) {}
  return { users: usersOut_(readUsers_()), audit: L, triggers: trig, permDefs: PERMS, defaultPerms: DEFAULT_PERMS };
}

/* ---------- ตรวจสุขภาพระบบ (ผู้ดูแล) ---------- */
function healthCheck(token) {
  auth_(token, ROLE_ADMIN);
  var out = [], add = function (st, name, detail, fix) { out.push({ st: st, name: name, detail: detail || '', fix: fix || '' }); };
  var tz = ''; try { tz = Session.getScriptTimeZone(); } catch (e) {}
  add(tz === 'Asia/Bangkok' ? 'ok' : 'bad', 'เขตเวลาโปรเจ็กต์', tz || 'อ่านไม่ได้', 'วางไฟล์ appsscript.json จากชุดติดตั้ง หรือ ตั้งค่าโปรเจ็กต์ > เขตเวลา = (GMT+07:00) Bangkok');
  var stz = ''; try { stz = ss_().getSpreadsheetTimeZone(); } catch (e) {}
  add(stz === 'Asia/Bangkok' ? 'ok' : 'warn', 'เขตเวลาของไฟล์ชีต', stz || 'อ่านไม่ได้', 'ในชีต: ไฟล์ > การตั้งค่า > เขตเวลา = (GMT+07:00) Bangkok');
  var miss = [[SH_P, P_HEAD], [SH_M, M_HEAD], [SH_S, S_HEAD], [SH_U, U_HEAD], [SH_SUP, SUP_HEAD], [SH_PO, PO_HEAD], [SH_CL, CL_HEAD], [SH_LOG, LOG_HEAD], [SH_IMG, IMG_HEAD]]
    .filter(function (x) { return !ss_().getSheetByName(x[0]); }).map(function (x) { return x[0]; });
  add(miss.length ? 'bad' : 'ok', 'แท็บข้อมูลครบ', miss.length ? 'ขาด: ' + miss.join(', ') : 'ครบ 9 แท็บ', 'เปิด Apps Script แล้วรันฟังก์ชัน setup อีกครั้ง');
  var trig = []; try { trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) {}
  var lack = JOBS.filter(function (j) { return trig.indexOf(j) < 0; });
  add(lack.length ? 'warn' : 'ok', 'งานอัตโนมัติ (อีเมล/สำรองข้อมูล)', lack.length ? 'ยังไม่ได้ตั้ง: ' + lack.join(', ') : 'ตั้งครบ 3 งาน', 'หน้าตั้งค่า > แจ้งเตือนทางอีเมล > ตั้งเวลางานอัตโนมัติ');
  var U = readUsers_(), admins = U.vals.filter(function (r) { return r[2] === ROLE_ADMIN && isActive_(r[3]); });
  var old = U.vals.filter(function (r) { return String(r[6]).indexOf('v3$') !== 0; }).map(function (r) { return r[0]; });
  add(admins.length >= 1 && admins.length <= 3 ? 'ok' : 'warn', 'จำนวนผู้ดูแล', admins.length + ' คน', 'ควรมีผู้ดูแล 1–3 คน ที่เหลือเป็นเจ้าหน้าที่');
  add(old.length ? 'warn' : 'ok', 'รหัสผ่านแบบเข้ารหัสใหม่', old.length ? 'ยังเป็นแบบเก่า: ' + old.join(', ') : 'ทุกบัญชีใช้แบบใหม่แล้ว', 'ให้บัญชีเหล่านี้เข้าสู่ระบบ 1 ครั้ง ระบบจะอัปเกรดเอง');
  try {
    var f = DriveApp.getFileById(ss_().getId()), acc = String(f.getSharingAccess()), eds = f.getEditors().length;
    add(/ANYONE/.test(acc) ? 'bad' : 'ok', 'การแชร์ไฟล์ชีต', /ANYONE/.test(acc) ? 'ใครมีลิงก์ก็เปิดได้' : 'จำกัดเฉพาะคนที่ระบุ', 'ในชีต: ปุ่มแชร์ > สิทธิ์เข้าถึงทั่วไป = จำกัด');
    add(eds <= 3 ? 'ok' : 'warn', 'ผู้แก้ไขไฟล์ชีตได้', 'เจ้าของ + ' + eds + ' คน', 'คนที่แก้ชีตได้ข้ามระบบสิทธิ์ทั้งหมด ควรเหลือเฉพาะผู้ดูแล');
  } catch (e) { add('warn', 'การแชร์ไฟล์ชีต', 'ตรวจไม่ได้ (' + String(e.message || e).slice(0, 80) + ')', 'ตรวจเองที่ปุ่มแชร์ของชีต: ควรเป็น "จำกัด"'); }
  var lb = readSettings_().lastBackup, age = lb ? (Date.now() - (typeof lb === 'number' ? lb : new Date(lb).getTime())) / 864e5 : null;
  add(age !== null && age <= 8 ? 'ok' : 'warn', 'สำรองข้อมูลล่าสุด', age === null || isNaN(age) ? 'ยังไม่เคยสำรอง' : Math.floor(age) + ' วันที่แล้ว', 'หน้าตั้งค่า > สำรองข้อมูลตอนนี้ และเปิดสำรองอัตโนมัติ');
  var q = -1; try { q = MailApp.getRemainingDailyQuota(); } catch (e) {}
  add(q > 10 ? 'ok' : 'warn', 'โควตาอีเมลวันนี้', q < 0 ? 'ตรวจไม่ได้' : 'ส่งได้อีก ' + q + ' ฉบับ', 'โควตารีเซ็ตทุกวัน (Gmail ทั่วไป 100 ฉบับ/วัน)');
  var rows = mSheet_().getLastRow(), cells = 0; ss_().getSheets().forEach(function (sh) { cells += sh.getMaxRows() * sh.getMaxColumns(); });
  var pct = Math.round(cells / 1e7 * 100);
  add(pct < 70 ? 'ok' : pct < 90 ? 'warn' : 'bad', 'พื้นที่ Google Sheets', 'ใช้ ' + pct + '% ของ 10 ล้านเซลล์ · ความเคลื่อนไหว ' + rows + ' แถว', 'ถ้าเกิน 70% ให้ย้ายข้อมูลปีเก่าไปไฟล์ใหม่');
  var url = ''; try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  add(url ? 'ok' : 'warn', 'การเผยแพร่เว็บแอป', url ? 'เผยแพร่แล้ว' : 'ยังไม่ได้ Deploy', 'Deploy > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป');
  var score = Math.round(out.reduce(function (a, x) { return a + (x.st === 'ok' ? 1 : x.st === 'warn' ? 0.5 : 0); }, 0) / out.length * 100);
  return { score: score, items: out, t: Date.now() };
}

/* ---------- อีเมล / งานอัตโนมัติ / สำรองข้อมูล ---------- */
var JOBS = ['morningJob', 'eveningJob', 'backupJob'];
function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (JOBS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('morningJob').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('eveningJob').timeBased().everyDays(1).atHour(18).create();
  ScriptApp.newTrigger('backupJob').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).create();
}
function installTriggers(token) {
  var s = auth_(token, ROLE_ADMIN);
  installTriggers_(); log_(s.name, 'ตั้งเวลางานอัตโนมัติ', JOBS.join(', '));
  return adminInfo(token);
}
function emailWrap_(title, body) {
  var shop = esc_(readSettings_().shop);
  return '<div style="font-family:Tahoma,sans-serif;max-width:560px;margin:auto;color:#0b2633">'
    + '<div style="background:#0b7fb8;color:#fff;padding:16px 20px;border-radius:12px 12px 0 0"><div style="font-size:13px;opacity:.85">' + shop + '</div>'
    + '<div style="font-size:20px;font-weight:bold">' + title + '</div></div>'
    + '<div style="border:1px solid #d9e6ec;border-top:0;padding:16px 20px;border-radius:0 0 12px 12px">' + body + '</div>'
    + '<div style="font-size:12px;color:#6a8592;text-align:center;margin-top:10px">ส่งอัตโนมัติจากระบบสต็อกสหกรณ์ · ปิดการแจ้งเตือนได้ที่หน้าตั้งค่า</div></div>';
}
function tableHtml_(head, rows, alignRight) {
  return '<table style="width:100%;border-collapse:collapse;font-size:14px"><tr>' + head.map(function (h, i) {
    return '<th style="text-align:' + (i && alignRight ? 'right' : 'left') + ';padding:6px;border-bottom:2px solid #d9e6ec;color:#6a8592">' + h + '</th>';
  }).join('') + '</tr>' + rows.map(function (r) {
    return '<tr>' + r.map(function (c, i) { return '<td style="padding:6px;border-bottom:1px solid #eef3f5;text-align:' + (i && alignRight ? 'right' : 'left') + '">' + c + '</td>'; }).join('') + '</tr>';
  }).join('') + '</table>';
}
function buildMorning_() {
  var low = productsOut_(loadP_()).filter(function (p) { return p.active && p.stock <= p.min; }).sort(function (a, b) { return a.stock - b.stock; });
  var pend = posOut_().filter(function (p) { return p.status === 'สั่งแล้ว'; });
  var sups = {}; suppliersOut_().forEach(function (x) { sups[x.id] = x.name; });
  var body = low.length
    ? '<p>มีสินค้าถึงจุดสั่งซื้อ <b>' + low.length + ' รายการ</b></p>' + tableHtml_(['สินค้า', 'คงเหลือ', 'จุดสั่งซื้อ', 'ผู้จำหน่าย'], low.map(function (p) {
      return [esc_(p.name), (p.stock <= 0 ? '<b style="color:#c43a32">หมด</b>' : p.stock + ' ' + esc_(p.unit)), p.min, esc_(sups[p.sup] || '-')]; }), true)
    : '<p>สต็อกทุกรายการเพียงพอ ไม่มีสินค้าที่ต้องสั่งเพิ่ม</p>';
  if (pend.length) body += '<p style="margin-top:16px">ใบสั่งซื้อที่รอรับของ ' + pend.length + ' ใบ: ' + pend.map(function (p) { return esc_(p.no + ' (' + (p.supName || '-') + ')'); }).join(', ') + '</p>';
  return { count: low.length, subject: (low.length ? '⚠ สินค้าใกล้หมด ' + low.length + ' รายการ' : 'สต็อกเพียงพอ') + ' · ' + fmt_(new Date(), 'd/M/yyyy'),
    html: emailWrap_('แจ้งเตือนสินค้าใกล้หมด', body) };
}
function buildSummary_(key) {
  var d = daySummary_(key), closed = closingsOut_().filter(function (c) { return c.date === key; })[0];
  var top = Object.keys(d.items).map(function (k) { return d.items[k]; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 8);
  var body = tableHtml_(['', ''], [['ยอดขายสุทธิ', '<b>' + money_(d.sales) + '</b>'], ['กำไรขั้นต้น', money_(d.sales - d.cost)], ['จำนวนบิล', d.nbills],
    ['เงินสด (หักคืน)', money_(d.cash - d.refund)], ['โอน', money_(d.transfer)], ['คืนสินค้า/ยกเลิก', money_(d.ret)]], true)
    + (top.length ? '<p style="margin-top:16px"><b>สินค้าขายดี</b></p>' + tableHtml_(['สินค้า', 'จำนวน', 'ยอด'], top.map(function (i) { return [esc_(i.name), i.q, money_(i.v)]; }), true) : '')
    + '<p style="margin-top:16px">' + (closed ? 'ปิดยอดแล้วโดย ' + esc_(closed.by) + ' · นับได้ ' + money_(closed.counted) + ' · ผลต่าง <b style="color:' + (Math.abs(closed.diff) < 0.01 ? '#14804a' : '#c43a32') + '">' + money_(closed.diff) + '</b>'
      : '<b style="color:#a86400">ยังไม่ได้ปิดยอดเงินสดของวันนี้</b>') + '</p>';
  var low = productsOut_(loadP_()).filter(function (p) { return p.active && p.stock <= p.min; }).length;
  if (low) body += '<p>สินค้าใกล้หมด ' + low + ' รายการ</p>';
  return { subject: 'สรุปยอดขาย ' + fmt_(new Date(key + 'T12:00:00'), 'd/M/yyyy') + ' · ' + money_(d.sales), html: emailWrap_('สรุปยอดขายประจำวัน', body), sales: d.sales };
}
function send_(subject, html) {
  var to = emailList_();
  if (!to.length) return false;
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html, name: readSettings_().shop });
  return true;
}
function morningJob() { if (!readSettings_().notifyLow) return; var m = buildMorning_(); if (m.count) send_(m.subject, m.html); }
function eveningJob() { if (!readSettings_().notifySummary) return; var key = dkey_(new Date()), s = buildSummary_(key); if (s.sales || daySummary_(key).nbills) send_(s.subject, s.html); }
function sendTestEmail(token) {
  var s = auth_(token, ROLE_ADMIN), to = emailList_();
  if (!to.length) throw E_('ยังไม่ได้ใส่อีเมลผู้รับในหน้าตั้งค่า');
  var m = buildMorning_(), d = buildSummary_(dkey_(new Date()));
  send_('[ทดสอบ] ' + m.subject, m.html); send_('[ทดสอบ] ' + d.subject, d.html);
  log_(s.name, 'ทดสอบส่งอีเมล', to.join(', '));
  return { to: to, html: m.html + '<div style="height:24px"></div>' + d.html, remaining: (function () { try { return MailApp.getRemainingDailyQuota(); } catch (e) { return null; } })() };
}
function backupJob(force) {
  if (force !== true && !readSettings_().backup) return null;
  var file = DriveApp.getFileById(ss_().getId()), name = 'สำรอง ' + ss_().getName() + ' ' + fmt_(new Date(), 'yyyy-MM-dd HHmm');
  var it = DriveApp.getFoldersByName('สำรองข้อมูลสต็อกสหกรณ์'), folder = it.hasNext() ? it.next() : DriveApp.createFolder('สำรองข้อมูลสต็อกสหกรณ์');
  var copy = file.makeCopy(name, folder), all = [], fi = folder.getFiles();
  while (fi.hasNext()) all.push(fi.next());
  all.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); }).slice(12).forEach(function (f) { f.setTrashed(true); }); // เก็บ 12 ชุดล่าสุด
  setSetting_('lastBackup', new Date());
  return { name: name, url: copy.getUrl(), folder: folder.getUrl() };
}
function backupNow(token) {
  var s = auth_(token, ROLE_ADMIN), r = backupJob(true);
  log_(s.name, 'สำรองข้อมูล', r.name);
  return r;
}
