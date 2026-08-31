/*******************************************************************
 * ระบบค้นหาค่าสี & คุณภาพน้ำตาล — Mitr Lao Sugar
 * Backend (Google Apps Script Web App) + Google Sheets = ฐานข้อมูล
 *
 * โครงสร้างชีต:
 *   DATA   : ข้อมูลหลัก 1 แถว = 1 Lot  (คอลัมน์ A–T)
 *   CONFIG : รหัส PIN และค่า Spec
 *
 * ผู้ใช้ 2 บทบาท (แยกด้วย PIN ใน CONFIG):
 *   QC        : บันทึก/แก้ไขค่าคุณภาพ
 *   WAREHOUSE : ค้นหา/อ่านอย่างเดียว (แก้ตำแหน่ง+จำนวนคงเหลือได้)
 *******************************************************************/

var SS_ID      = '1Fhyt32uFb7eyerdXvM9zuYOB9sOGwsF4SX1my_936aM'; // Google Sheet เป้าหมาย
var DATA_SHEET = 'DATA';
var CFG_SHEET  = 'CONFIG';

/* คอลัมน์ DATA (A–T) — ต้องตรงกับหัวตารางที่ setup() สร้าง */
var COL = {
  lot:1, grade:2, prod:3, bbf:4,            // A B C D
  color:5, pol:6, moist:7, invert:8,        // E F G H
  ma:9, ash:10, sediment:11,                // I J K
  status:12, location:13, qtyTon:14, bags:15, // L M N O
  customer:16, certNo:17, source:18,        // P Q R
  updatedBy:19, updatedAt:20,               // S T
  recheck:21                                // U = วันที่สุ่มค่าสีใหม่ (วันตรวจเช็คซ้ำ)
};
var DATA_COLS = 21;   // จำนวนคอลัมน์ที่ใช้จริงใน DATA
var HEADERS = ['Lot (Strike)','เกรด','วันผลิต','Best before','สี Color (ICU)','Pol (%)',
  'ความชื้น (%)','อินเวิร์ต (%)','M.A. (mm)','Conductivity Ash','Sediment','สถานะ Spec',
  'ตำแหน่งเก็บ','คงเหลือ (ตัน)','จำนวนกระสอบ','ลูกค้า','Cert No.','แผ่นที่มา',
  'ผู้บันทึกล่าสุด','เวลาแก้ไข'];

/* ---------- ค่า Spec เริ่มต้น (แก้ได้ในชีต CONFIG) ---------- */
var SPEC_DEFAULT = {
  COLOR_MAX:1200, POL_MIN:99.00, MOIST_MAX:0.20,
  INVERT_MAX:0.80, MA_MIN:0.75, MA_MAX:1.20
};

/* =================================================================
 *  WEB APP ENTRY
 * ================================================================= */
function doGet(e) {
  // เรียกแบบ API (JSONP) จาก PWA บน GitHub Pages
  if (e && e.parameter && e.parameter.api) return apiHandler(e);
  // เปิดหน้าเว็บปกติ (ในตัว Apps Script)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('iFind — ค่าสีน้ำตาล Mitr Lao')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico');
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/* ---------- JSONP API (สำหรับ PWA ข้ามโดเมน) ---------- */
function apiHandler(e){
  var p = e.parameter, action = p.api, cb = p.callback || 'callback', out;
  try {
    if      (action === 'getRecords')    out = getRecords();
    else if (action === 'getLatestInfo') out = getLatestInfo();
    else if (action === 'checkPin')      out = checkPin(p.pin);
    else if (action === 'saveQC')        out = saveQC(p.pin, JSON.parse(p.rec));
    else if (action === 'saveStock')     out = saveStock(p.pin, p.lot, p.location, p.qtyTon, p.bags);
    else if (action === 'changePin')    out = changePinServer(p.oldPin, p.newPin);
    else if (action === 'importUpsert')  out = importUpsert(p.pin, JSON.parse(p.rows));
    else if (action === 'getCustomers')  out = getCustomers();
    else if (action === 'saveCustomers') out = saveCustomers(p.pin, JSON.parse(p.rows));
    else if (action === 'getVersion')    out = getVersion();
    else out = {error:'unknown action'};
  } catch (err) { out = {error: String(err && err.message ? err.message : err)}; }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* =================================================================
 *  AUTH — ตรวจ PIN แยกบทบาท (ตรวจฝั่งเซิร์ฟเวอร์ทุกครั้งที่เขียน)
 * ================================================================= */
function checkPin(pin){
  var cfg = getConfig();
  pin = String(pin||'').trim();
  if (pin && pin === String(cfg.PIN_QC))        return {ok:true, role:'qc',        label:'ฝ่าย QC'};
  if (pin && pin === String(cfg.PIN_WAREHOUSE)) return {ok:true, role:'warehouse', label:'คลังสินค้า'};
  return {ok:false};
}
function requireRole(pin, role){
  var r = checkPin(pin);
  if (!r.ok) throw new Error('PIN ไม่ถูกต้อง');
  if (role && r.role !== role && !(role==='warehouse' && r.role==='qc'))
    throw new Error('บทบาทนี้ไม่มีสิทธิ์ทำรายการ');
  return r;
}

function changePinServer(oldPin, newPin){
  var cfg = getConfig();
  oldPin = String(oldPin||'').trim();
  newPin = String(newPin||'').trim();
  if (newPin.length !== 4 || !/^\d{4}$/.test(newPin)) return {ok:false, error:'PIN ใหม่ต้องเป็นตัวเลข 4 หลัก'};
  if (oldPin !== String(cfg.PIN_QC)) return {ok:false, error:'PIN เดิมไม่ถูกต้อง'};
  var sh = ss().getSheetByName(CFG_SHEET);
  if (!sh) return {ok:false, error:'ไม่พบชีต CONFIG'};
  var vals = sh.getRange(1,1,sh.getLastRow()||1,2).getValues();
  var found = false;
  for (var i = 0; i < vals.length; i++){
    if (String(vals[i][0]).trim() === 'PIN_QC'){ sh.getRange(i+1,2).setValue(newPin); found = true; break; }
  }
  if (!found) sh.appendRow(['PIN_QC', newPin]);
  return {ok:true};
}

/* =================================================================
 *  READ — ดึงข้อมูล + ค้นหา
 * ================================================================= */
function getConfig(){
  var sh = ss().getSheetByName(CFG_SHEET);
  var cfg = {};
  for (var k in SPEC_DEFAULT) cfg[k] = SPEC_DEFAULT[k];
  cfg.PIN_QC = '1111'; cfg.PIN_WAREHOUSE = '2222';
  if (sh){
    var v = sh.getRange(1,1,sh.getLastRow()||1,2).getValues();
    v.forEach(function(row){ if(row[0]) cfg[String(row[0]).trim()] = row[1]; });
  }
  return cfg;
}
function spec(){
  var c = getConfig();
  return {COLOR_MAX:+c.COLOR_MAX, POL_MIN:+c.POL_MIN, MOIST_MAX:+c.MOIST_MAX,
          INVERT_MAX:+c.INVERT_MAX, MA_MIN:+c.MA_MIN, MA_MAX:+c.MA_MAX};
}

/* ข้อมูลสรุปสำหรับหน้า Home (ไม่ต้องล็อกอิน) */
function getLatestInfo(){
  var sh = ss().getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 2) return {count:0, lastDate:'', lastBy:''};
  var last = sh.getLastRow();
  var upd = sh.getRange(2,COL.updatedAt,last-1,1).getValues();
  var by  = sh.getRange(2,COL.updatedBy,last-1,1).getValues();
  var maxD = null, maxBy = '';
  for (var i=0;i<upd.length;i++){
    var d = upd[i][0];
    if (d instanceof Date && (!maxD || d > maxD)){ maxD = d; maxBy = by[i][0]; }
  }
  return {
    count: last-1,
    lastDate: maxD ? Utilities.formatDate(maxD, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
    lastBy: maxBy || ''
  };
}

/* คืนรายการทั้งหมด (สำหรับหน้าแอป) — ปลอดภัยเพราะ read-only
   ใช้ CacheService ให้ตอบไวขึ้นมาก (แคช 5 นาที ล้างอัตโนมัติเมื่อมีการบันทึก) */
function getRecords(){
  var cache = CacheService.getScriptCache();
  var hit = cache.get('records');
  if (hit) return JSON.parse(hit);
  var res = buildRecords();
  try{ cache.put('records', JSON.stringify(res), 300); }catch(e){}
  return res;
}
function bustCache(){ try{ CacheService.getScriptCache().remove('records'); CacheService.getScriptCache().remove('customers'); }catch(e){} }

/* อ่านสเปกลูกค้าจากชีต "Spec" — ส่ง raw rows ให้ client ไป parse เอง (ปรับ parser ได้โดยไม่ต้อง redeploy) */
function getCustomers(){
  var cache = CacheService.getScriptCache();
  var hit = cache.get('customers'); if (hit) return JSON.parse(hit);
  var sh = ss().getSheetByName('Spec');
  var out = {rows: []};
  if (sh && sh.getLastRow() >= 1){
    var vals = sh.getDataRange().getValues();
    out.rows = vals.map(function(row){ return row.map(function(v){
      if (v instanceof Date){ var d=v.getDate(),m=v.getMonth()+1,y=v.getFullYear(); return (d<10?'0':'')+d+'/'+(m<10?'0':'')+m+'/'+y; }
      return (v===null||v===undefined)?'':String(v);
    }); });
  }
  try{ cache.put('customers', JSON.stringify(out), 300); }catch(e){}
  return out;
}

/* บันทึกสเปกลูกค้าจากแอป → เขียนทับชีต "Spec" (QC เท่านั้น)
   รับข้อมูลแบบย่อ (structured objects) แล้วประกอบแถว 13 คอลัมน์ที่เซิร์ฟเวอร์ (URL สั้น) */
var SPEC_HEAD = ['Customer Code','Customer','Customer1','Customer2','Color','Polarization','Moisture','Invert Sugar','Target M.A','Conductivity Ash','Sediment','',''];
function _sn(v){ return (v===''||v===null||v===undefined||isNaN(+v)) ? null : +v; }
function custRowFromObj(c, code){
  var cmin=_sn(c.cmin), cmax=_sn(c.cmax), pmin=_sn(c.pmin), momax=_sn(c.momax),
      ivmax=_sn(c.ivmax), mamin=_sn(c.mamin), mamax=_sn(c.mamax), sdmax=_sn(c.sdmax);
  var color = (cmin!=null&&cmax!=null) ? (cmin+'-'+cmax+' ICU')
            : (cmin!=null ? (cmin+' ICU Min') : (cmax!=null ? (cmax+' ICU Max') : '-'));
  var ma = (mamin!=null||mamax!=null) ? ((mamin!=null?mamin:'')+' - '+(mamax!=null?mamax:'')+' mm') : '-';
  var t = String(c.t||'DCR');
  var typeStr = /sugar/i.test(t) ? t : (t+' Sugar');
  var name = String(c.n||''), cust1 = name+typeStr;
  return [code, name, cust1, '-', color,
    (pmin!=null?pmin+' % Min':'-'), (momax!=null?momax+' % Max':'-'),
    (ivmax!=null?ivmax+' % Max':'-'), ma, '-',
    (sdmax!=null?sdmax+' ppm Max':'-'), typeStr, cust1];
}
var SPEC_VER = '2026-08-24-dateLock-v3';   // ใช้ตรวจว่า deploy เวอร์ชันล่าสุดหรือยัง (getVersion)
function getVersion(){ return {ok:true, version: SPEC_VER}; }
function saveCustomers(pin, custs){
  requireRole(pin, 'qc');
  var sh = ss().getSheetByName('Spec');
  if (!sh) { sh = ss().insertSheet('Spec'); }
  custs = custs || [];
  // รับได้ทั้งแบบย่อ (object) และแบบแถวเต็ม (array) เพื่อกันเวอร์ชัน client ไม่ตรง
  var rows = custs.map(function(c, i){ return Array.isArray(c) ? c : custRowFromObj(c, i+1); });
  var all = [SPEC_HEAD].concat(rows);
  var oldR = sh.getLastRow(), oldC = sh.getLastColumn();
  if (oldR > 0) sh.getRange(1, 1, oldR, Math.max(oldC, SPEC_HEAD.length)).clearContent();
  sh.getRange(1, 1, all.length, SPEC_HEAD.length).setValues(all);
  CacheService.getScriptCache().remove('customers');
  return {ok:true, count: rows.length};
}
function buildRecords(){
  var sh = ss().getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 2) return {rows:[], spec:spec()};
  var last = sh.getLastRow();
  var vals = sh.getRange(2,1,last-1,DATA_COLS).getValues();
  var sp = spec();
  var rows = vals.filter(function(r){ return r[COL.lot-1] !== '' ; }).map(function(r){
    var o = {
      lot:str(r[COL.lot-1]), grade:str(r[COL.grade-1]), prod:dstr(r[COL.prod-1]), bbf:dstr(r[COL.bbf-1]),
      color:num(r[COL.color-1]), pol:num(r[COL.pol-1]), moist:num(r[COL.moist-1]),
      invert:num(r[COL.invert-1]), ma:num(r[COL.ma-1]), ash:num(r[COL.ash-1]), sediment:num(r[COL.sediment-1]),
      location:str(r[COL.location-1]), qtyTon:num(r[COL.qtyTon-1]), bags:num(r[COL.bags-1]),
      customer:str(r[COL.customer-1]), certNo:str(r[COL.certNo-1]), source:str(r[COL.source-1]),
      updatedBy:str(r[COL.updatedBy-1]), updatedAt:dstr(r[COL.updatedAt-1]),
      recheck:dstr(r[COL.recheck-1])
    };
    o.status = evalStatus(o, sp);
    return o;
  });
  return {rows:rows, spec:sp};
}

/* =================================================================
 *  WRITE — บันทึก/แก้ไข (QC เท่านั้น) ; คลังแก้ตำแหน่ง+จำนวนได้
 * ================================================================= */
function saveQC(pin, rec){
  var u = requireRole(pin, 'qc');
  var sh = ss().getSheetByName(DATA_SHEET);
  var rowIdx = findRowByLot(sh, rec.lot);
  var now = new Date();
  var write = function(rng){
    rng.getCell(1,COL.grade).setValue(rec.grade||'');
    rng.getCell(1,COL.prod).setValue(toDate(rec.prod||''));
    rng.getCell(1,COL.bbf).setValue(toDate(rec.bbf||''));
    rng.getCell(1,COL.color).setValue(numOrBlank(rec.color));
    rng.getCell(1,COL.pol).setValue(numOrBlank(rec.pol));
    rng.getCell(1,COL.moist).setValue(numOrBlank(rec.moist));
    rng.getCell(1,COL.invert).setValue(numOrBlank(rec.invert));
    rng.getCell(1,COL.ma).setValue(numOrBlank(rec.ma));
    rng.getCell(1,COL.ash).setValue(numOrBlank(rec.ash));
    rng.getCell(1,COL.sediment).setValue(numOrBlank(rec.sediment));
    rng.getCell(1,COL.status).setValue(evalStatus(rec, spec()));
    rng.getCell(1,COL.updatedBy).setValue(u.label);
    rng.getCell(1,COL.updatedAt).setValue(now);
  };
  if (rowIdx > 0){
    write(sh.getRange(rowIdx,1,1,20));
  } else {
    var r = sh.getLastRow()+1;
    sh.getRange(r,COL.lot).setValue(rec.lot);
    write(sh.getRange(r,1,1,20));
  }
  bustCache();
  return {ok:true};
}

/* คลัง: แก้ตำแหน่งเก็บ + จำนวนคงเหลือ (ไม่ต้องใช้ PIN — คลังเข้าได้เลย) */
function saveStock(pin, lot, location, qtyTon, bags){
  var sh = ss().getSheetByName(DATA_SHEET);
  var rowIdx = findRowByLot(sh, lot);
  if (rowIdx < 1) throw new Error('ไม่พบ Lot: '+lot);
  sh.getRange(rowIdx,COL.location).setValue(location||'');
  sh.getRange(rowIdx,COL.qtyTon).setValue(numOrBlank(qtyTon));
  sh.getRange(rowIdx,COL.bags).setValue(numOrBlank(bags));
  sh.getRange(rowIdx,COL.updatedBy).setValue('คลังสินค้า');
  sh.getRange(rowIdx,COL.updatedAt).setValue(new Date());
  bustCache();
  return {ok:true};
}

/* =================================================================
 *  IMPORT — อัปโหลด Excel: แมทเลข Strike แล้วอัปเดตค่าคุณภาพ (QC เท่านั้น)
 *  - Strike ที่มีอยู่แล้ว → อัปเดตเฉพาะคอลัมน์คุณภาพ (คงตำแหน่ง/สต๊อก/ลูกค้าเดิม)
 *  - Strike ใหม่ → เพิ่มแถวใหม่
 *  รับข้อมูลทีละก้อน (chunk) จากฝั่ง client เพื่อไม่ให้ URL ยาวเกิน
 * ================================================================= */
var IMPORT_ORDER = ['lot','grade','prod','color','pol','moist','invert','ma','ash','sediment','analyst','updatedAt','recheck'];
function importUpsert(pin, rows){
  var u = requireRole(pin, 'qc');
  if (!rows || !rows.length) return {ok:true, updated:0, added:0};
  // rows มาเป็น array ตำแหน่ง (ประหยัดความยาว URL) → แปลงกลับเป็น object
  rows = rows.map(function(a){
    if (a && !Array.isArray(a)) return a;               // เผื่อส่งมาเป็น object
    var rec = {}; for (var k=0;k<IMPORT_ORDER.length;k++) rec[IMPORT_ORDER[k]] = a[k];
    return rec;
  });
  var sh = ss().getSheetByName(DATA_SHEET);
  var last = sh.getLastRow();
  var map = {};
  if (last >= 2){
    var col = sh.getRange(2, COL.lot, last-1, 1).getValues();
    for (var i=0;i<col.length;i++){ var s=String(col[i][0]).trim(); if(s) map[s]=i+2; }
  }
  var sp = spec(), now = new Date(), updated = 0, appended = [];
  function writeQuality(rng, rec, status){
    rng.getCell(1,COL.grade).setValue(rec.grade||'');
    if (rec.prod) rng.getCell(1,COL.prod).setValue(toDate(rec.prod));
    rng.getCell(1,COL.color).setValue(numOrBlank(rec.color));
    rng.getCell(1,COL.pol).setValue(numOrBlank(rec.pol));
    rng.getCell(1,COL.moist).setValue(numOrBlank(rec.moist));
    rng.getCell(1,COL.invert).setValue(numOrBlank(rec.invert));
    rng.getCell(1,COL.ma).setValue(numOrBlank(rec.ma));
    rng.getCell(1,COL.ash).setValue(numOrBlank(rec.ash));
    rng.getCell(1,COL.sediment).setValue(numOrBlank(rec.sediment));
    rng.getCell(1,COL.status).setValue(status);
    if (rec.recheck) rng.getCell(1,COL.recheck).setValue(toDate(rec.recheck));
    rng.getCell(1,COL.updatedBy).setValue(rec.analyst ? String(rec.analyst) : u.label);
    rng.getCell(1,COL.updatedAt).setValue(rec.updatedAt ? toDate(rec.updatedAt) : now);
  }
  rows.forEach(function(rec){
    var lot = String(rec.lot||'').trim(); if(!lot) return;
    var status = evalStatus(rec, sp);
    if (map[lot]){
      writeQuality(sh.getRange(map[lot],1,1,DATA_COLS), rec, status);
      updated++;
    } else {
      var row = []; for (var k=0;k<DATA_COLS;k++) row[k]='';
      row[COL.lot-1]=lot;               row[COL.grade-1]=rec.grade||'';
      row[COL.prod-1]=toDate(rec.prod||''); row[COL.color-1]=numOrBlank(rec.color);
      row[COL.pol-1]=numOrBlank(rec.pol);       row[COL.moist-1]=numOrBlank(rec.moist);
      row[COL.invert-1]=numOrBlank(rec.invert); row[COL.ma-1]=numOrBlank(rec.ma);
      row[COL.ash-1]=numOrBlank(rec.ash);       row[COL.sediment-1]=numOrBlank(rec.sediment);
      row[COL.status-1]=status;         row[COL.source-1]='อัปโหลด';
      row[COL.recheck-1]=toDate(rec.recheck||'');
      row[COL.updatedBy-1]=rec.analyst?String(rec.analyst):u.label;
      row[COL.updatedAt-1]=rec.updatedAt ? toDate(rec.updatedAt) : now;
      appended.push(row);
    }
  });
  if (appended.length) sh.getRange(sh.getLastRow()+1, 1, appended.length, DATA_COLS).setValues(appended);
  bustCache();
  return {ok:true, updated:updated, added:appended.length};
}

/* =================================================================
 *  SPEC LOGIC
 * ================================================================= */
function evalStatus(o, sp){
  var fails = 0, n = 0;
  function chk(v, ok){ if(v===''||v===null||isNaN(v)) return; n++; if(!ok) fails++; }
  chk(o.color,  +o.color  <= sp.COLOR_MAX);
  chk(o.pol,    +o.pol    >= sp.POL_MIN);
  chk(o.moist,  +o.moist  <= sp.MOIST_MAX);
  chk(o.invert, +o.invert <= sp.INVERT_MAX);
  chk(o.ma,     +o.ma >= sp.MA_MIN && +o.ma <= sp.MA_MAX);
  if (n === 0) return '';
  if (fails === 0) return 'ผ่าน';
  if (fails === 1) return 'เฝ้าระวัง';
  return 'ไม่ผ่าน';
}

/* =================================================================
 *  HELPERS
 * ================================================================= */
function ss(){ return SpreadsheetApp.openById(SS_ID); }
function str(v){ return v===null||v===undefined?'':String(v).trim(); }
function num(v){ if(v===''||v===null||v===undefined) return ''; var n=parseFloat(v); return isNaN(n)?'':n; }
function numOrBlank(v){ if(v===''||v===null||v===undefined) return ''; var n=parseFloat(v); return isNaN(n)?'':n; }
function dstr(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return str(v);
}
/* แปลง "dd/MM/yyyy" หรือ "dd/MM/yyyy HH:mm" → Date object (กัน Google Sheets ตีความสลับวัน/เดือน)
   เขียน Date object ลงชีตแทนข้อความ แล้ว dstr อ่านกลับเป็น dd/MM/yyyy เสมอ ทุกเครื่อง */
function toDate(s){
  if (s instanceof Date) return s;
  s = String(s||'').trim(); if (!s) return '';
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (!m) return s;
  var d=+m[1], mo=+m[2], y=+m[3], h=m[4]?+m[4]:0, mi=m[5]?+m[5]:0;
  if (d<1||d>31||mo<1||mo>12) return s;
  return new Date(y, mo-1, d, h, mi);
}
function findRowByLot(sh, lot){
  if (sh.getLastRow() < 2) return -1;
  var col = sh.getRange(2,COL.lot,sh.getLastRow()-1,1).getValues();
  lot = String(lot).trim();
  for (var i=0;i<col.length;i++){ if(String(col[i][0]).trim() === lot) return i+2; }
  return -1;
}

/* =================================================================
 *  SETUP — สร้างชีต DATA + CONFIG (รันครั้งเดียว)
 * ================================================================= */
function setup(){
  var s = ss();
  var d = s.getSheetByName(DATA_SHEET) || s.insertSheet(DATA_SHEET);
  if (d.getLastRow() === 0){
    d.getRange(1,1,1,HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#B0701F').setFontColor('#FFFFFF');
    d.setFrozenRows(1); d.setColumnWidths(1,20,110);
  }
  var c = s.getSheetByName(CFG_SHEET) || s.insertSheet(CFG_SHEET);
  if (c.getLastRow() === 0){
    var cfg = [
      ['PIN_QC','1111'], ['PIN_WAREHOUSE','2222'],
      ['COLOR_MAX',1200], ['POL_MIN',99.00], ['MOIST_MAX',0.20],
      ['INVERT_MAX',0.80], ['MA_MIN',0.75], ['MA_MAX',1.20]
    ];
    c.getRange(1,1,cfg.length,2).setValues(cfg);
    c.getRange(1,1,cfg.length,1).setFontWeight('bold');
  }
  SpreadsheetApp.getUi && SpreadsheetApp.getActive(); // no-op guard
  Logger.log('Setup เสร็จ: สร้างชีต DATA + CONFIG แล้ว (PIN เริ่มต้น QC=1111, คลัง=2222 — เปลี่ยนในชีต CONFIG)');
}
