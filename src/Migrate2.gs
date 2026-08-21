/*******************************************************************
 * Migrate2.gs — นำเข้าข้อมูลจากแผ่น " 25 26" (ข้อมูลหลักฤดูกาล 25-26)
 *
 * แผ่น " 25 26" มี ~1,840 แถว คอลัมน์:
 *   A=วันที่, B=STRIKE, C=วันที่ผลิต, D=เวลา,
 *   E=อุณหภูมิ(ก่อน), F=อุณหภูมิ(หลัง),
 *   G=Moisture(ก่อนอบ), H=Moisture(หลังอบ),
 *   I=%POL, J=Colour, K=Sediment, L=MA,
 *   M=CV, N=%RS, O=%Ash, P=เกรด(TestB),
 *   Q=สรุปผล, R=ผู้วิเคราะห์, S=หัวหน้ากะ, T=กะ
 *
 * วิธีใช้:
 *   1) เปิด Extensions → Apps Script
 *   2) รันฟังก์ชัน migrate2()
 *   3) ข้อมูลจะ APPEND ต่อท้าย DATA sheet (ไม่ลบข้อมูลเดิม)
 *
 * หมายเหตุ: รันฟังก์ชัน clearAndMigrate2() ถ้าต้องการลบข้อมูลเดิมทั้งหมด
 *           แล้วนำเข้าใหม่จากแผ่น " 25 26" อย่างเดียว
 *******************************************************************/

function migrate2() {
  _doMigrate2(false);
}

function clearAndMigrate2() {
  _doMigrate2(true);
}

function _doMigrate2(clearFirst) {
  var s = ss();
  var data = s.getSheetByName(DATA_SHEET);
  if (!data) { setup(); data = s.getSheetByName(DATA_SHEET); }

  var src = s.getSheetByName(' 25 26');
  if (!src) {
    Logger.log('ERROR: ไม่พบแผ่น " 25 26" — ตรวจสอบว่า Import ไฟล์ Excel แล้ว');
    return;
  }

  if (clearFirst && data.getLastRow() > 1) {
    data.getRange(2, 1, data.getLastRow() - 1, 20).clearContent();
    Logger.log('ลบข้อมูลเดิมทั้งหมดแล้ว');
  }

  var sp = spec();
  var vals = src.getDataRange().getValues();
  var out = [];
  var skipped = 0;

  // row 0 = title, row 1 = headers, data starts at row 2
  for (var i = 2; i < vals.length; i++) {
    var r = vals[i];
    var strike = String(r[1] || '').trim();        // B = STRIKE
    if (!strike || strike === '' || isNaN(+strike)) { skipped++; continue; }

    var prodDate = fmtDate(r[2]);                   // C = วันที่ผลิต
    var color    = numVal(r[9]);                     // J = Colour
    var pol      = numVal(r[8]);                     // I = %POL
    var moist    = numVal(r[7]);                     // H = Moisture (หลังอบ)
    var invert   = numVal(r[13]);                    // N = %RS (Invert/Reducing Sugar)
    var ma       = numVal(r[11]);                    // L = MA
    var ash      = numVal(r[14]);                    // O = %Ash
    var sediment = numVal(r[10]);                    // K = Sediment
    var grade    = String(r[16] || '').trim();       // Q = สรุปผล (DCR/Organic/etc)
    var analyst  = String(r[17] || '').trim();       // R = ผู้วิเคราะห์
    var dateStr  = fmtDate(r[0]);                    // A = วันที่
    var timeStr  = String(r[3] || '').trim();        // D = เวลา

    // กรอง #REF! ออก
    if (/^#/.test(grade)) grade = '';
    if (/^#/.test(analyst)) analyst = '';

    // ถ้าเกรดว่าง ลองดู Test B
    if (!grade) {
      var testB = String(r[15] || '').trim();
      if (/^#/.test(testB)) testB = '';
      grade = testB ? 'DCR' : '';
    }

    // ประเมินสถานะ
    var obj = { color: color, pol: pol, moist: moist, invert: invert, ma: ma };
    var status = evalStatus(obj, sp);

    // สร้างเวลาแก้ไข
    var updatedAt = dateStr;
    if (timeStr) updatedAt += ' ' + timeStr;

    // row = [Lot, เกรด, วันผลิต, BBF, สี, Pol, ชื้น, Invert, MA, Ash, Sediment,
    //        สถานะ, ตำแหน่ง, ตัน, กระสอบ, ลูกค้า, CertNo, แผ่นที่มา, ผู้บันทึก, เวลาแก้ไข]
    out.push([
      strike,           // A: Lot (Strike)
      grade,            // B: เกรด
      prodDate,         // C: วันผลิต
      '',               // D: Best before
      color,            // E: สี
      pol,              // F: Pol
      moist,            // G: ความชื้น
      invert,           // H: Invert
      ma,               // I: MA
      ash,              // J: Ash
      sediment,         // K: Sediment
      status,           // L: สถานะ
      '',               // M: ตำแหน่ง
      '',               // N: ตัน
      '',               // O: กระสอบ
      '',               // P: ลูกค้า
      '',               // Q: Cert No.
      '25 26',          // R: แผ่นที่มา
      analyst || 'ย้ายข้อมูล',  // S: ผู้บันทึก
      updatedAt         // T: เวลาแก้ไข
    ]);
  }

  if (out.length > 0) {
    var startRow = data.getLastRow() + 1;
    data.getRange(startRow, 1, out.length, 20).setValues(out);
    // ล้างแคชแบบปลอดภัย (ทำงานได้แม้ยังไม่ได้อัปเดต Code.gs เวอร์ชันใหม่)
    try {
      if (typeof bustCache === 'function') bustCache();
      else CacheService.getScriptCache().remove('records');
    } catch (e) {}
  }

  Logger.log('=== Migrate2 เสร็จสิ้น ===');
  Logger.log('นำเข้า: ' + out.length + ' แถว');
  Logger.log('ข้าม: ' + skipped + ' แถว (ไม่มีเลข Strike)');
  Logger.log('ข้อมูลทั้งหมดใน DATA: ' + (data.getLastRow() - 1) + ' แถว');
}

/* ---------- Helper ---------- */
function numVal(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = +v;
  return isNaN(n) ? '' : n;
}

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var d = v.getDate(), m = v.getMonth() + 1, y = v.getFullYear();
    return (d < 10 ? '0' : '') + d + '/' + (m < 10 ? '0' : '') + m + '/' + y;
  }
  return String(v).trim();
}
