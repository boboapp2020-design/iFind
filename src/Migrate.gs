/*******************************************************************
 * Migrate.gs — รวบข้อมูลจากใบ COA ทุกแท็บ → ชีต DATA
 *
 * วิธีใช้ (รันครั้งเดียวหลังนำเข้าไฟล์ .xlsx):
 *   1) File → Import → Upload  "Update ค่าสี 2526.xlsx"
 *      เลือก "Insert new sheet(s)"  → ได้ 52 แท็บ COA
 *   2) เปิด Extensions → Apps Script  → รันฟังก์ชัน  migrate()
 *   3) ดูผลใน Executions log ว่าแผ่นไหนเข้าได้/จับไม่ได้
 *
 * อ่านเฉพาะคอลัมน์ A–T ของแต่ละแผ่น  จับตำแหน่งค่าจาก "หัวตาราง"
 * (Color, Polarization, Lot Number …) จึงรองรับ COA หลายรูปแบบ
 *******************************************************************/

function migrate(){
  var s = ss();
  var data = s.getSheetByName(DATA_SHEET);
  if (!data){ setup(); data = s.getSheetByName(DATA_SHEET); }

  // ล้างข้อมูลเดิม (เก็บหัวตารางไว้)
  if (data.getLastRow() > 1) data.getRange(2,1,data.getLastRow()-1,20).clearContent();

  var sp = spec();
  var out = [], report = [];
  var skip = {}; skip[DATA_SHEET]=1; skip[CFG_SHEET]=1;

  s.getSheets().forEach(function(sh){
    var name = sh.getName();
    if (skip[name]) return;
    var res = extractSheet(sh, sp);
    if (res.rows.length === 0){ report.push('— '+name+': ข้าม ('+res.reason+')'); return; }
    for (var i=0;i<res.rows.length;i++) out.push(res.rows[i]);
    report.push('✓ '+name+': '+res.rows.length+' แถว');
  });

  if (out.length){
    data.getRange(2,1,out.length,20).setValues(out);
    data.getRange(2,1,out.length,20).setVerticalAlignment('middle');
  }
  var summary = 'ย้ายข้อมูลเสร็จ: รวม '+out.length+' แถว\n'+report.join('\n');
  Logger.log(summary);
  return {count:out.length, sheets:report};
}

/* ---------- ดึงข้อมูลจาก 1 แผ่น COA ---------- */
function extractSheet(sh, sp){
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {rows:[], reason:'ว่าง'};
  var nCol = Math.min(20, sh.getMaxColumns());     // อ่านไม่เกินคอลัมน์ T
  var V = sh.getRange(1,1,lastRow,nCol).getValues(); // A:T

  // หาแถวหัวตาราง: ต้องมีทั้ง Color และ Polarization
  var hr = -1;
  for (var i=0;i<V.length;i++){
    var joined = rowText(V[i]);
    if (joined.indexOf('color') >= 0 && joined.indexOf('polari') >= 0){ hr = i; break; }
  }
  if (hr < 0) return {rows:[], reason:'ไม่พบหัวตาราง Color/Polarization'};

  function findCol(re){
    for (var j=0;j<V[hr].length;j++){ if (re.test(String(V[hr][j]).toLowerCase())) return j; }
    return -1;
  }
  var cLot   = findCol(/lot\s*number/),  cProd = findCol(/production/), cBbf = findCol(/best\s*before/);
  var cColor = findCol(/color/),         cPol  = findCol(/polari/),     cMoist = findCol(/moisture/);
  var cInv   = findCol(/invert/),        cMa   = findCol(/target\s*m|m\.\s*a/), cAsh = findCol(/ash/);
  var cSed   = findCol(/sediment/);
  if (cLot < 0) cLot = 1; // สำรอง: คอลัมน์ B

  // ค่าระดับใบ (ใช้กับทุกแถวของใบนั้น)
  var grade = normGrade(scanValueAfter(V, /grade of sugar/), sh.getName());
  var cust  = scanValueAfter(V, /company of customer/);
  var cert  = scanValueAfter(V, /certificate no/);

  var rows = [];
  for (var r=hr+1; r<V.length; r++){
    var line = rowText(V[r]);
    if (/specification|method|approved|reporter|remark/.test(line)) break; // ถึงท้ายตาราง
    var lotv = String(V[r][cLot]||'').trim();
    if (!lotv || !/[0-9]/.test(lotv)) continue;             // Lot ต้องมีตัวเลข
    var color = pick(V[r],cColor), pol = pick(V[r],cPol);
    if (color === '' && pol === '') continue;               // ไม่ใช่แถวข้อมูล

    var o = {
      lot:lotv, grade:grade, prod:dstr(V[r][cProd>=0?cProd:0]), bbf:dstr(V[r][cBbf>=0?cBbf:0]),
      color:color, pol:pol, moist:pick(V[r],cMoist), invert:pick(V[r],cInv),
      ma:pick(V[r],cMa), ash:pick(V[r],cAsh), sediment:pick(V[r],cSed)
    };
    rows.push([
      o.lot, o.grade, o.prod, o.bbf, o.color, o.pol, o.moist, o.invert, o.ma, o.ash, o.sediment,
      evalStatus(o, sp), '', '', '', cust, cert, sh.getName(), 'ย้ายข้อมูล', new Date()
    ]);
  }
  return {rows:rows, reason: rows.length?'':'ไม่มีแถวข้อมูล'};
}

/* ---------- helpers ---------- */
function rowText(arr){ return arr.map(function(x){return String(x).toLowerCase();}).join('|'); }
function pick(row, col){ if (col < 0) return ''; return num(row[col]); }

/* หาค่าถัดจาก label (ขวามือในแถวเดียวกัน ไม่งั้นดูแถวถัดไป) */
function scanValueAfter(V, re){
  for (var i=0;i<V.length;i++){
    for (var j=0;j<V[i].length;j++){
      if (re.test(String(V[i][j]).toLowerCase())){
        for (var k=j+1;k<V[i].length;k++){ var v=String(V[i][k]).replace(/[:：]/g,'').trim(); if(v) return v; }
        if (V[i+1]) for (var k2=0;k2<V[i+1].length;k2++){ var v2=String(V[i+1][k2]).trim(); if(v2) return v2; }
      }
    }
  }
  return '';
}
function normGrade(raw, name){
  var t = (String(raw)+' '+String(name)).toLowerCase();
  if (t.indexOf('white') >= 0) return 'White';
  if (t.indexOf('dcr') >= 0)   return 'DCR';
  raw = String(raw).replace(/sugar/i,'').trim();
  return raw || '';
}
