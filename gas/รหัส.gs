// ============================================================
// Apps Script Web App — รับข้อมูลผลทดสอบคอนกรีตจาก Cowork
// วิธี Deploy:
//   1. เปิด Sheet "ผลทดสอบคอนกรีต" → Extensions → Apps Script
//   2. วางโค้ดนี้ทั้งหมด → Save
//   3. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy URL ที่ได้ → แจ้งกลับมาเพื่อใส่ใน Cowork skill
// ============================================================

const SHEET_NAME = 'ผลทดสอบคอนกรีต';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      return respond(false, 'ไม่พบ Sheet: ' + SHEET_NAME);
    }

    sh.appendRow([
      data.sample_date,   // A: วันที่เก็บตัวอย่าง
      data.test_date,     // B: วันที่ทดสอบ
      data.age_days,      // C: อายุ(วัน)
      data.formula_name,  // D: ชื่อสูตร
      data.cube_size,     // E: ขนาด cube
      data.result1_kn,    // F: ลูก1(kN)
      data.result2_kn,    // G: ลูก2(kN)
      data.result3_kn,    // H: ลูก3(kN)
      data.avg_kn,        // I: เฉลี่ย kN
      data.avg_mpa,       // J: เฉลี่ย MPa
      data.avg_ksc        // K: เฉลี่ย KSC
    ]);

    Logger.log('บันทึกสำเร็จ: ' + JSON.stringify(data));
    return respond(true, 'บันทึกสำเร็จ');

  } catch (err) {
    Logger.log('Error: ' + err.message);
    return respond(false, err.message);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return respond(false, 'ไม่พบ Sheet');
    const rows = sh.getDataRange().getValues();
    const headers = rows[0];
    const data = rows.slice(1).filter(r => r[0]).map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        const v = r[i];
        if (v instanceof Date) {
          obj[h] = Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
        } else {
          obj[h] = v;
        }
      });
      return obj;
    });

    // อ่าน tab วัตถุดิบ (ปูน/หิน/ทราย ทุก 10 วัน)
    var materials = [];
    var matSh = ss.getSheetByName('วัตถุดิบ');
    if (matSh) {
      var matRows = matSh.getDataRange().getValues();
      var matH = matRows[0];
      materials = matRows.slice(1).filter(r => r[0]).map(r => {
        var obj = {};
        matH.forEach((h, i) => { obj[h] = r[i] === '' ? null : r[i]; });
        return obj;
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: data, materials: materials }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return respond(false, err.message);
  }
}

// ============================================================
// นำเข้าข้อมูลจาก Sheet "สำเนาของ คำกำลังอัดก้อนปูน"
// รัน function นี้ครั้งเดียวเพื่อ import ข้อมูลเก่า
// ============================================================
function importFromOldSheet() {
  var SOURCE_ID = '1Ii0Ocr-PAIp1If2R-wMskzVteHkrGwtVuXXpvqKdKz4';
  var TARGET_NAME = 'ผลทดสอบคอนกรีต';

  var src = SpreadsheetApp.openById(SOURCE_ID);
  var srcSheet = src.getSheetByName('ค่ากำอัดอัดทั้งปี69');
  if (!srcSheet) { Logger.log('ไม่พบ sheet ค่ากำอัดอัดทั้งปี69'); return; }
  Logger.log('พบ sheet: ' + srcSheet.getName() + ' (' + (srcSheet.getLastRow()-1) + ' แถว)');
  var rows = srcSheet.getDataRange().getValues();
  var headers = rows[0].map(function(x){ return String(x).trim(); });
  var ci = {};
  headers.forEach(function(n,i){ ci[n] = i; });

  // หา index คอลัมน์
  var iDate   = ci['วันที่'] !== undefined ? ci['วันที่'] : 0;
  var iForm   = ci['สูตร']  !== undefined ? ci['สูตร']  : 1;
  var iSet    = ci['ชุดที่'] !== undefined ? ci['ชุดที่'] : 2;
  var iAge    = ci['อายุ (วัน)'] !== undefined ? ci['อายุ (วัน)'] : (ci['อายุ(วัน)'] !== undefined ? ci['อายุ(วัน)'] : 3);
  var iKN     = ci['แรงกด (kN)'] !== undefined ? ci['แรงกด (kN)'] : (ci['แรงกด(kN)'] !== undefined ? ci['แรงกด(kN)'] : 4);
  var iKSC    = ci['กำลังอัด (KSC)'] !== undefined ? ci['กำลังอัด (KSC)'] : (ci['กำลังอัด(KSC)'] !== undefined ? ci['กำลังอัด(KSC)'] : 5);

  // จัดกลุ่ม 3 ก้อน
  var groups = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[iDate]) continue;
    var rawDate = r[iDate];
    var age     = Number(r[iAge] || 0);
    var formula = String(r[iForm] || '').trim();
    var set     = String(r[iSet] || '1').trim();
    var kn      = Number(r[iKN]  || 0);
    var ksc     = Number(r[iKSC] || 0);

    // แปลงวันที่
    var testDate = '';
    if (rawDate instanceof Date) {
      testDate = Utilities.formatDate(rawDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    } else {
      var parts = String(rawDate).split('/');
      if (parts.length === 3) {
        var y = Number(parts[2]) < 100 ? 2000 + Number(parts[2]) : Number(parts[2]);
        testDate = y + '-' + parts[0].padStart(2,'0') + '-' + parts[1].padStart(2,'0');
      }
    }
    if (!testDate) continue;
    // เฉพาะปี 2026 (พ.ศ. 2569) เท่านั้น
    if (testDate.slice(0,4) !== '2026') continue;

    // วันที่ในชีต = วันที่เก็บตัวอย่าง, คำนวณวันทดสอบ = วันเก็บ + อายุ
    var sampleDate = testDate;
    var td2 = new Date(sampleDate);
    td2.setDate(td2.getDate() + age);
    testDate = Utilities.formatDate(td2, 'Asia/Bangkok', 'yyyy-MM-dd');

    var key = sampleDate + '|' + formula + '|' + set + '|' + age;
    if (!groups[key]) groups[key] = { sampleDate:sampleDate, testDate:testDate, age:age, formula:formula, kns:[], kscs:[] };
    groups[key].kns.push(kn);
    groups[key].kscs.push(ksc);
  }

  // ล้างข้อมูลเดิมทั้งหมด (เก็บแถวหัวไว้)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TARGET_NAME);
  if (!sh) { Logger.log('ไม่พบ target sheet: ' + TARGET_NAME); return; }
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  }
  Logger.log('ล้างข้อมูลเดิมแล้ว');

  var added = 0;
  var keys = Object.keys(groups);
  for (var k = 0; k < keys.length; k++) {
    var g = groups[keys[k]];
    if (g.kns.length < 1) continue;
    var r1 = g.kns[0] || '';
    var r2 = g.kns[1] || '';
    var r3 = g.kns[2] || '';
    var avgKn  = g.kns.reduce(function(a,b){return a+b;},0) / g.kns.length;
    var avgKsc = g.kscs.reduce(function(a,b){return a+b;},0) / g.kscs.length;
    sh.appendRow([
      g.sampleDate,
      g.testDate,
      g.age,
      g.formula,
      '15x15',
      r1, r2, r3,
      Math.round(avgKn*100)/100,
      '',
      Math.round(avgKsc*10)/10
    ]);
    added++;
  }

  Logger.log('นำเข้าสำเร็จ: ' + added + ' แถว');
  try { SpreadsheetApp.getUi().alert('✅ นำเข้าสำเร็จ ' + added + ' แถว'); }
  catch(e) { Logger.log('done'); }
}

// ============================================================
// เมนูเลือกเดือน — สร้างตอนเปิด Sheet
// ============================================================
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📅 เลือกเดือน');

  // หาเดือนทั้งหมดที่มีในชีต
  if (sh && sh.getLastRow() > 1) {
    var dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var months = {};
    dates.forEach(function(r) {
      var d = r[0];
      var ym = d instanceof Date
        ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
        : String(d).slice(0, 7);
      if (ym && ym.length === 7) months[ym] = true;
    });
    var sortedMonths = Object.keys(months).sort().reverse();
    sortedMonths.forEach(function(ym) {
      var parts = ym.split('-');
      var be = Number(parts[0]) + 543;
      var label = 'เดือน ' + parts[1] + '/' + String(be).slice(2) + '  (' + ym + ')';
      menu.addItem(label, 'showMonth_' + ym.replace('-', '_'));
    });
  }

  menu.addSeparator();
  menu.addItem('✅ แสดงทั้งหมด', 'showAllRows');
  menu.addItem('🔒 แค่เดือนล่าสุด', 'showLatestMonthOnly');
  menu.addToUi();
}

// dynamic stubs — Apps Script ต้องการ function จริง
// สร้างอัตโนมัติตอน onOpen แต่ต้องมี handler กลาง
function handleMonthMenu(ym) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() <= 1) return;
  var dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var rowYM = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    if (rowYM === ym) sh.showRows(i + 2, 1);
    else sh.hideRows(i + 2, 1);
  }
  try { SpreadsheetApp.getUi().alert('✅ แสดงเฉพาะ ' + ym); } catch(e) {}
}

function showAllRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) sh.showRows(2, sh.getLastRow() - 1);
  try { SpreadsheetApp.getUi().alert('✅ แสดงทั้งหมดแล้ว'); } catch(e) {}
}

// stubs สำหรับแต่ละเดือน (เพิ่มได้ถ้ามีเดือนใหม่)
function showMonth_2026_02(){handleMonthMenu('2026-02');}
function showMonth_2026_03(){handleMonthMenu('2026-03');}
function showMonth_2026_04(){handleMonthMenu('2026-04');}
function showMonth_2026_05(){handleMonthMenu('2026-05');}
function showMonth_2026_06(){handleMonthMenu('2026-06');}
function showMonth_2026_07(){handleMonthMenu('2026-07');}
function showMonth_2026_08(){handleMonthMenu('2026-08');}
function showMonth_2026_09(){handleMonthMenu('2026-09');}
function showMonth_2026_10(){handleMonthMenu('2026-10');}
function showMonth_2026_11(){handleMonthMenu('2026-11');}
function showMonth_2026_12(){handleMonthMenu('2026-12');}

// ============================================================
// ซ่อนแถวเดือนก่อน แสดงแค่เดือนล่าสุด
// ============================================================
function showLatestMonthOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { Logger.log('ไม่พบ sheet'); return; }

  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return; // มีแค่หัว

  // หาเดือนล่าสุดจากคอลัมน์ A (วันที่เก็บตัวอย่าง)
  var dates = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var latestYM = '';
  dates.forEach(function(r) {
    var d = r[0];
    var ym = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    if (ym > latestYM) latestYM = ym;
  });

  if (!latestYM) return;
  Logger.log('เดือนล่าสุด: ' + latestYM);

  // แสดง/ซ่อนทีละแถว
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var ym = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    var row = i + 2; // +2 เพราะเริ่มแถว 2
    if (ym === latestYM) {
      sh.showRows(row, 1);
    } else {
      sh.hideRows(row, 1);
    }
  }

  Logger.log('เสร็จสิ้น — แสดงเฉพาะ ' + latestYM);
  try { SpreadsheetApp.getUi().alert('✅ แสดงเฉพาะเดือน ' + latestYM); }
  catch(e) { Logger.log('done'); }
}

// ============================================================
// แก้ค่าเฉลี่ย kN และ KSC ทุกแถวให้ถูกต้อง
// คอลัมน์: F=ลูก1, G=ลูก2, H=ลูก3, I=เฉลี่ยkN, K=เฉลี่ยKSC
// ============================================================
function fixAllKscValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { Logger.log('ไม่พบ sheet'); return; }

  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  var range = sh.getRange(2, 1, lastRow - 1, 11); // A2:K
  var values = range.getValues();
  var fixed = 0;

  for (var i = 0; i < values.length; i++) {
    var kn1 = Number(values[i][5]) || 0; // F
    var kn2 = Number(values[i][6]) || 0; // G
    var kn3 = Number(values[i][7]) || 0; // H

    if (kn1 > 0 && kn2 > 0 && kn3 > 0) {
      var avgKn  = Math.round((kn1 + kn2 + kn3) / 3 * 100) / 100;
      var avgMpa = Math.round(avgKn / 22.5 * 100) / 100;
      var avgKsc = Math.round(avgKn / 22.5 * 10.197 * 10) / 10;

      values[i][8]  = avgKn;  // I: เฉลี่ย kN
      values[i][9]  = avgMpa; // J: เฉลี่ย MPa
      values[i][10] = avgKsc; // K: เฉลี่ย KSC
      fixed++;
    }
  }

  range.setValues(values);
  Logger.log('แก้ไขแล้ว ' + fixed + ' แถว');
  try { SpreadsheetApp.getUi().alert('✅ แก้ไขแล้ว ' + fixed + ' แถว'); }
  catch(e) { Logger.log('done'); }
}

function respond(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: success, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
