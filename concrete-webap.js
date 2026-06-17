// ============================================================
// Apps Script Web App - รับข้อมูลผลทดสอบคอนกรีต
// วิธี Deploy:
//   1. เปิด Sheet "ผลทดสอบคอนกรีต" > Extensions > Apps Script
//   2. วางโค้ดนี้ทั้งหมด > Save
//   3. Deploy > New deployment > Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   4. Copy URL ที่ได้ > แก้ CONCRETE_URL ใน index.html
// ============================================================

const SHEET_NAME = 'ผลทดสอบคอนกรีต';
const SPREADSHEET_ID = '1ZYGnV8AqyR3a0uTNRftkguEdBPrQOjLQbWU6ZhpiFyk';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);

    if (!sh) {
      return respond(false, 'ไม่พบ Sheet: ' + SHEET_NAME);
    }

    // หาแถวสุดท้ายจาก column A จริง (ไม่นับสูตรที่ลากลงไป)
    var colA = sh.getRange('A:A').getValues();
    var lastRow = 1;
    for (var i = colA.length - 1; i >= 1; i--) {
      if (colA[i][0] !== '') { lastRow = i + 1; break; }
    }
    var newRow = lastRow + 1;
    sh.getRange(newRow, 1, 1, 11).setValues([[
      data.sample_date,
      data.test_date,
      data.age_days,
      data.formula_name,
      data.cube_size,
      data.result1_kn,
      data.result2_kn,
      data.result3_kn,
      data.avg_kn,
      data.avg_mpa,
      data.avg_ksc
    ]]);

    return respond(true, 'บันทึกสำเร็จ');

  } catch (err) {
    return respond(false, err.message);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return respond(false, 'ไม่พบ Sheet: ' + SHEET_NAME);
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

    // อ่าน tab วัตถุดิบ
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

function importFromOldSheet() {
  var SOURCE_ID = '1Ii0Ocr-PAIp1If2R-wMskzVteHkrGwtVuXXpvqKdKz4';
  var TARGET_NAME = 'ผลทดสอบคอนกรีต';

  var src = SpreadsheetApp.openById(SOURCE_ID);
  var srcSheet = src.getSheetByName('ค่ากำอัดอัดทั้งปี69');
  if (!srcSheet) { Logger.log('ไม่พบ sheet'); return; }
  var rows = srcSheet.getDataRange().getValues();
  var headers = rows[0].map(function(x){ return String(x).trim(); });
  var ci = {};
  headers.forEach(function(n,i){ ci[n] = i; });

  var iDate = ci['วันที่'] !== undefined ? ci['วันที่'] : 0;
  var iForm = ci['สูตร']  !== undefined ? ci['สูตร']  : 1;
  var iSet  = ci['ชุดที่'] !== undefined ? ci['ชุดที่'] : 2;
  var iAge  = ci['อายุ (วัน)'] !== undefined ? ci['อายุ (วัน)'] : (ci['อายุ(วัน)'] !== undefined ? ci['อายุ(วัน)'] : 3);
  var iKN   = ci['แรงกด (kN)'] !== undefined ? ci['แรงกด (kN)'] : (ci['แรงกด(kN)'] !== undefined ? ci['แรงกด(kN)'] : 4);
  var iKSC  = ci['กำลังอัด (KSC)'] !== undefined ? ci['กำลังอัด (KSC)'] : (ci['กำลังอัด(KSC)'] !== undefined ? ci['กำลังอัด(KSC)'] : 5);

  var groups = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[iDate]) continue;
    var rawDate = r[iDate];
    var age = Number(r[iAge] || 0);
    var formula = String(r[iForm] || '').trim();
    var set = String(r[iSet] || '1').trim();
    var kn  = Number(r[iKN]  || 0);
    var ksc = Number(r[iKSC] || 0);

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
    if (testDate.slice(0,4) !== '2026') continue;

    var sampleDate = testDate;
    var td2 = new Date(sampleDate);
    td2.setDate(td2.getDate() + age);
    testDate = Utilities.formatDate(td2, 'Asia/Bangkok', 'yyyy-MM-dd');

    var key = sampleDate + '|' + formula + '|' + set + '|' + age;
    if (!groups[key]) groups[key] = { sampleDate:sampleDate, testDate:testDate, age:age, formula:formula, kns:[], kscs:[] };
    groups[key].kns.push(kn);
    groups[key].kscs.push(ksc);
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(TARGET_NAME);
  if (!sh) { Logger.log('ไม่พบ target sheet'); return; }
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  }

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
      g.sampleDate, g.testDate, g.age, g.formula, '15x15',
      r1, r2, r3,
      Math.round(avgKn*100)/100, '',
      Math.round(avgKsc*10)/10
    ]);
    added++;
  }
  Logger.log('นำเข้าสำเร็จ: ' + added + ' แถว');
  try { SpreadsheetApp.getUi().alert('นำเข้าสำเร็จ ' + added + ' แถว'); }
  catch(e) { Logger.log('done'); }
}

function onOpen() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('เลือกเดือน');

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
  menu.addItem('แสดงทั้งหมด', 'showAllRows');
  menu.addItem('แค่เดือนล่าสุด', 'showLatestMonthOnly');
  menu.addToUi();
}

function handleMonthMenu(ym) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  try { SpreadsheetApp.getUi().alert('แสดงเฉพาะ ' + ym); } catch(e) {}
}

function showAllRows() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) sh.showRows(2, sh.getLastRow() - 1);
  try { SpreadsheetApp.getUi().alert('แสดงทั้งหมดแล้ว'); } catch(e) {}
}

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

function showLatestMonthOnly() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;
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
  for (var i = 0; i < dates.length; i++) {
    var d = dates[i][0];
    var ym = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM')
      : String(d).slice(0, 7);
    if (ym === latestYM) sh.showRows(i + 2, 1);
    else sh.hideRows(i + 2, 1);
  }
  try { SpreadsheetApp.getUi().alert('แสดงเฉพาะเดือน ' + latestYM); }
  catch(e) { Logger.log('done'); }
}

function fixAllKscValues() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow <= 1) return;
  var range = sh.getRange(2, 1, lastRow - 1, 11);
  var values = range.getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var kn1 = Number(values[i][5]) || 0;
    var kn2 = Number(values[i][6]) || 0;
    var kn3 = Number(values[i][7]) || 0;
    if (kn1 > 0 && kn2 > 0 && kn3 > 0) {
      values[i][8]  = Math.round((kn1+kn2+kn3)/3*100)/100;
      values[i][9]  = Math.round((kn1+kn2+kn3)/3/22.5*100)/100;
      values[i][10] = Math.round((kn1+kn2+kn3)/3/22.5*10.197*10)/10;
      fixed++;
    }
  }
  range.setValues(values);
  try { SpreadsheetApp.getUi().alert('แก้แล้ว ' + fixed + ' แถว'); }
  catch(e) { Logger.log('done'); }
}

function setupMaterialsSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('วัตถุดิบ');
  if (!sh) sh = ss.insertSheet('วัตถุดิบ');
  var headers = ['วันที่','ปูนรวม','ปูนเสาเหล็ก','ปูนI18','หิน3/4','หิน1','ทราย'];
  var firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (firstRow[0] === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  try { SpreadsheetApp.getUi().alert('พร้อมแล้ว! กรอกข้อมูลใน tab "วัตถุดิบ" ได้เลย'); }
  catch(e) { Logger.log('done'); }
}

// Auto-fill เดือนปี (คอลัมน์ L) เมื่อกรอกวันที่ในคอลัมน์ A
function onEdit(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sh = range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  if (range.getColumn() !== 1) return;
  var row = range.getRow();
  if (row <= 1) return;
  var val = range.getValue();
  if (!val) return;
  var d = new Date(val);
  if (isNaN(d.getTime())) return;
  var ym = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM');
  sh.getRange(row, 12).setValue(ym);
}

function respond(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: success, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
