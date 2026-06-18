// ============================================================
// สร้าง Sheet "งานรอผลิต" ใน pile-production-data
// วิธีใช้:
//   1. เปิด Sheet pile-production-data → Extensions → Apps Script
//   2. เพิ่ม function นี้ต่อท้ายโค้ดเดิม
//   3. กด Run → "createPendingSheet"
//   4. กด Sync Dashboard → Sync ทันที
// ============================================================

const PENDING_SS_ID = '16keXb-jdSY6UOk1r-ZpR4ZqIYUfu9hQK7U1x7hDCof0';

function createPendingSheet() {
  var ss = SpreadsheetApp.openById(PENDING_SS_ID);

  // ลบ sheet เก่าถ้ามี (ป้องกัน duplicate)
  var existing = ss.getSheetByName('งานรอผลิต');
  if (existing) ss.deleteSheet(existing);

  var sh = ss.insertSheet('งานรอผลิต');

  // ===== Header =====
  var headers = ['วันที่', 'งานรอผลิต (m³)'];
  sh.getRange(1, 1, 1, 2).setValues([headers]);
  sh.getRange(1, 1, 1, 2)
    .setBackground('#2F5FD0')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  // ===== ข้อมูลตัวอย่าง =====
  var today = new Date();
  var sampleData = [
    [today, 1250],
  ];
  sh.getRange(2, 1, sampleData.length, 2).setValues(sampleData);

  // ===== Format =====
  // คอลัมน์ A: วันที่
  sh.getRange('A2:A1000').setNumberFormat('dd/MM/yyyy');
  // คอลัมน์ B: ตัวเลข
  sh.getRange('B2:B1000').setNumberFormat('#,##0.00');

  // สีแถวข้อมูล
  sh.getRange(2, 1, 1, 2).setBackground('#EEF3FC');

  // ขนาดคอลัมน์
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(2, 180);

  // Freeze header
  sh.setFrozenRows(1);

  // ===== คำอธิบาย =====
  sh.getRange('D1').setValue('📌 วิธีใช้');
  sh.getRange('D2').setValue('• กรอกวันที่ในคอลัมน์ A (format: วัน/เดือน/ปี)');
  sh.getRange('D3').setValue('• กรอกคิวงานรอผลิต (m³) ในคอลัมน์ B');
  sh.getRange('D4').setValue('• แถวล่างสุด = ข้อมูลล่าสุดที่จะแสดงใน Dashboard');
  sh.getRange('D5').setValue('• กด Sync Dashboard → Sync ทันที ทุกครั้งที่อัพเดท');
  sh.getRange('D1:D5').setFontColor('#64748b').setFontSize(11);
  sh.getRange('D1').setFontWeight('bold').setFontColor('#1e293b');
  sh.setColumnWidth(4, 400);

  sh.getRange('D1:D5').setBorder(true, true, true, true, false, false,
    '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('D1:D5').setBackground('#f8fafc');

  // ===== เสร็จ =====
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    '✅ สร้าง Sheet "งานรอผลิต" สำเร็จ!\n\n' +
    'กรอกข้อมูลแล้วกด Sync Dashboard → Sync ทันที\n' +
    'เพื่ออัพเดทตัวเลขบน Dashboard'
  );
  Logger.log('✅ Sheet งานรอผลิต สร้างเสร็จแล้ว');
}
