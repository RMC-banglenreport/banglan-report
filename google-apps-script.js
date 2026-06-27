// ============================================================
// Google Apps Script — sync Google Sheet → Supabase
// วิธีใช้:
//   1. เปิด Google Sheet → Extensions → Apps Script
//   2. วางโค้ดทั้งหมดนี้ แทนที่โค้ดเดิม
//   3. แก้ SUPABASE_URL และ SUPABASE_KEY ด้านล่าง
//   4. กด Save → Run "syncAll" ครั้งแรก (อนุญาต permission)
//   5. ตั้ง trigger: Triggers → Add → onEdit / time-driven
// ============================================================

const SUPABASE_URL = 'https://npxzerdirspwunuckcqr.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5weHplcmRpcnNwd3VudWNrY3FyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAxMjUxMjIsImV4cCI6MjA5NTcwMTEyMn0.4C1MucMeqPozXSfErLM44at7dykfzfFQvpVnoqmrMQI';
const SS_MAIN_ID   = '16keXb-jdSY6UOk1r-ZpR4ZqIYUfu9hQK7U1x7hDCof0';
const SS_DAMAGE_ID = '1VYLouTO0BpHkLPnsEZWA2-JsTHqmPM0U1v8GfyTRIoM';

// ============================================================
// MAIN: sync ข้อมูลทั้งหมดจาก Sheet → Supabase
// ============================================================
function syncAll() {
  const ss = SpreadsheetApp.openById(SS_MAIN_ID);

  writeComputedColumns(ss);
  syncMonthly(ss);
  syncDecades(ss);
  syncCauses(ss);

  // เพิ่มคอลัมน์ใหม่ถ้ายังไม่มี (ทำอัตโนมัติ ไม่ต้องกดเอง)
  addPersonnelColumns(ss);

  // sync ข้อมูลเสียหาย (อยู่ใน Sheet การผลิตเดิม)
  syncDamageItems(ss);
  syncDamageSales(ss);
  syncPendingWork(ss);

  try {
    SpreadsheetApp.getUi().alert('✅ Sync สำเร็จ! ข้อมูลอัปเดตแล้ว');
  } catch(e) {
    Logger.log('✅ Sync สำเร็จ! ข้อมูลอัปเดตแล้ว');
  }
}

// ============================================================
// คำนวณ % ผลิต และ ส่วนต่างสต็อค แล้วเขียนกลับลง Sheet "รายเดือน"
// % ผลิต = MAX(คิวผลิตจริง/(170×วันผลิต), แพผลิต/(12×วันผลิต)) × 100
// ส่วนต่างสต็อค = สต็อคจริง − (สต็อคเดือนก่อน + ผลิต − ขนส่ง)
// ============================================================
function writeComputedColumns(ss) {
  const sheet = ss.getSheetByName('รายเดือน');
  if (!sheet) { Logger.log('ไม่พบ Sheet "รายเดือน"'); return; }

  const allValues = sheet.getDataRange().getValues();
  if (allValues.length < 2) return;

  const headers = allValues[0];
  const col = {};
  headers.forEach((h, i) => { col[String(h).trim()] = i; });

  // ตรวจสอบคอลัมน์ที่ต้องการ
  const needCols = ['ปี','เดือน','คิวผลิตจริง','แพผลิต','วันผลิต','คิวขนส่ง','สต็อคคงเหลือ'];
  for (const c of needCols) {
    if (col[c] === undefined) { Logger.log(`❌ ไม่พบคอลัมน์ "${c}" — ข้าม writeComputedColumns`); return; }
  }

  // หาหรือสร้างคอลัมน์ % ผลิต และ ส่วนต่างสต็อค
  function ensureCol(name) {
    let idx = col[name];
    if (idx === undefined) {
      idx = headers.length;
      col[name] = idx;
      headers.push(name);
      sheet.getRange(1, idx + 1).setValue(name)
        .setBackground('#d9ead3').setFontWeight('bold').setHorizontalAlignment('center');
    }
    return idx;
  }
  const pctCol  = ensureCol('% ผลิต');
  const diffCol = ensureCol('ส่วนต่างสต็อค');

  // เรียงแถวข้อมูล (ปี, เดือน) เพื่อคำนวณสต็อคสะสม
  const dataRows = [];
  for (let i = 1; i < allValues.length; i++) {
    const r = allValues[i];
    const yr = Number(r[col['ปี']]), mo = Number(r[col['เดือน']]);
    if (!yr || !mo) continue;
    dataRows.push({ rowIdx: i, yr, mo, r });
  }
  dataRows.sort((a, b) => a.yr !== b.yr ? a.yr - b.yr : a.mo - b.mo);

  // คำนวณและเขียนทีละแถว
  const prevStock = {}; // key = "yr-mo" → stock value

  for (const d of dataRows) {
    const r = d.r;
    const produced     = Number(r[col['คิวผลิตจริง']]) || 0;
    const raftProduced = Number(r[col['แพผลิต']])       || 0;
    const days         = Number(r[col['วันผลิต']])       || 1;
    const transport    = Number(r[col['คิวขนส่ง']])      || 0;
    const stock        = Number(r[col['สต็อคคงเหลือ']]) || 0;

    // % ผลิต
    const pctQ = produced      / (170 * days) * 100;
    const pctR = raftProduced  / (12  * days) * 100;
    const pct  = Math.max(pctQ, pctR);

    // ส่วนต่างสต็อค (ต้องรู้สต็อคเดือนก่อน)
    let diff = '';
    const prevKey = d.mo === 1
      ? `${d.yr - 1}-12`
      : `${d.yr}-${d.mo - 1}`;
    if (prevStock[prevKey] !== undefined) {
      diff = stock - (prevStock[prevKey] + produced - transport);
    }

    // บันทึก stock เดือนนี้ไว้ใช้เดือนหน้า
    prevStock[`${d.yr}-${d.mo}`] = stock;

    // เขียนลง sheet
    sheet.getRange(d.rowIdx + 1, pctCol  + 1).setValue(produced > 0 ? Math.round(pct * 10) / 10 : '');
    sheet.getRange(d.rowIdx + 1, diffCol + 1).setValue(diff === '' ? '' : Math.round(diff * 10) / 10);
  }

  Logger.log(`✅ writeComputedColumns: คำนวณ ${dataRows.length} แถว`);
}

// ============================================================
// เสาเสีย: sync 4 sheets → damage_* tables
// ============================================================
function readSheetRows(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) { Logger.log('ไม่พบ Sheet "' + name + '"'); return null; }
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { headers: [], data: [] };
  const headers = rows[0];
  const col = {}; headers.forEach((h, i) => { col[String(h).trim()] = i; });
  return { col, rows };
}

function syncDamageYearly(ss) {
  const r = readSheetRows(ss, 'เสาเสีย-รายปี'); if (!r) return;
  const c = r.col;
  const data = r.rows.slice(1).filter(x => x[c['ปี']]);
  const recs = data.map(x => ({ year: Number(x[c['ปี']]), sales: Number(x[c['ยอดขาย']]) || null, loss: Number(x[c['มูลค่าเสียหาย']]) || null }));
  if (recs.length) {
    deleteFromSupabase('damage_yearly', 'year=gte.0');
    upsertToSupabase('damage_yearly', recs, 'year');
  }
  Logger.log('✅ เสาเสีย-รายปี: ' + recs.length + ' แถว');
}

function syncDamageMonthly(ss) {
  const r = readSheetRows(ss, 'เสาเสีย-รายเดือน'); if (!r) return;
  const c = r.col;
  const data = r.rows.slice(1).filter(x => x[c['ปี']] && x[c['เดือน']]);
  const recs = data.map(x => ({ year: Number(x[c['ปี']]), month: Number(x[c['เดือน']]), sales: Number(x[c['ยอดขาย']]) || null, loss: Number(x[c['มูลค่าเสียหาย']]) || null }));
  if (recs.length) {
    const years = [...new Set(recs.map(v => v.year))];
    years.forEach(y => deleteFromSupabase('damage_monthly', 'year=eq.' + y));
    upsertToSupabase('damage_monthly', recs, 'year,month');
  }
  Logger.log('✅ เสาเสีย-รายเดือน: ' + recs.length + ' แถว');
}

function syncDamageCauses(ss) {
  const r = readSheetRows(ss, 'เสาเสีย-สาเหตุ'); if (!r) return;
  const c = r.col;
  const data = r.rows.slice(1).filter(x => x[c['ปี']] && x[c['สาเหตุ']]);
  const recs = data.map((x, i) => ({ year: Number(x[c['ปี']]), sort_order: i + 1, cause_name: String(x[c['สาเหตุ']]), value: Number(x[c['มูลค่า']]) || 0 }));
  if (recs.length) {
    const years = [...new Set(recs.map(v => v.year))];
    years.forEach(y => deleteFromSupabase('damage_causes', 'year=eq.' + y));
    upsertToSupabase('damage_causes', recs, 'year,sort_order');
  }
  Logger.log('✅ เสาเสีย-สาเหตุ: ' + recs.length + ' แถว');
}

function syncDamageCustomers(ss) {
  const r = readSheetRows(ss, 'เสาเสีย-ลูกค้า'); if (!r) return;
  const c = r.col;
  const data = r.rows.slice(1).filter(x => x[c['ปี']] && x[c['ลูกค้า']]);
  const recs = data.map((x, i) => ({ year: Number(x[c['ปี']]), sort_order: i + 1, customer_name: String(x[c['ลูกค้า']]), count: Number(x[c['จำนวนครั้ง']]) || 0 }));
  if (recs.length) {
    const years = [...new Set(recs.map(v => v.year))];
    years.forEach(y => deleteFromSupabase('damage_customers', 'year=eq.' + y));
    upsertToSupabase('damage_customers', recs, 'year,sort_order');
  }
  Logger.log('✅ เสาเสีย-ลูกค้า: ' + recs.length + ' แถว');
}

// ============================================================
// สร้าง 4 sheets เสาเสีย พร้อมข้อมูลเริ่มต้น (รันครั้งเดียว)
// เลือกฟังก์ชัน setupDamageSheets แล้วกด Run
// ============================================================
function setupDamageSheets() {
  const ss = SpreadsheetApp.openById('16keXb-jdSY6UOk1r-ZpR4ZqIYUfu9hQK7U1x7hDCof0');

  function mk(name, headers, rows) {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    const hr = sh.getRange(1, 1, 1, headers.length);
    hr.setBackground('#d65b5b').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }

  mk('เสาเสีย-รายปี', ['ปี','ยอดขาย','มูลค่าเสียหาย'], [
    [2565,201534492.77,1668996.73],
    [2566,229414514.05,1632024.54],
    [2568,195137597.38,1459639.97],
    [2569,87867775.23,582095.78],
  ]);

  mk('เสาเสีย-รายเดือน', ['ปี','เดือน','ยอดขาย','มูลค่าเสียหาย'], [
    [2569,1,17345153.30,110734.79],
    [2569,2,19459745.47,87310.30],
    [2569,3,21730581.59,134545.11],
    [2569,4,11730756.37,67319.12],
    [2569,5,17601538.50,182186.46],
  ]);

  mk('เสาเสีย-สาเหตุ', ['ปี','สาเหตุ','มูลค่า'], [
    [2569,'เสาเข็ม Fail',223526.62],
    [2569,'เสาร้าวในกอง/ขนส่ง/ปีกแตก',104645.72],
    [2569,'เสาเข็มหัวแตก',51108.10],
    [2569,'ตอกเอียง/เทสต์ไม่ผ่าน/ผิดหมุด',47622.16],
    [2569,'พนักงานบริษัท',40865],
    [2569,'เสาเสียในสต๊อค',34270],
    [2569,'ผู้รับเหมา',30398.08],
    [2569,'สั่งผลิตผิด/ส่งผิด/ปรับปรุง',28666.32],
    [2569,'เสาเข็มหายจากนับสต๊อค',15005],
    [2569,'ปั้นจั่นลากหัก/ชน',5988.79],
    [2569,'อื่นๆ',0],
    [2569,'ตอกสไลด์',0],
    [2569,'เสาเข็มบริษัทอื่น',0],
  ]);

  mk('เสาเสีย-ลูกค้า', ['ปี','ลูกค้า','จำนวนครั้ง'], [
    [2569,'บจก.วิทวี',5],[2569,'คุณนงลักษณ์',4],[2569,'บ.ดับบลิวเฮ้าส์',4],
    [2569,'บจก.เคพีวาย',2],[2569,'บจก.กาญจนเควต',2],[2569,'บ้านพักอาศัย (คุณสนิท)',2],
    [2569,'บจก.ไทยร็อคเฟอร์เทค',2],[2569,'บจก.อาร์ตคอนกรีต',2],[2569,'คุณติณน์',2],
    [2569,'บจก.ทีวายเค',1],[2569,'บ.หอมศีล',1],[2569,'บจก.เหรียญทอง',1],
    [2569,'บจก.ไอยราวาณิชย์',1],[2569,'บ.ไทยมั่นคงพลาสติก',1],[2569,'บจก.พีช พลัส พร็อพเพอร์ตี้',1],
    [2569,'บจก.บุญโสฬส',1],[2569,'บจก.เอรอส',1],[2569,'บจก.บีบี เอสเตท',1],
    [2569,'บจก.ลิ้มพัฒนาดีซีเมนต์',1],[2569,'เอริชเฮ้าส์',1],[2569,'บจก.โกเด้นไพรซ์',1],
  ]);

  try { SpreadsheetApp.getUi().alert('✅ สร้าง 4 sheets เสาเสียแล้ว — กด syncAll เพื่อส่งขึ้น Supabase'); }
  catch(e) { Logger.log('✅ สร้าง 4 sheets เสาเสียแล้ว'); }
}

// ============================================================
// Sync Sheet "รายเดือน" → production_monthly
// ============================================================
function syncMonthly(ss) {
  const sheet = ss.getSheetByName('รายเดือน');
  if (!sheet) { Logger.log('ไม่พบ Sheet "รายเดือน"'); return; }

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });

  // กรองเฉพาะแถวที่มีข้อมูลจริง (ต้องมีคิวผลิตจริง > 0)
  const producedIdx = col['คิวผลิตจริง'];
  const data = rows.slice(1).filter(r => {
    if (!r[0] || !r[1]) return false; // ต้องมี ปี และ เดือน
    if (producedIdx === undefined) return false;
    const val = r[producedIdx];
    return val !== '' && val !== null && Number(val) > 0;
  });

  if (data.length === 0) { Logger.log('ไม่มีข้อมูลรายเดือน'); return; }

  const records = data.map(r => ({
    year:          Number(r[col['ปี']]),
    month:         Number(r[col['เดือน']]),
    order_qty:     col['คิวสั่งผลิต']   !== undefined ? (Number(r[col['คิวสั่งผลิต']])   || null) : null,
    produced:      col['คิวผลิตจริง']    !== undefined ? (Number(r[col['คิวผลิตจริง']])    || null) : null,

    cancel:        col['ค้าง-ยกเลิก']   !== undefined ? (Number(r[col['ค้าง-ยกเลิก']])   || null) : null,
    days:          col['วันผลิต']        !== undefined ? (Number(r[col['วันผลิต']])        || null) : null,
    per_day:       col['เฉลี่ยคิว/วัน'] !== undefined ? (Number(r[col['เฉลี่ยคิว/วัน']]) || null) : null,
    transport:     col['คิวขนส่ง']       !== undefined ? (Number(r[col['คิวขนส่ง']])       || null) : null,
    stock:         col['สต็อคคงเหลือ']  !== undefined ? (Number(r[col['สต็อคคงเหลือ']])  || null) : null,
    raft_order:    col['แพสั่ง']         !== undefined ? (Number(r[col['แพสั่ง']])         || null) : null,
    raft_produced: col['แพผลิต']         !== undefined ? (Number(r[col['แพผลิต']])         || null) : null,
    raft_cancel:   col['แพค้าง']         !== undefined ? (Number(r[col['แพค้าง']])         || null) : null,
    is_sample:     (() => {
      const k = 'ข้อมูลจริง (TRUE=จริง)';
      const v = col[k] !== undefined ? r[col[k]] : true;
      return v === false || v === 'FALSE' ? true : false;
    })(),
  }));

  // ลบข้อมูลปีที่มีใน Sheet ก่อน แล้ว insert ใหม่ทั้งหมด
  const years = [...new Set(records.map(r => r.year))];
  years.forEach(y => deleteFromSupabase('production_monthly', `year=eq.${y}`));
  upsertToSupabase('production_monthly', records, 'year,month');
  Logger.log(`✅ รายเดือน: sync ${records.length} แถว (ลบ+insert ใหม่)`);
}

// ============================================================
// Sync Sheet "คุณภาพ10วัน" → quality_decade
// ============================================================
function syncDecades(ss) {
  const sheet = ss.getSheetByName('คุณภาพ10วัน');
  if (!sheet) { return; }

  const rows = sheet.getDataRange().getValues();
  const data = rows.slice(1).filter(r => r[0] && r[1]);

  if (data.length === 0) return;

  const records = data.map(r => ({
    year:         Number(r[0]),
    sort_order:   Number(r[1]),
    label:        String(r[2]),
    strength:     r[3] !== '' ? Number(r[3]) : null,
    cement_total: r[4] !== '' ? Number(r[4]) : null,
    cement_big:   r[5] !== '' ? Number(r[5]) : null,
    cement_i18:   r[6] !== '' ? Number(r[6]) : null,
    rock34:       r[7] !== '' ? Number(r[7]) : null,
    rock1:        r[8] !== '' ? Number(r[8]) : null,
    sand:         r[9] !== '' ? Number(r[9]) : null,
  }));

  const years = [...new Set(records.map(r => r.year))];
  years.forEach(y => deleteFromSupabase('quality_decade', `year=eq.${y}`));
  upsertToSupabase('quality_decade', records, 'year,sort_order');
  Logger.log(`✅ คุณภาพ10วัน: sync ${records.length} แถว (ลบ+insert ใหม่)`);
}

// ============================================================
// Sync Sheet "สาเหตุค้าง" → cancel_causes
// ============================================================
function syncCauses(ss) {
  const sheet = ss.getSheetByName('สาเหตุค้าง');
  if (!sheet) { Logger.log('ไม่พบ Sheet "สาเหตุค้าง"'); return; }

  const rows = sheet.getDataRange().getValues();
  const data = rows.slice(1).filter(r => r[0] && r[1] !== '');

  if (data.length === 0) return;

  const records = data.map(r => ({
    year:         Number(r[0]),
    month:        Number(r[1]),
    cause_index:  Number(r[2]),
    cause_label:  String(r[3]),
    count:        Number(r[4]) || 0,
  }));

  const years = [...new Set(records.map(r => r.year))];
  years.forEach(y => deleteFromSupabase('cancel_causes', `year=eq.${y}`));
  upsertToSupabase('cancel_causes', records, 'year,month,cause_index');
  Logger.log(`✅ สาเหตุค้าง: sync ${records.length} แถว (ลบ+insert ใหม่)`);
}

// ============================================================
// Helper: ลบ records จาก Supabase ตาม filter
// ============================================================
function deleteFromSupabase(table, filter) {
  const url = `${SUPABASE_URL}/rest/v1/${table}?${filter}`;
  const options = {
    method: 'delete',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
    },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200 && code !== 204) {
    Logger.log(`⚠️ Delete ${table}: ${code}`);
  }
}

// ============================================================
// Helper: truncate damage_items (ลบทุกแถวแน่นอน)
// ============================================================
function truncateDamageItems() {
  var url = SUPABASE_URL + '/rest/v1/damage_items?id=gt.0';
  var res = UrlFetchApp.fetch(url, {
    method: 'delete',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Prefer': 'return=minimal'
    },
    muteHttpExceptions: true
  });
  Logger.log('truncate damage_items: HTTP ' + res.getResponseCode());
  Utilities.sleep(500);
}

// ============================================================
// Helper: insert records เป็น batch (ไม่ใช้ on_conflict=id)
// ============================================================
function insertInBatches(table, records, batchSize) {
  const url = `${SUPABASE_URL}/rest/v1/${table}`;
  for (let i = 0; i < records.length; i += batchSize) {
    const batch = records.slice(i, i + batchSize);
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Prefer': 'return=minimal',
      },
      payload: JSON.stringify(batch),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 200 && code !== 201) {
      Logger.log(`❌ insert batch ${i}: HTTP ${code} — ${res.getContentText()}`);
      throw new Error(`Insert error: ${code}`);
    }
    Logger.log(`✅ insert batch ${i}–${Math.min(i+batchSize, records.length)-1}`);
  }
}

// ============================================================
// Helper: upsert records ไปยัง Supabase
// ============================================================
function upsertToSupabase(table, records, onConflict) {
  if (!records || records.length === 0) return;

  const url = `${SUPABASE_URL}/rest/v1/${table}?on_conflict=${onConflict}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Prefer': 'resolution=merge-duplicates',
    },
    payload: JSON.stringify(records),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code !== 200 && code !== 201) {
    Logger.log(`❌ Error ${table}: ${code} — ${response.getContentText()}`);
    throw new Error(`Supabase error ${code}: ${response.getContentText()}`);
  }
}

// ============================================================
// Auto-sync เมื่อแก้ไข Sheet (trigger onEdit)
// ============================================================
function onEdit(e) {
  const sheetName = e.source.getActiveSheet().getName();
  const validSheets = ['รายเดือน', 'คุณภาพ10วัน', 'สาเหตุค้าง',
    'รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB',
    'เสียหายในโรงงานบางเลน รหัสID',
    'ยอดขายประจำเดือน'];

  if (!validSheets.includes(sheetName)) return;

  // debounce: รอ 2 วินาทีก่อน sync เพื่อไม่ให้ sync บ่อยเกินไป
  SpreadsheetApp.flush();
  syncAll_silent();
}

function syncAll_silent() {
  try {
    const ss = SpreadsheetApp.openById(SS_MAIN_ID);
    syncMonthly(ss);
    syncDecades(ss);
    syncCauses(ss);
    syncDamageItems(ss);
    syncDamageSales(ss);
    Logger.log('✅ Auto-sync สำเร็จ');
  } catch(e) {
    Logger.log('❌ Auto-sync error: ' + e.message);
  }
}

// ============================================================
// สร้าง Menu ใน Google Sheet
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Sync Dashboard')
    .addItem('🔄 Sync ทันที', 'syncAll')
    .addItem('📋 ดู Log', 'showLog')
    .addSeparator()
    .addItem('🔀 รวมสาเหตุเข้าคอลัมน์ H (ทำครั้งเดียว)', 'mergeCauseToH')
    .addItem('🔧 จัดคอลัมน์ ID: ชุดผู้รับเหมา=I, ชื่อพนักงาน=J (ทำครั้งเดียว)', 'fixIDColumns')
    .addToUi();
}

// รวมคอลัมน์ G (ประเภทเสียหาย) เข้า H (สาเหตุ) — รันครั้งเดียว
function mergeCauseToH() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'เสียหายในโรงงานบางเลน รหัสID';
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { SpreadsheetApp.getUi().alert('ไม่พบ Sheet: ' + sheetName); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const colG = 7; // G = ประเภทเสียหาย
  const colH = 8; // H = สาเหตุ

  const gVals = sheet.getRange(2, colG, lastRow - 1, 1).getValues();
  const hVals = sheet.getRange(2, colH, lastRow - 1, 1).getValues();

  const newH = [];
  const clearG = [];

  for (let i = 0; i < gVals.length; i++) {
    const g = String(gVals[i][0] || '').trim();
    const h = String(hVals[i][0] || '').trim();

    if (g && !h) {
      newH.push([g]);   // H ว่าง → ย้าย G มาใส่ H
      clearG.push([i + 2, true]);
    } else if (g && h) {
      newH.push([g + ' — ' + h]);  // ทั้งคู่มี → รวมกัน
      clearG.push([i + 2, true]);
    } else {
      newH.push([h]);   // G ว่าง → เก็บ H เดิม
    }
  }

  sheet.getRange(2, colH, newH.length, 1).setValues(newH);
  clearG.forEach(([row]) => sheet.getRange(row, colG).clearContent());

  SpreadsheetApp.getUi().alert('✅ รวมสาเหตุเข้าคอลัมน์ H เรียบร้อย\nคอลัมน์ G ถูกล้างแล้ว');
}

// ผสานเซลล์คอลัมน์ I ตาม H และใส่ dropdown เฉพาะแถวที่ G = "ผู้รับเหมา"
function setupContractorColumn(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const colG = headers.indexOf('ประเภทเสียหาย') + 1; // 1-based
  const colH = headers.indexOf('สาเหตุ') + 1;
  const colI = headers.indexOf('ชุดผู้รับเหมา') + 1;
  if (!colG || !colH || !colI) { Logger.log('⚠️ ไม่พบคอลัมน์ที่ต้องการใน ID sheet'); return; }

  const contractorChoices = ['ชุดยกเสาข้างใน','ยกเสาด้านนอก','เข้าแบบ','ถอดแบบ','ชุดจี้คอนกรีต'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(contractorChoices, true)
    .setAllowInvalid(true)
    .build();

  // อ่าน merge ranges ของคอลัมน์ H
  const merges = sheet.getRange(2, colH, lastRow - 1, 1).getMergedRanges();
  const mergedRows = new Set(); // แถวที่อยู่ใน merge range แล้ว

  // ล้าง merge เก่าในคอลัมน์ I ก่อน
  sheet.getRange(2, colI, lastRow - 1, 1).breakApart();

  // อ่านค่าคอลัมน์ G ทั้งหมด
  const gVals = sheet.getRange(2, colG, lastRow - 1, 1).getValues();

  merges.forEach(mr => {
    const startRow = mr.getRow();
    const numRows  = mr.getNumRows();
    const typeVal  = String(sheet.getRange(startRow, colG).getValue() || '').trim();

    // ผสาน I เฉพาะถ้า G = ผู้รับเหมา
    if (/รับเหมา/i.test(typeVal)) {
      if (numRows > 1) {
        sheet.getRange(startRow, colI, numRows, 1).merge();
      }
      sheet.getRange(startRow, colI).setDataValidation(rule);
    }

    for (let r = startRow; r < startRow + numRows; r++) mergedRows.add(r);
  });

  // แถวที่ไม่ได้ merge — ตรวจ G แต่ละแถว
  for (let i = 0; i < gVals.length; i++) {
    const row = i + 2;
    if (mergedRows.has(row)) continue;
    const typeVal = String(gVals[i][0] || '').trim();
    if (/รับเหมา/i.test(typeVal)) {
      sheet.getRange(row, colI).setDataValidation(rule);
    } else {
      sheet.getRange(row, colI).clearDataValidations();
    }
  }

  Logger.log('✅ ID: ตั้ง dropdown + merge ชุดผู้รับเหมาเรียบร้อย');
}

// จัดคอลัมน์ ID: ลบ ชุดคนงาน, ให้ ชุดผู้รับเหมา=I, ชื่อพนักงาน=J
function fixIDColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('เสียหายในโรงงานบางเลน รหัสID');
  if (!sheet) { SpreadsheetApp.getUi().alert('ไม่พบ Sheet ID'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  // ลบคอลัมน์ที่ชื่อ "ชุดคนงาน" ออก (จากขวาไปซ้าย)
  for (let i = headers.length - 1; i >= 0; i--) {
    if (headers[i] === 'ชุดคนงาน') {
      sheet.deleteColumn(i + 1);
      Logger.log(`🗑️ ลบคอลัมน์ "ชุดคนงาน" (col ${i+1})`);
    }
  }

  // อ่าน headers ใหม่
  const h2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const idxContractor = h2.indexOf('ชุดผู้รับเหมา'); // 0-based
  const idxEmployee   = h2.indexOf('ชื่อพนักงาน');

  // ถ้า ชุดผู้รับเหมา อยู่หลัง ชื่อพนักงาน → สลับตำแหน่ง
  if (idxContractor > idxEmployee && idxContractor >= 0 && idxEmployee >= 0) {
    const lastRow = sheet.getLastRow();
    const colC = idxContractor + 1; // 1-based
    const colE = idxEmployee + 1;

    // copy ข้อมูลสลับกัน
    const cVals = sheet.getRange(1, colC, lastRow, 1).getValues();
    const eVals = sheet.getRange(1, colE, lastRow, 1).getValues();
    sheet.getRange(1, colE, lastRow, 1).setValues(cVals);
    sheet.getRange(1, colC, lastRow, 1).setValues(eVals);
    Logger.log('✅ สลับ ชุดผู้รับเหมา ↔ ชื่อพนักงาน แล้ว');
  }

  SpreadsheetApp.getUi().alert('✅ จัดคอลัมน์เรียบร้อย\nI = ชุดผู้รับเหมา | J = ชื่อพนักงาน');
}

function showLog() {
  SpreadsheetApp.getUi().alert(
    'ดู Log ได้ที่:\nApps Script → View → Logs (Ctrl+Enter)'
  );
}

// ============================================================
// Sync Sheet "รายการเสียหาย" → damage_items
// คอลัมน์ใน Sheet: ปี | เดือน | รหัส | ลูกค้า | ประเภท | สาเหตุ | มูลค่า(บาท)
// รหัส: REB = รับคืนสินค้า, ROB = เสียหายหน้างาน, ID = เสียหายในโรงงาน
// ============================================================
// Helper: แปลงชื่อเดือนไทย → เลขเดือน 1-12
// ============================================================
function parseThaiMonthYear(text) {
  const MONTHS = {
    'มกราคม':1,'กุมภาพันธ์':2,'มีนาคม':3,'เมษายน':4,
    'พฤษภาคม':5,'มิถุนายน':6,'กรกฎาคม':7,'สิงหาคม':8,
    'กันยายน':9,'ตุลาคม':10,'พฤศจิกายน':11,'ธันวาคม':12
  };
  const s = String(text).trim();
  for (const [name, num] of Object.entries(MONTHS)) {
    if (s.includes(name)) {
      const yearMatch = s.match(/\d{4}/);
      const year = yearMatch ? Number(yearMatch[0]) : null;
      return { month: num, year };
    }
  }
  return null;
}

// Helper: แปลงตัวเลขที่อาจมี comma เช่น "11,730,756.37" → number
function toNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  return Number(String(val).replace(/,/g, '')) || 0;
}

// Helper: แปลง วันที่ "7/4/69" หรือ Date object → { month, year (พ.ศ.) }
function parseDateCell(val) {
  if (!val) return null;
  // ถ้าเป็น Date object จาก Google Sheet
  if (val instanceof Date) {
    const m = val.getMonth() + 1;
    const y = val.getFullYear() + 543; // ค.ศ. → พ.ศ.
    return { month: m, year: y };
  }
  // ถ้าเป็น string เช่น "7/4/69"
  const s = String(val).trim();
  const parts = s.split('/');
  if (parts.length === 3) {
    const m = Number(parts[1]);
    let y = Number(parts[2]);
    if (y < 100) y += 2500; // 69 → 2569
    return { month: m, year: y };
  }
  return null;
}

// ============================================================
// Sync Sheet "รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB"
// และ "เสียหายในโรงงานบางเลน รหัสID" → damage_items
// ============================================================
function syncDamageItems(ss) {
  const records = [];

  // --- Sheet 1: REB-ROB ---
  // คอลัมน์: เดือน/ปี | เลขที่ใบรับคืน | รายการสินค้า | จำนวน/คัน | ยอดเงินสุทธิ | ประเภท | สาเหตุ... | ชื่อลูกค้า
  const sheetRR = ss.getSheetByName('รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB');
  if (!sheetRR) {
    Logger.log('⚠️ ไม่พบ Sheet REB-ROB');
  } else {
    const rows = sheetRR.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    // หา index คอลัมน์
    const ci = {};
    headers.forEach((h, i) => { ci[h] = i; });

    // คอลัมน์อาจใช้ชื่อต่างกัน — map อัตโนมัติ
    const colMonthYear      = ci['เดือน/ปี']        ?? 0;
    const colBillNo         = ci['เลขที่ใบรับคืน']  ?? 1;
    const colAmount         = ci['ยอดเงินสุทธิ']    ?? 4;
    const colType           = ci['ประเภท']           ?? 5;
    const colCause          = headers.findIndex(h => h.includes('สาเหตุ'));  // G
    const colCustomer       = ci['ชื่อลูกค้า']       ?? 7;
    const colContractor     = ci['ชุดผู้รับเหมา']    ?? -1;
    const colEmployee       = ci['ชื่อพนักงาน']      ?? -1;

    let lastYM           = null;
    let lastBillNo       = '';
    let lastCode         = 'REB';
    let lastCustomer     = null;
    let lastCause        = null;
    let lastType         = null;
    let lastContractor   = null;
    let lastEmployee     = null;

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];

      const billCell = String(r[colBillNo] || '').trim();
      if (billCell) {
        lastBillNo = billCell;

        // ดึง code_type จากเลขที่บิล
        const codeMatch = billCell.match(/^(REB|ROB)/i);
        lastCode = codeMatch ? codeMatch[1].toUpperCase() : 'REB';

        // ดึง ปี/เดือน จากเลขที่บิล เช่น REB6904-0001 → 69=ปี 04=เดือน
        const ymFromBill = billCell.match(/(?:REB|ROB)(\d{2})(\d{2})-/i);
        if (ymFromBill) {
          lastYM = { year: Number('25' + ymFromBill[1]), month: Number(ymFromBill[2]) };
        }

        lastCustomer   = String(r[colCustomer]   || '').trim() || null;
        lastType       = String(r[colType]        || '').trim() || null;
        lastCause      = colCause >= 0 ? (String(r[colCause] || '').trim() || null) : null;
        lastContractor = colContractor >= 0 ? (String(r[colContractor] || '').trim() || null) : null;
        lastEmployee   = colEmployee >= 0 ? (String(r[colEmployee] || '').trim() || null) : null;
      }

      if (!lastYM || !lastBillNo) continue;

      const amount = toNum(r[colAmount]);
      if (amount === 0) continue;

      records.push({
        year:             lastYM.year,
        month:            lastYM.month,
        code_type:        lastCode,
        damage_group:     'รับคืน-เสียหายหน้างาน',
        customer_name:    lastCustomer,
        cause:            lastType || lastCause,
        reason:           lastCause,
        amount:           amount,
        bill_no:          lastBillNo,
        contractor_team:  lastContractor,
        employee_name:    lastEmployee,
      });
    }
    Logger.log(`REB-ROB: อ่านได้ ${records.length} แถว`);
  }

  // --- Sheet 2: ID ---
  // คอลัมน์: วันที่ | เลขที่บิล | รายการสินค้า | จำนวน/คัน | ยอดเงินสุทธิ | หักเงินผู้รับเหมา | ประเภทเสียหาย | สาเหตุ
  const countBefore = records.length;
  const sheetID = ss.getSheetByName('เสียหายในโรงงานบางเลน รหัสID');
  if (!sheetID) {
    Logger.log('⚠️ ไม่พบ Sheet ID');
  } else {
    const rows = sheetID.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim());
    const ci = {};
    headers.forEach((h, i) => { ci[h] = i; });

    const colBillNo      = ci['เลขที่บิล']         ?? 0;
    const colAmount      = ci['ยอดเงินสุทธิ']      ?? 3;
    const colType        = ci['ประเภทเสียหาย']     ?? 5;
    const colCause       = ci['สาเหตุ']            ?? 6;
    const colContractorID = ci['ชุดผู้รับเหมา']   ?? -1;
    const colEmployeeID   = ci['ชื่อพนักงาน']      ?? -1;

    let lastYM = null; // carry forward year/month

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];

      // ดึง ปี/เดือน จากเลขที่บิลก่อน (น่าเชื่อถือกว่า date cell)
      // รูปแบบ: ID2-6904-001 → 69=ปี, 04=เดือน
      const bill = String(r[colBillNo] || '').trim();
      if (bill) {
        const m = bill.match(/ID\d*-(\d{2})(\d{2})-/i);
        if (m) {
          lastYM = { year: Number('25' + m[1]), month: Number(m[2]) };
        }
      }

      if (!lastYM) continue;

      const amount = toNum(r[colAmount]);
      // รับทั้ง + และ - (บางแถวเป็นการปรับปรุง)
      if (amount === 0) continue;

      const typeVal = String(r[colType] || '').trim() || null;
      const cause   = String(r[colCause] || '').trim() || null;

      const contractorID = colContractorID >= 0 ? (String(r[colContractorID] || '').trim() || null) : null;
      const employeeID   = colEmployeeID   >= 0 ? (String(r[colEmployeeID]   || '').trim() || null) : null;

      records.push({
        year:             lastYM.year,
        month:            lastYM.month,
        code_type:        'ID',
        damage_group:     'เสียหายในโรงงาน',
        customer_name:    null,
        cause:            typeVal || cause,
        reason:           cause,
        amount:           amount,
        bill_no:          bill || null,
        contractor_team:  contractorID,
        employee_name:    employeeID,
      });
    }
    Logger.log(`ID: อ่านได้ ${records.length - countBefore} แถว`);
  }

  if (records.length === 0) { Logger.log('damage_items: ไม่มีข้อมูล'); return; }

  // ลบทั้งหมดก่อน แล้ว insert ใหม่
  truncateDamageItems();
  insertInBatches('damage_items', records, 200);
  Logger.log(`✅ damage_items: sync รวม ${records.length} แถว`);
}

// ============================================================
// เพิ่มคอลัมน์ ชุดผู้รับเหมา และ ชื่อพนักงาน ในชีทเสียหาย (ถ้ายังไม่มี)
// รันครั้งเดียวจาก Menu หรือ Script Editor
// ============================================================
function addPersonnelColumns(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SS_MAIN_ID);

  // REB-ROB: เฉพาะชื่อพนักงาน (blowcount) ลบ ชุดผู้รับเหมา ออกถ้ามี
  const sheetRR = ss.getSheetByName('รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB');
  if (sheetRR) {
    const headersRR = sheetRR.getRange(1, 1, 1, sheetRR.getLastColumn()).getValues()[0].map(h => String(h).trim());
    // ลบคอลัมน์ ชุดผู้รับเหมา ออก (ถ้ามี) — ต้องลบจากขวาไปซ้าย
    for (let i = headersRR.length - 1; i >= 0; i--) {
      if (headersRR[i] === 'ชุดผู้รับเหมา') {
        sheetRR.deleteColumn(i + 1);
        Logger.log('🗑️ REB-ROB: ลบคอลัมน์ "ชุดผู้รับเหมา" ออกแล้ว');
      }
    }
    // เพิ่ม ชื่อพนักงาน ถ้ายังไม่มี
    const headersRR2 = sheetRR.getRange(1, 1, 1, sheetRR.getLastColumn()).getValues()[0].map(h => String(h).trim());
    if (!headersRR2.includes('ชื่อพนักงาน')) {
      sheetRR.getRange(1, sheetRR.getLastColumn() + 1).setValue('ชื่อพนักงาน');
      Logger.log('✅ REB-ROB: เพิ่มคอลัมน์ "ชื่อพนักงาน"');
    }
  }

  // ID: ชุดผู้รับเหมา (I) + ชื่อพนักงาน (J)
  const sheetID = ss.getSheetByName('เสียหายในโรงงานบางเลน รหัสID');
  if (sheetID) {
    let headersID = sheetID.getRange(1, 1, 1, sheetID.getLastColumn()).getValues()[0].map(h => String(h).trim());
    let nextColID = sheetID.getLastColumn() + 1;

    if (!headersID.includes('ชุดผู้รับเหมา')) {
      sheetID.getRange(1, nextColID).setValue('ชุดผู้รับเหมา');
      Logger.log('✅ ID: เพิ่มคอลัมน์ "ชุดผู้รับเหมา"');
      nextColID++;
      headersID = sheetID.getRange(1, 1, 1, sheetID.getLastColumn()).getValues()[0].map(h => String(h).trim());
    }
    if (!headersID.includes('ชื่อพนักงาน')) {
      sheetID.getRange(1, nextColID).setValue('ชื่อพนักงาน');
      Logger.log('✅ ID: เพิ่มคอลัมน์ "ชื่อพนักงาน"');
    }

    // dropdown ชุดผู้รับเหมา เฉพาะแถวที่ G = ผู้รับเหมา
    setupContractorColumn(sheetID);
  }
}

// ============================================================
// Sync Sheet "ยอดขายประจำเดือน" → damage_sales
// คอลัมน์: เดือน/ปี | ยอดขาย/บาท  (เดือน/ปี เป็นชื่อเดือนไทย เช่น "เมษายน 2569")
// ============================================================
function syncDamageSales(ss) {
  const sheet = ss.getSheetByName('ยอดขายประจำเดือน');
  if (!sheet) { Logger.log('⚠️ ไม่พบ Sheet "ยอดขายประจำเดือน"'); return; }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return;

  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const ym = parseThaiMonthYear(r[0]);
    if (!ym) continue;
    const sales = toNum(r[1]);
    if (sales === 0) continue;
    records.push({ year: ym.year, month: ym.month, sales });
  }

  if (records.length === 0) { Logger.log('damage_sales: ไม่มีข้อมูล'); return; }

  const years = [...new Set(records.map(r => r.year))];
  years.forEach(y => deleteFromSupabase('damage_sales', `year=eq.${y}`));
  upsertToSupabase('damage_sales', records, 'year,month');
  Logger.log(`✅ damage_sales: sync ${records.length} แถว`);
}

// ============================================================
// สร้าง 3 Tab เสียหายใน Spreadsheet การผลิต (รันครั้งเดียว)
// เลือก setupDamageDetailSheets → Run
// ============================================================
function setupDamageDetailSheets() {
  const ss = SpreadsheetApp.openById(SS_MAIN_ID);

  // ลบ Sheet เก่าที่ชื่อผิด (ถ้ามี)
  ['รายการเสียหาย','ยอดขาย-รายเดือน'].forEach(name => {
    const old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  });

  function mk(name, headers, sampleRows, color) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    else sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(color).setFontColor('#fff').setFontWeight('bold')
      .setHorizontalAlignment('center');
    if (sampleRows.length) {
      sh.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
    }
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('✅ สร้าง Sheet: ' + name);
  }

  // Tab 1: ยอดขายประจำเดือน
  mk('ยอดขายประจำเดือน',
    ['เดือน/ปี', 'ยอดขาย/บาท'],
    [
      ['มกราคม 2569', 17345153.30],
      ['กุมภาพันธ์ 2569', 19459745.47],
      ['มีนาคม 2569', 21730581.59],
      ['เมษายน 2569', 11730756.37],
      ['พฤษภาคม 2569', 17601538.50],
    ],
    '#38761d'
  );

  // Tab 2: REB-ROB
  mk('รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB',
    ['เดือน/ปี','เลขที่ใบรับคืน','รายการสินค้า','จำนวน/คัน','ยอดเงินสุทธิ','ประเภท','สาเหตุ/รูปคืน-เสียหน้างาน','ชื่อลูกค้า','หน่วยงานของลูกค้า'],
    [
      ['เมษายน 2569','REB6904-0001','เสาเข็ม I-0.26x9.00 เมตร',1,3500,'เสาร้าวในกอง/ขนส่ง/ปีกแตก','','บจก.ตัวอย่าง','บางเลน'],
      ['เมษายน 2569','ROB6904-0001','เสาเข็ม I-0.22x7.00 เมตร',2,5040,'เสาเข็ม Fail','','คุณตัวอย่าง','อยุธยา'],
    ],
    '#b45f06'
  );

  // Tab 3: ID — คอลัมน์ A = เดือน/ปี (สูตรอัตโนมัติ), B = เลขที่บิล
  const shID = (()=>{
    const name = 'เสียหายในโรงงานบางเลน รหัสID';
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    else sh.clearContents();

    const headers = ['เดือน/ปี','เลขที่บิล','รายการสินค้า','จำนวน/คัน','ยอดเงินสุทธิ','หักเงินผู้รับเหมา','ประเภทเสียหาย','สาเหตุ'];
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground('#1155cc').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');

    // ข้อมูลตัวอย่างแถว 2
    sh.getRange(2,2,1,7).setValues([['ID2-6904-001','เสาเข็ม I-0.18x7.00 เมตร',1,1861.80,'','อื่นๆ','ปรับปรุงสินค้า']]);

    // สูตร เดือน/ปี ใน A2:A500 — ดึงจากเลขที่บิล (คอลัมน์ B)
    const formulas = [];
    for(let row=2;row<=500;row++){
      formulas.push(['=IFERROR(IF(B'+row+'="","",CHOOSE(VALUE(MID(REGEXEXTRACT(B'+row+',"\\d{4}"),3,2)),"มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม")&" 25"&LEFT(REGEXEXTRACT(B'+row+',"\\d{4}"),2)),"")']);
    }
    sh.getRange(2,1,499,1).setFormulas(formulas);

    sh.setFrozenRows(1);
    sh.autoResizeColumns(1,headers.length);
    Logger.log('✅ สร้าง Sheet: ' + name);
    return sh;
  })();

  // Tab 2 REB-ROB: เพิ่มสูตร เดือน/ปี อัตโนมัติใน col A (ดึงจาก col B = เลขที่ใบรับคืน)
  const shRR = ss.getSheetByName('รับคืนสินค้า-เสียหายหน้างาน บางเลน รหัสREB-ROB');
  if (shRR) {
    const fmlsRR = [];
    for(let row=2;row<=500;row++){
      fmlsRR.push(['=IFERROR(IF(B'+row+'="","",CHOOSE(VALUE(MID(REGEXEXTRACT(B'+row+',"\\d{4}"),3,2)),"มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม")&" 25"&LEFT(REGEXEXTRACT(B'+row+',"\\d{4}"),2)),"")']);
    }
    shRR.getRange(2,1,499,1).setFormulas(fmlsRR);
    Logger.log('✅ ใส่สูตร เดือน/ปี ใน REB-ROB Sheet');
  }

  try {
    SpreadsheetApp.getUi().alert('✅ สร้าง Sheets สำเร็จ!\nคอลัมน์ A "เดือน/ปี" คำนวณจากเลขที่บิลอัตโนมัติ');
  } catch(e) {
    Logger.log('✅ setupDamageDetailSheets สำเร็จ');
  }
}

// ============================================================
// sync งานรอผลิต → Supabase pending_work
// ============================================================
function syncPendingWork(ss) {
  var sheet = ss.getSheetByName('งานรอผลิต');
  if (!sheet) { Logger.log('⚠️ ไม่พบ Sheet "งานรอผลิต" — ข้าม'); return; }

  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) { Logger.log('⚠️ งานรอผลิต: ไม่มีข้อมูล'); return; }

  var data = rows.slice(1).filter(function(r){ return r[0] && r[1] !== ''; }).map(function(r) {
    var d = r[0];
    var dateStr = d instanceof Date
      ? Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd')
      : String(d).slice(0, 10);
    return { updated_date: dateStr, value_m3: Number(r[1]) || null };
  });

  deleteFromSupabase('pending_work', 'id=gte.1');
  if (data.length > 0) insertInBatches('pending_work', data, 200);
  Logger.log('✅ pending_work: ' + data.length + ' แถว');
}
