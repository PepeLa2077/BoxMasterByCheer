/**
 * ═══════════════════════════════════════════════════════════
 *  BangYai BoxMaster by Cheer — Google Sheets Backend
 * ═══════════════════════════════════════════════════════════
 *  วิธีติดตั้ง (ทำครั้งเดียว):
 *  1. สร้าง Google Sheet ใหม่ (ตั้งชื่อเช่น "BangYai BoxMaster DB")
 *  2. เมนู ส่วนขยาย (Extensions) → Apps Script
 *  3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ → กดบันทึก (ไอคอนแผ่นดิสก์)
 *  4. กด Deploy → New deployment → เฟือง → เลือก "Web app"
 *     - Execute as:      Me (บัญชีของคุณ)
 *     - Who has access:  Anyone
 *     → กด Deploy → Authorize access (อนุญาตสิทธิ์บัญชีตัวเอง)
 *  5. คัดลอก Web app URL (ลงท้ายด้วย /exec)
 *     ไปวางในแอป: หน้า "จัดการ" → เชื่อมต่อ Google Sheets → กดเชื่อมต่อ
 *
 *  หมายเหตุ: ถ้าแก้โค้ดนี้ภายหลัง ต้อง Deploy → Manage deployments
 *  → แก้ไข (ดินสอ) → Version: New version → Deploy ถึงจะมีผล
 * ═══════════════════════════════════════════════════════════
 */

var DATA_SHEET = '_data';     // ชีตซ่อนเก็บข้อมูลดิบ (อย่าแก้เอง)
var BACKUP_SHEET = '_backup'; // สำรองข้อมูลรายวัน เก็บย้อนหลัง 30 วัน
var BACKUP_KEEP = 30;

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'load';
  if (action === 'load') return jsonOut(loadData());
  return jsonOut({ ok: false, error: 'unknown action' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save') {
      saveData(body.data);
      return jsonOut({ ok: true, updatedAt: body.data.updatedAt || null });
    }
    return jsonOut({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function loadData() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sh) return { ok: true, data: null };
  var raw = sh.getRange(1, 1).getValue();
  if (!raw) return { ok: true, data: null };
  try { return { ok: true, data: JSON.parse(raw) }; }
  catch (e) { return { ok: true, data: null }; }
}

function saveData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // กันเขียนชนกันเมื่อหลายเครื่องบันทึกพร้อมกัน
  try {
    var sh = ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET);
    dailyBackup(ss, sh); // สำรองข้อมูลเดิมของวันนี้ไว้ก่อนเขียนทับ
    sh.getRange(1, 1).setValue(JSON.stringify(data));
    try { sh.hideSheet(); } catch (e) {}
    mirrorReadable(ss, data);
  } finally {
    lock.releaseLock();
  }
}

/** สำรองข้อมูลวันละ 1 snapshot (ค่าแรกของวัน) — ย้อนหลัง 30 วัน
 *  กู้คืน: เปิดชีต _backup (คลิกขวาชื่อชีต → Unhide ถ้าซ่อนอยู่),
 *  คัดลอก JSON ของวันที่ต้องการไปวางทับช่อง A1 ของชีต _data */
function dailyBackup(ss, dataSh) {
  var raw = dataSh.getRange(1, 1).getValue();
  if (!raw) return;
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var bk = ss.getSheetByName(BACKUP_SHEET) || ss.insertSheet(BACKUP_SHEET);
  try { bk.hideSheet(); } catch (e) {}
  var last = bk.getLastRow();
  if (last > 0 && bk.getRange(last, 1).getValue() === today) return; // วันนี้สำรองแล้ว
  bk.appendRow([today, raw]);
  var rows = bk.getLastRow();
  if (rows > BACKUP_KEEP) bk.deleteRows(1, rows - BACKUP_KEEP);
}

/** สร้างชีตอ่านง่ายสำหรับคน: Lot ยา / กล่อง / ประวัติ (เขียนทับทุกครั้งที่บันทึก) */
function mirrorReadable(ss, data) {
  writeTab(ss, 'Lot ยา',
    ['ยา', 'Lot', 'วันหมดอายุ'],
    (data.lots || []).map(function (l) { return [l.drug, l.lotNo, l.expiry]; }));

  writeTab(ss, 'กล่อง',
    ['กล่อง', 'อยู่ที่', 'วันหมดอายุกล่อง'],
    (data.boxes || []).map(function (b) { return [b.no, b.ward || 'ห้องยา', b.expiry]; }));

  writeTab(ss, 'ประวัติการแลกเปลี่ยน',
    ['วันที่นำมาเปลี่ยน', 'กล่อง', 'Ward', 'เภสัชกร', 'ยาที่ใช้ไป', 'ปัญหาที่พบ', 'หมายเหตุ', 'วันหมดอายุกล่องใหม่'],
    (data.exchanges || []).map(function (x) {
      var items = (x.items || []).map(function (i) {
        return i.drug + ' (Lot ' + i.lotNo + ') x' + i.qty;
      }).join(' ; ');
      return [x.date, x.boxNo, x.ward, x.pharmacist, items, x.problem, x.note || '', x.newExpiry];
    }));
}

function writeTab(ss, name, head, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
}
