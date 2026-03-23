// ============================================================
//  ระบบน้ำประปาหมู่บ้าน — Google Apps Script v2
//  รองรับ auto-sync ทุกเดือน + เปิดได้ทุกเครื่อง
// ============================================================

const SPREADSHEET_ID = ''; // ← ใส่ ID ของ Google Sheet ถ้าจะใช้ ID แทน active sheet
                             //   ถ้าว่างไว้จะใช้ SpreadsheetApp.getActiveSpreadsheet()

// ชื่อ Sheet format: "น้ำ_มกราคม_2568"
const SHEET_PREFIX = 'น้ำ_';

// ===== MAIN ENTRY POINTS =====

function doGet(e) {
  const params = e ? e.parameter : {};
  try {
    // listSheets — ดึงรายชื่อ Sheet ทั้งหมดที่มีข้อมูล
    if (params.action === 'listSheets') {
      return jsonResponse(listAllSheets());
    }
    // getResidents — ดึงทะเบียนบ้านจาก Sheet "ทะเบียนบ้าน"
    if (params.action === 'getResidents') {
      return jsonResponse(getResidentsData());
    }
    // test connection
    if (params.test) {
      return jsonResponse({ success: true, message: 'เชื่อมต่อสำเร็จ ✅', timestamp: new Date().toISOString() });
    }
    // ดึงข้อมูลเดือนที่ระบุ
    if (params.month && params.year) {
      return jsonResponse(getSheetData(params.month, parseInt(params.year)));
    }
    return jsonResponse({ success: false, error: 'ระบุ month และ year' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    // บันทึกทะเบียนบ้าน
    if (payload.action === 'saveResidents' && payload.residents) {
      writeResidentsData(payload.residents);
      return jsonResponse({ success: true, message: `บันทึกทะเบียนบ้าน ${payload.residents.length} หลังสำเร็จ` });
    }
    if (payload.rows && payload.month && payload.year) {
      writeSheetData(payload.month, payload.year, payload.rows);
      return jsonResponse({ success: true, message: `บันทึก ${payload.rows.length} รายการสำเร็จ` });
    }
    return jsonResponse({ success: false, error: 'ข้อมูลไม่ครบ' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ===== RESIDENTS SHEET (ทะเบียนบ้าน) =====

const RESIDENTS_SHEET_NAME = 'ทะเบียนบ้าน';

function getResidentsData() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(RESIDENTS_SHEET_NAME);
  if (!sheet) return { success: true, data: [], message: 'ไม่พบ Sheet ทะเบียนบ้าน' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const values = range.getValues();
  const data = values
    .filter(row => row[0] && row[1]) // must have house + name
    .map(row => ({
      house:        String(row[0]).trim(),
      name:         String(row[1]).trim(),
      phone:        String(row[2] || '').trim(),
      initialMeter: parseFloat(row[3]) || 0,
      status:       String(row[4] || 'active').trim() || 'active'
    }));
  return { success: true, data, total: data.length };
}

function writeResidentsData(residents) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(RESIDENTS_SHEET_NAME);
  if (!sheet) {
    // หา Sheet แรก (index 0) และเปลี่ยนชื่อ ถ้าชื่อไม่ใช่ ทะเบียนบ้าน
    const allSheets = ss.getSheets();
    if (allSheets[0].getName() !== RESIDENTS_SHEET_NAME) {
      // สร้าง Sheet ใหม่ที่ตำแหน่ง 0
      sheet = ss.insertSheet(RESIDENTS_SHEET_NAME, 0);
    } else {
      sheet = allSheets[0];
    }
    setupResidentsHeader(sheet);
  }

  // Clear existing data rows (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();

  if (!residents.length) return;

  const data = residents.map(r => [r.house, r.name, r.phone || '', r.initialMeter || 0, r.status || 'active']);
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  formatResidentsSheet(sheet, data.length);
}

function setupResidentsHeader(sheet) {
  const headers = ['บ้านเลขที่', 'ชื่อ-สกุล', 'เบอร์โทร', 'มิเตอร์เริ่มต้น', 'สถานะ'];
  const header = sheet.getRange(1, 1, 1, 5);
  header.setValues([headers]);
  header.setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 100);
}

function formatResidentsSheet(sheet, dataRows) {
  if (!dataRows) return;
  for (let i = 0; i < dataRows; i++) {
    sheet.getRange(i + 2, 1, 1, 5).setBackground(i % 2 === 0 ? '#f5f8fc' : '#ffffff');
  }
  sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  sheet.getRange(2, 1, dataRows, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 5, dataRows, 1).setHorizontalAlignment('center');
}

// ===== LIST ALL SHEETS =====

function listAllSheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const result = [];
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(SHEET_PREFIX)) {
      // parse "น้ำ_มกราคม_2568" → { month: "มกราคม", year: 2568 }
      const parts = name.replace(SHEET_PREFIX, '').split('_');
      if (parts.length >= 2) {
        const month = parts[0];
        const year = parseInt(parts[1]);
        if (year > 2500 && month) {
          result.push({ month, year, sheetName: name });
        }
      }
    }
  });
  // Sort newest first
  result.sort((a, b) => b.year !== a.year ? b.year - a.year : MONTHS_TH.indexOf(b.month) - MONTHS_TH.indexOf(a.month));
  return { success: true, sheets: result, total: result.length };
}

// ===== GET SHEET DATA =====

function getSheetData(month, year) {
  const ss = getSpreadsheet();
  const sheetName = `${SHEET_PREFIX}${month}_${year}`;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [], message: `ไม่พบ Sheet: ${sheetName}` };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, data: [] };

  const range = sheet.getRange(2, 1, lastRow - 1, 8);
  const values = range.getValues();
  const data = values
    .filter(row => row[0] && row[1]) // must have house + name
    .map(row => ({
      seq:    row[0],
      house:  String(row[1]).trim(),
      name:   String(row[2]).trim(),
      prev:   parseFloat(row[3]) || 0,
      curr:   parseFloat(row[4]) || 0,
      units:  parseFloat(row[5]) || 0,
      amount: parseFloat(row[6]) || 0,
      remark: String(row[7] || '').trim()
    }));
  return { success: true, data, month, year, total: data.length };
}

// ===== WRITE SHEET DATA =====

function writeSheetData(month, year, rows) {
  const ss = getSpreadsheet();
  const sheetName = `${SHEET_PREFIX}${month}_${year}`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheetHeader(sheet);
  }

  // Clear existing data rows (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();

  if (!rows.length) return;

  const data = rows.map(r => [r.seq, r.house, r.name, r.prev, r.curr, r.units, r.amount, r.remark || '']);
  sheet.getRange(2, 1, data.length, 8).setValues(data);
  formatSheet(sheet, data.length);

  // Update yearly summary
  updateYearlySummary(ss, month, year, rows);
}

// ===== HELPERS =====

function setupSheetHeader(sheet) {
  const headers = ['ลำดับ', 'บ้านเลขที่', 'ชื่อ-สกุล', 'เลขมิเตอร์ครั้งก่อน', 'เลขมิเตอร์ครั้งหลัง', 'จำนวนหน่วย', 'จำนวนเงิน (บาท)', 'หมายเหตุ'];
  const header = sheet.getRange(1, 1, 1, 8);
  header.setValues([headers]);
  header.setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(8, 180);
}

function formatSheet(sheet, dataRows) {
  if (!dataRows) return;
  // Alternate row colors
  for (let i = 0; i < dataRows; i++) {
    const row = sheet.getRange(i + 2, 1, 1, 8);
    row.setBackground(i % 2 === 0 ? '#f5f8fc' : '#ffffff');
  }
  // Amount column — right align, number format
  sheet.getRange(2, 7, dataRows, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  sheet.getRange(2, 4, dataRows, 3).setNumberFormat('#,##0').setHorizontalAlignment('center');
  sheet.getRange(2, 1, dataRows, 1).setHorizontalAlignment('center');
  // Total row
  const totalRow = dataRows + 2;
  sheet.getRange(totalRow, 1, 1, 8).setBackground('#e3f0ff').setFontWeight('bold');
  sheet.getRange(totalRow, 3).setValue('รวม');
  sheet.getRange(totalRow, 6).setFormula(`=SUM(F2:F${dataRows + 1})`);
  sheet.getRange(totalRow, 7).setFormula(`=SUM(G2:G${dataRows + 1})`).setNumberFormat('#,##0.00');
}

function updateYearlySummary(ss, month, year, rows) {
  try {
    const summaryName = `สรุปรายปี_${year}`;
    let summary = ss.getSheetByName(summaryName);
    if (!summary) {
      summary = ss.insertSheet(summaryName);
      summary.getRange(1, 1, 1, 4).setValues([['เดือน', 'จำนวนบ้าน', 'หน่วยรวม', 'รายรับรวม']]);
      summary.getRange(1, 1, 1, 4).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    }
    const totalUnits = rows.reduce((s, r) => s + (parseFloat(r.units) || 0), 0);
    const totalAmount = rows.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0);
    // Find existing row for this month
    const lastRow = summary.getLastRow();
    let found = false;
    if (lastRow > 1) {
      const existing = summary.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existing.length; i++) {
        if (existing[i][0] === month) {
          summary.getRange(i + 2, 1, 1, 4).setValues([[month, rows.length, totalUnits, totalAmount]]);
          found = true; break;
        }
      }
    }
    if (!found) summary.appendRow([month, rows.length, totalUnits, totalAmount]);
    summary.getRange(2, 4, summary.getLastRow() - 1, 1).setNumberFormat('#,##0.00');
  } catch(e) { /* non-critical */ }
}

function getSpreadsheet() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

const MONTHS_TH = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

// ===== SHEET MENU (เปิดใน Google Sheets) =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💧 ระบบน้ำประปา')
    .addItem('สร้าง Sheet เดือนนี้', 'createCurrentMonthSheet')
    .addItem('ดู URL สำหรับเว็บ', 'showWebAppUrl')
    .addToUi();
}

function createCurrentMonthSheet() {
  const now = new Date();
  const month = MONTHS_TH[now.getMonth() + 1];
  const year = now.getFullYear() + 543;
  const ss = getSpreadsheet();
  const name = `${SHEET_PREFIX}${month}_${year}`;
  if (!ss.getSheetByName(name)) {
    const sheet = ss.insertSheet(name);
    setupSheetHeader(sheet);
    SpreadsheetApp.getUi().alert(`✅ สร้าง Sheet "${name}" สำเร็จ`);
  } else {
    SpreadsheetApp.getUi().alert(`Sheet "${name}" มีอยู่แล้ว`);
  }
}

function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(`URL สำหรับเว็บ:\n\n${url}\n\nนำ URL นี้ไปใส่ในหน้า ตั้งค่า > Google Sheets Sync`);
}
