// ใส่ ID ของ Google Sheets (ตัวแปรนี้จะถูกเรียกใช้ได้จากทุกไฟล์ .gs)
const SHEET_ID = '1yOIGTO-bqiwXmLU-ULKxtrGTs2ikWibnkB_tH31xzxY';

// ฟังก์ชันตั้งค่าตาราง (เลือกรันฟังก์ชันนี้จากหน้า Editor แค่ครั้งแรกครั้งเดียว)
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // สร้างตาราง Transactions ถ้ายังไม่มี
  if (!ss.getSheetByName('Transactions')) {
    const sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['Timestamp', 'Date', 'Type', 'Description', 'Amount', 'RecordedBy']);
    sheet.getRange("A1:F1").setFontWeight("bold");
  }
  
  // สร้างตาราง Receipts ถ้ายังไม่มี
  if (!ss.getSheetByName('Receipts')) {
    const sheet = ss.insertSheet('Receipts');
    sheet.appendRow(['Timestamp', 'Date', 'ShopName', 'ItemsDetail', 'TotalAmount', 'RecordedBy', 'ReceiptID']);
    sheet.getRange("A1:G1").setFontWeight("bold");
  }

  // สร้างตาราง DB_ระบบบันทึกรายรับรายจ่าย (ผู้ใช้งาน) ถ้ายังไม่มี
  if (!ss.getSheetByName('DB_ระบบบันทึกรายรับรายจ่าย')) {
    const sheet = ss.insertSheet('DB_ระบบบันทึกรายรับรายจ่าย');
    sheet.appendRow(['Username', 'Password', 'Role', 'Name']);
    sheet.getRange("A1:D1").setFontWeight("bold");
    
    // สร้างผู้ใช้เริ่มต้น (Admin) รหัสผ่านคือ admin
    const defaultPassword = generateSecurePassword('admin');
    sheet.appendRow(['admin', defaultPassword, 'admin', 'System Admin']);
  }
}