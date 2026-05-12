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
    sheet.appendRow(['Timestamp', 'Date', 'ShopName', 'ItemsDetail', 'TotalAmount', 'RecordedBy']);
    sheet.getRange("A1:F1").setFontWeight("bold");
  }
}