function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("ระบบบันทึกรายรับรายจ่าย และ ใบเสร็จ")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ระบบจะรู้จักตัวแปร SHEET_ID ที่ดึงมาจาก setup.gs อัตโนมัติ
// ==========================================
// ระบบเข้าสู่ระบบ (Authentication)
// ==========================================

function authenticate(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย"); 
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };

    // ดึงเฉพาะคอลัมน์ A (Username) มาตรวจสอบก่อนเพื่อความเร็ว
    const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i][0] === username) {
        rowIndex = i + 2; // +2 เพราะเริ่มต้นที่บรรทัด 2 และ Index ของ Array เริ่มที่ 0
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };
    }

    // เมื่อเจอ Username แล้ว ค่อยไปดึงข้อมูลที่เหลือของแถวนั้น (B, C, D)
    const userData = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
    const storedPasswordString = userData[1].toString();
    const role = userData[2];
    const name = userData[3];

    // 1. ตรวจสอบว่าข้อมูลในชีตเป็นระบบ Hash+Salt (มีเครื่องหมาย $) หรือยัง
    if (storedPasswordString.includes("$")) {
      const parts = storedPasswordString.split("$");
      const salt = parts[0];
      const storedHash = parts[1];

      const rawHash = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        password + salt,
      );
      const computedHash = rawHash
        .map(function (byte) {
          return ("0" + (byte & 0xff).toString(16)).slice(-2);
        })
        .join("");

      if (computedHash === storedHash) {
        return {
          success: true,
          data: { username: username, role: role, name: name },
        };
      } else {
        return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
      }
    } else {
      // 2. กรณีรหัสผ่านธรรมดา
      if (password === storedPasswordString) {
        return {
          success: true,
          data: { username: username, role: role, name: name },
        };
      } else {
        return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
      }
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 2. ฟังก์ชันเพิ่มผู้ใช้ใหม่ (Create)
// ==========================================
// ระบบจัดการผู้ใช้งาน (User Management)
// ==========================================

// 1. ฟังก์ชันดึงรายชื่อผู้ใช้ทั้งหมด (Read) - เพิ่มการเช็กสิทธิ์ Admin
function getUsers(currentUserRole) {
  if (currentUserRole !== "admin") {
    return { success: false, message: "สิทธิ์ของคุณไม่เพียงพอในการดูรายชื่อผู้ใช้" };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];

  // ดึงเฉพาะคอลัมน์ A (Username), C (Role), D (Name)
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  let usersList = [];
  data.forEach((row) => {
    if (row[0]) {
      usersList.push({
        username: row[0],
        role: row[2],
        name: row[3],
      });
    }
  });
  return usersList;
}

// 2. ฟังก์ชันเพิ่มผู้ใช้ใหม่ (Create) - เพิ่ม LockService
function addUser(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // รอคิว 30 วินาที

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
    const lastRow = sheet.getLastRow();

    if (!formObj.username || !formObj.password || !formObj.name) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // เช็กว่า Username นี้มีคนใช้ไปหรือยัง? (ดึงเฉพาะคอลัมน์แรกมาเช็ก)
    if (lastRow > 1) {
      const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < usernames.length; i++) {
        if (usernames[i][0] === formObj.username) {
          return { success: false, message: "Username นี้มีในระบบแล้ว" };
        }
      }
    }

    const securePassword = generateSecurePassword(formObj.password);

    sheet.appendRow([
      formObj.username,
      securePassword,
      formObj.role,
      formObj.name,
    ]);

    return { success: true, message: "เพิ่มผู้ใช้งานสำเร็จ!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 3. ฟังก์ชันอัปเดตข้อมูลผู้ใช้งาน (Update) - เพิ่ม LockService
function updateUser(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "ไม่พบข้อมูลในระบบ" };

    const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i][0] === formObj.username) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, 3).setValue(formObj.role);
        sheet.getRange(rowNum, 4).setValue(formObj.name);

        if (formObj.password && formObj.password.trim() !== "") {
          const securePassword = generateSecurePassword(formObj.password);
          sheet.getRange(rowNum, 2).setValue(securePassword);
        }
        return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
      }
    }
    return { success: false, message: "ไม่พบชื่อผู้ใช้งานในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ระบบรักษาความปลอดภัย (Security: Hash + Salt)
// ==========================================

// ฟังก์ชันสร้าง Hash + Salt พร้อมกัน
function generateSecurePassword(plainPassword) {
  // 1. สร้าง Salt เป็นตัวอักษรสุ่มยาวๆ (UUID)
  const salt = Utilities.getUuid();

  // 2. เอารหัสผ่านดิบ มาบวกกับ Salt แล้วค่อยเอาไป Hash
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plainPassword + salt,
  );
  const hash = rawHash
    .map(function (byte) {
      return ("0" + (byte & 0xff).toString(16)).slice(-2);
    })
    .join("");

  // 3. ส่งค่ากลับไปแบบแพ็คคู่ (Salt $ Hash) เพื่อเอาไปบันทึกลงช่องเดียวกัน
  return salt + "$" + hash;
}

// ⚠️ ฟังก์ชันใช้ครั้งเดียว: รันเพื่อแปลงรหัสผ่านเก่าในชีตให้เป็น Hash + Salt
function migrateOldPasswords() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("ชีต1"); // อย่าลืมเช็กชื่อแท็บให้ตรงกันนะครับ
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let currentPassword = data[i][1].toString();

    // เช็กว่ารหัสนี้มี Salt หรือยัง (ดูจากการมีเครื่องหมาย $ ซ่อนอยู่ไหม)
    if (currentPassword && !currentPassword.includes("$")) {
      let secureValue = generateSecurePassword(currentPassword);
      // อัปเดตกลับไปที่ชีตในคอลัมน์ที่ 2
      sheet.getRange(i + 1, 2).setValue(secureValue);
    }
  }
}

// 3. บันทึกรายรับ-รายจ่าย (เพิ่มช่องว่างไว้สำหรับคอลัมน์ที่ 7)
function saveTransaction(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // --- เริ่มด่านตรวจคนเข้าเมือง (Validation) ---
    if (!formObj.date) {
      return { success: false, message: "กรุณาระบุวันที่ทำรายการ" };
    }

    const inputDate = new Date(formObj.date);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (inputDate > today) {
      return {
        success: false,
        message: "ไม่สามารถบันทึกรายการล่วงหน้า (วันที่ในอนาคต) ได้",
      };
    }

    const amount = parseFloat(formObj.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: "จำนวนเงินต้องมากกว่า 0 บาท" };
    }
    // --- จบด่านตรวจ ---

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const timestamp = new Date();

    const txId = getNextId("Transactions", "TX");

    sheet.appendRow([
      timestamp,
      formObj.date,
      formObj.type,
      formObj.description,
      amount,
      formObj.username,
      "", 
      txId,
    ]);

    return {
      success: true,
      message: "บันทึกข้อมูลสำเร็จ (รหัสอ้างอิง: " + txId + ")",
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 4. บันทึกใบเสร็จ (พ่วง ID และ Validation)
function saveReceipt(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // --- ด่านตรวจของใบเสร็จ (Validation) ---
    if (!formObj.date || !formObj.shopName) {
      return {
        success: false,
        message: "กรุณาระบุวันที่และชื่อร้านค้าให้ครบถ้วน",
      };
    }

    const inputDate = new Date(formObj.date);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (inputDate > today) {
      return {
        success: false,
        message: "ไม่สามารถบันทึกใบเสร็จล่วงหน้า (วันที่ในอนาคต) ได้",
      };
    }

    const items = JSON.parse(formObj.items);
    if (items.length === 0) {
      return {
        success: false,
        message: "ใบเสร็จต้องมีสินค้าอย่างน้อย 1 รายการ",
      };
    }
    // --- จบด่านตรวจ ---

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const receiptSheet = ss.getSheetByName("Receipts");
    const txSheet = ss.getSheetByName("Transactions");
    const timestamp = new Date();

    const rcId = getNextId("Receipts", "RC");
    const txId = getNextId("Transactions", "TX");

    let total = 0;
    items.forEach((item) => {
      total += parseFloat(item.price) * parseInt(item.qty);
    });

    receiptSheet.appendRow([
      timestamp,
      formObj.date,
      formObj.shopName,
      formObj.items,
      total,
      formObj.username,
      rcId,
    ]);

    txSheet.appendRow([
      timestamp,
      formObj.date,
      "expense",
      "ใบเสร็จ: " + formObj.shopName + " (อ้างอิง " + rcId + ")",
      total,
      formObj.username,
      formObj.items,
      txId,
    ]);

    return {
      success: true,
      message: "บันทึกใบเสร็จสำเร็จ (รหัส: " + rcId + ")",
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 5. ดึงข้อมูล Dashboard (เพิ่ม filterType มารับค่า)
function getDashboardData(username, role, filterType) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const txSheet = ss.getSheetByName("Transactions");
  if (!txSheet)
    return {
      balance: 0,
      income: [0],
      expense: [0],
      labels: ["รวม"],
      recentItems: [],
    };
  const data = txSheet.getDataRange().getValues();

  let balance = 0,
    incomeSum = 0,
    expenseSum = 0,
    recentItems = [];

  // 🛠️ สร้างตัวแปรดึงเวลา ณ ปัจจุบัน เอาไว้เทียบ
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let type = row[2];
    let amount = parseFloat(row[4]) || 0;
    const rowUsername = row[5];

    // 🛡️ Row-Level Security
    if (role === "user" && rowUsername !== username) {
      continue;
    }

    // ==========================================
    // 📅 ด่านกรองเวลา (Time Filter)
    // ==========================================
    if (row[1]) {
      // เช็กก่อนว่ามีข้อมูลวันที่ไหม
      const rowDate = new Date(row[1]); // แปลงวันที่ในชีตให้เป็นก้อนเวลาที่ JS รู้จัก

      if (filterType === "this_month") {
        // ถ้าเดือนไม่ตรง หรือ ปีไม่ตรง ให้ข้ามไปเลย!
        if (
          rowDate.getMonth() !== currentMonth ||
          rowDate.getFullYear() !== currentYear
        )
          continue;
      } else if (filterType === "last_month") {
        // ทริค Junior: ต้องระวังเดือนมกราคม (0) เพราะเดือนที่แล้วจะเป็นธันวาคม (11) ของปีที่แล้ว
        let targetMonth = currentMonth - 1;
        let targetYear = currentYear;
        if (targetMonth < 0) {
          targetMonth = 11;
          targetYear--;
        }
        if (
          rowDate.getMonth() !== targetMonth ||
          rowDate.getFullYear() !== targetYear
        )
          continue;
      } else if (filterType === "this_year") {
        if (rowDate.getFullYear() !== currentYear) continue;
      }
      // ถ้าเป็น 'all' ระบบก็จะข้ามด่านนี้ไปเลย (ไม่โดนดัก)
    }

    // ==========================================
    // ส่วนคำนวณยอดเงิน (ถ้าหลุดรอดด่านบนมาได้ ก็จะมาถูกบวกตรงนี้)
    // ==========================================
    if (type === "income") {
      incomeSum += amount;
      balance += amount;
    } else if (type === "expense") {
      expenseSum += amount;
      balance -= amount;
    }

    if (recentItems.length < 10) {
      recentItems.push({
        timestamp: row[0] ? row[0].toString() : "",
        date: row[1] ? row[1].toString() : "",
        type: type,
        description: row[3] || "-",
        amount: amount,
        user: row[5] || "ระบบ",
        itemsData: row[6] ? row[6].toString() : "",
        id: row[7] ? row[7].toString() : "",
      });
    }
  }
  return {
    balance: balance,
    income: [incomeSum],
    expense: [expenseSum],
    labels: ["รวม"],
    recentItems: recentItems,
  };
}

// ฟังก์ชันผู้ช่วย: สร้าง ID รันเลขตัวต่อไปอัตโนมัติ (เช่น TX-1001)
function getNextId(sheetName, prefix) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = sheet.getLastRow(); // หาเบอร์บรรทัดสุดท้ายที่มีข้อมูล

  // ถ้ามีแค่บรรทัดที่ 1 (หัวตาราง) แสดงว่ายังไม่มีข้อมูล ให้เริ่มที่ 1000
  if (lastRow <= 1) {
    return prefix + "-1000";
  }

  // ดึงค่า ID จากคอลัมน์ที่ 8 (คอลัมน์ H) ของบรรทัดสุดท้ายมาดู
  const lastId = sheet.getRange(lastRow, 8).getValue().toString();

  // สมมติ lastId คือ "TX-1005" -> เราจะหั่นด้วยเครื่องหมาย "-" แล้วเอาเลข 1005 มา + 1
  const parts = lastId.split("-");

  if (parts.length === 2) {
    const nextNumber = parseInt(parts[1]) + 1; // แปลงเป็นตัวเลขแล้วบวก 1
    return prefix + "-" + nextNumber; // ประกอบร่างกลับไปเป็น TX-1006
  } else {
    // กันเหนียว เผื่อข้อมูลเก่าไม่มี ID ให้เอาเลขบรรทัดมาใช้แทนก่อน
    return prefix + "-" + (lastRow + 1000);
  }
}

// ฟังก์ชันสำหรับยกเลิกรายการ (Void)
function voidTransactionInSheet(txId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "ไม่พบข้อมูลในระบบ" };

    // ดึงเฉพาะคอลัมน์ H (ID) เพื่อหาบรรทัด
    const ids = sheet.getRange(1, 8, lastRow, 1).getValues();

    for (let i = 1; i < ids.length; i++) {
      if (ids[i][0] === txId) {
        const rowNum = i + 1;
        const currentDesc = sheet.getRange(rowNum, 4).getValue();

        if (currentDesc.toString().includes("[ยกเลิก]")) {
          return { success: false, message: "รายการนี้ถูกยกเลิกไปแล้ว" };
        }

        sheet.getRange(rowNum, 5).setValue(0);
        sheet.getRange(rowNum, 4).setValue("[ยกเลิก] " + currentDesc);

        return { success: true, message: "ยกเลิกรายการเรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบรหัสรายการในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ฟังก์ชันสำหรับอัปเดตข้อมูล (Edit)
function updateTransactionInSheet(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "ไม่พบข้อมูลในระบบ" };

    const ids = sheet.getRange(1, 8, lastRow, 1).getValues();

    for (let i = 1; i < ids.length; i++) {
      if (ids[i][0] === formObj.id) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(formObj.date);
        sheet.getRange(rowNum, 4).setValue(formObj.description);
        sheet.getRange(rowNum, 5).setValue(parseFloat(formObj.amount));

        return { success: true, message: "อัปเดตรายการเรียบร้อยแล้ว" };
      }
    }
    return { success: false, message: "ไม่พบรหัสรายการในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
