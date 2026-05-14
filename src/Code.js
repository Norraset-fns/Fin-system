function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("ระบบบันทึกรายรับรายจ่าย และ ใบเสร็จ")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ระบบจะรู้จักตัวแปร SHEET_ID ที่ดึงมาจาก setup.gs อัตโนมัติ
// ==========================================
// ระบบเข้าสู่ระบบ (Authentication) - ปรับปรุงเพื่อประสิทธิภาพ (Performance)
// ==========================================

function authenticate(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย"); 
    const lastRow = sheet.getLastRow();
    
    // ถ้าไม่มีข้อมูลเลย ให้ตอบกลับว่าไม่พบผู้ใช้
    if (lastRow <= 1) return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };

    // ⚡ เทคนิค Senior: ดึงมาเฉพาะคอลัมน์ Username (คอลัมน์ A) เพื่อหาบรรทัดที่ต้องการ
    // ช่วยให้ประหยัดแรมมากกว่าการดึงทั้งแผ่น (getDataRange)
    const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i][0] === username) {
        rowIndex = i + 2; // +2 เพราะเริ่มบรรทัด 2 และ Array Index เริ่มที่ 0
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };
    }

    // เมื่อเจอบรรทัดที่ใช่แล้ว ค่อยดึงข้อมูลคอลัมน์อื่นๆ (B, C, D) ของแถวนั้นมาใช้งาน
    const userData = sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
    const storedPasswordString = userData[1].toString();
    const role = userData[2];
    const name = userData[3];

    // 1. ตรวจสอบรหัสผ่านระบบ Hash+Salt
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
      // 2. กรณีรหัสผ่านธรรมดา (เผื่อไว้ช่วงย้ายระบบ)
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

// ==========================================
// ระบบจัดการผู้ใช้งาน (User Management)
// ==========================================

// 1. ดึงรายชื่อผู้ใช้ (พร้อมเช็กสิทธิ์ Admin ที่ฝั่ง Server)
function getUsers(currentUserRole) {
  // 🔐 ป้องกันคนแอบเรียกฟังก์ชันผ่านหน้าเว็บโดยไม่ได้รับอนุญาต
  if (currentUserRole !== "admin") {
    return { success: false, message: "สิทธิ์ของคุณไม่เพียงพอในการดูรายชื่อผู้ใช้" };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];

  // ดึงเฉพาะข้อมูลที่จำเป็นไปโชว์ที่หน้าตาราง
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

// 2. เพิ่มผู้ใช้ใหม่ (พร้อมระบบ Lock ป้องกันข้อมูลพัง)
function addUser(formObj) {
  const lock = LockService.getScriptLock(); // 🔐 สร้างกุญแจล็อค
  try {
    lock.waitLock(30000); // รอคิวว่าง 30 วินาที ป้องกันบันทึกชนกัน

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
    const lastRow = sheet.getLastRow();

    if (!formObj.username || !formObj.password || !formObj.name) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // เช็ก Username ซ้ำ (ดึงแค่คอลัมน์แรกมาเช็ก จะเร็วกว่าดึงทั้งชีต)
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
    lock.releaseLock(); // 🔓 ปลดล็อคให้คนต่อไปเข้ามาใช้งานได้ (ต้องอยู่ใน finally เสมอ)
  }
}

// 3. อัปเดตข้อมูลผู้ใช้ (พร้อมระบบ Lock)
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

// ==========================================
// ระบบบันทึกรายการ (Transactions & Receipts)
// ==========================================

// 1. บันทึกรายรับ-รายจ่าย (พร้อมระบบ Lock ป้องกันการออกเลข ID ซ้ำ)
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
      return { success: false, message: "ไม่สามารถบันทึกรายการล่วงหน้าได้" };
    }

    const amount = parseFloat(formObj.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: "จำนวนเงินต้องมากกว่า 0 บาท" };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const timestamp = new Date();

    // ⚡ getNextId ต้องรันอยู่ข้างใน Lock เพื่อไม่ให้คนอื่นมาแย่งเลข ID เดียวกัน
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

    return { success: true, message: "บันทึกสำเร็จ (ID: " + txId + ")" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 2. บันทึกใบเสร็จ (พร้อมระบบ Lock)
function saveReceipt(formObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // --- ด่านตรวจของใบเสร็จ (Validation) ---
    if (!formObj.date || !formObj.shopName) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    const inputDate = new Date(formObj.date);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (inputDate > today) {
      return { success: false, message: "ไม่สามารถบันทึกใบเสร็จล่วงหน้าได้" };
    }

    const items = JSON.parse(formObj.items);
    if (items.length === 0) {
      return { success: false, message: "ต้องมีสินค้าอย่างน้อย 1 รายการ" };
    }

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

    // บันทึกลงหน้า Receipts (สำหรับเก็บรายละเอียดสินค้า)
    receiptSheet.appendRow([
      timestamp,
      formObj.date,
      formObj.shopName,
      formObj.items,
      total,
      formObj.username,
      rcId,
    ]);

    // บันทึกลงหน้า Transactions (สำหรับสรุปยอดบัญชี)
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

    return { success: true, message: "บันทึกใบเสร็จสำเร็จ (ID: " + rcId + ")" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 3. ดึงข้อมูล Dashboard (พร้อมระบบกรองข้อมูล)
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

    // 📅 ระบบกรองเวลา (Time Filter)
    if (row[1]) {
      const rowDate = new Date(row[1]);

      if (filterType === "this_month") {
        if (
          rowDate.getMonth() !== currentMonth ||
          rowDate.getFullYear() !== currentYear
        )
          continue;
      } else if (filterType === "last_month") {
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
    }

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
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return prefix + "-1000";
  }

  const lastId = sheet.getRange(lastRow, 8).getValue().toString();
  const parts = lastId.split("-");

  if (parts.length === 2) {
    const nextNumber = parseInt(parts[1]) + 1;
    return prefix + "-" + nextNumber;
  } else {
    return prefix + "-" + (lastRow + 1000);
  }
}

// ==========================================
// ระบบยกเลิกและแก้ไขรายการ (Void & Edit)
// ==========================================

// 1. ยกเลิกรายการ (Void) - ปรับยอดเงินเป็น 0 และใส่เครื่องหมาย [ยกเลิก]
function voidTransactionInSheet(txId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "ไม่พบข้อมูลในระบบ" };

    // ⚡ Surgical Read: ดึงมาแค่คอลัมน์ ID (คอลัมน์ H) เพื่อหาตำแหน่งบรรทัด
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

// 2. แก้ไขรายการ (Edit) - เขียนทับข้อมูลเดิม
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

// 3. ดึงประวัติรายการทั้งหมด (พร้อมระบบ Filter และ Security)
function getTransactionsHistory(username, role, searchQuery) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getDataRange().getValues();
    let results = [];
    const query = searchQuery ? searchQuery.toLowerCase() : "";

    // ⚡ เทคนิค Junior-to-Senior: วนลูปย้อนกลับ (i--) เพื่อเอาตัวล่าสุดขึ้นก่อน
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowUsername = row[5];
      const rowDesc = row[3].toString().toLowerCase();
      const rowId = row[7].toString().toLowerCase();
      const rowDate = row[1].toString();

      // 1. เช็กสิทธิ์ (Row-level Security): ถ้าไม่ใช่ Admin/Manager ให้ดูได้แค่ของตัวเอง
      if (role === "user" && rowUsername !== username) continue;

      // 2. เช็กคำค้นหา (ถ้ากรอกมา): ค้นหาใน รายการ, รหัส, หรือ วันที่
      if (query && !rowDesc.includes(query) && !rowId.includes(query) && !rowDate.includes(query)) {
        continue;
      }

      results.push({
        timestamp: row[0] ? row[0].toString() : "",
        date: row[1] ? row[1].toString() : "",
        type: row[2],
        description: row[3],
        amount: parseFloat(row[4]) || 0,
        user: row[5],
        itemsData: row[6] ? row[6].toString() : "",
        id: row[7] ? row[7].toString() : ""
      });

      // ⚡ ประสิทธิภาพ: จำกัดไว้แค่ 100 รายการ เพื่อไม่ให้แอปอืด
      if (results.length >= 100) break;
    }

    return results;
  } catch (error) {
    console.error("getTransactionsHistory error:", error.toString());
    return [];
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
