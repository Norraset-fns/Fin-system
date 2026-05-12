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
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย"); // อย่าลืมเช็กชื่อแท็บให้ตรงกับของคุณนะครับ
    const data = sheet.getDataRange().getValues();

    // วนลูปหา Username
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        const storedPasswordString = data[i][1].toString(); // ดึงข้อมูลก้อน salt$hash มา

        // 1. ตรวจสอบว่าข้อมูลในชีตเป็นระบบ Hash+Salt (มีเครื่องหมาย $) หรือยัง
        if (storedPasswordString.includes("$")) {
          // หั่นแยกเกลือกับแฮชออกจากกัน
          const parts = storedPasswordString.split("$");
          const salt = parts[0];
          const storedHash = parts[1];

          // นำรหัสผ่านที่กรอกมาหน้าเว็บ บวกกับเกลือ แล้วปั่นเป็น Hash เพื่อเทียบกัน
          const rawHash = Utilities.computeDigest(
            Utilities.DigestAlgorithm.SHA_256,
            password + salt,
          );
          const computedHash = rawHash
            .map(function (byte) {
              return ("0" + (byte & 0xff).toString(16)).slice(-2);
            })
            .join("");

          // ถ้า Hash ตรงกันเป๊ะ แปลว่ารหัสถูกต้อง!
          if (computedHash === storedHash) {
            return {
              success: true,
              data: {
                username: data[i][0],
                role: data[i][2],
                name: data[i][3],
              },
            };
          } else {
            return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
          }
        } else {
          // 2. เผื่อกรณีฉุกเฉิน (มีบางบัญชียังเป็นข้อความธรรมดา ยังไม่ได้แปลง)
          if (password === storedPasswordString) {
            return {
              success: true,
              data: {
                username: data[i][0],
                role: data[i][2],
                name: data[i][3],
              },
            };
          } else {
            return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
          }
        }
      }
    }

    // ถ้าวนลูปจนจบแล้วยังไม่เจอ Username
    return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 2. ฟังก์ชันเพิ่มผู้ใช้ใหม่ (Create)
// ==========================================
// ระบบจัดการผู้ใช้งาน (User Management)
// ==========================================

// 1. ฟังก์ชันดึงรายชื่อผู้ใช้ทั้งหมด (Read)
function getUsers() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");

  // ดึงข้อมูลตั้งแต่บรรทัดที่ 2 ถึงบรรทัดสุดท้าย คอลัมน์ที่ 1-4
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

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

// 2. ฟังก์ชันเพิ่มผู้ใช้ใหม่ (Create)
function addUser(formObj) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");

    // --- ด่านตรวจ (Validation) ---
    if (!formObj.username || !formObj.password || !formObj.name) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    // เช็กว่า Username นี้มีคนใช้ไปหรือยัง?
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === formObj.username) {
        return {
          success: false,
          message: "Username นี้มีในระบบแล้ว กรุณาใช้ชื่ออื่น",
        };
      }
    }

    // 🔐 เปลี่ยนมาใช้ฟังก์ชันแบบมี Salt
    const securePassword = generateSecurePassword(formObj.password);

    // ถ้าผ่านด่านทั้งหมด ก็บันทึกคนใหม่ลงบรรทัดสุดท้าย
    sheet.appendRow([
      formObj.username,
      securePassword, // สิ่งที่จะถูกบันทึกจะเป็นรูปแบบ (salt$hash)
      formObj.role,
      formObj.name,
    ]);

    return { success: true, message: "เพิ่มผู้ใช้งานสำเร็จ!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 3. ฟังก์ชันอัปเดตข้อมูลผู้ใช้งาน (Update)
function updateUser(formObj) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DB_ระบบบันทึกรายรับรายจ่าย");
    const data = sheet.getDataRange().getValues();

    // วนลูปหา Username ที่ต้องการแก้ไข (Username อยู่คอลัมน์ที่ 1 หรือ index 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === formObj.username) {
        const rowNum = i + 1;

        // อัปเดต ชื่อ และ ตำแหน่ง
        sheet.getRange(rowNum, 3).setValue(formObj.role); // คอลัมน์ C: Role
        sheet.getRange(rowNum, 4).setValue(formObj.name); // คอลัมน์ D: Name

        // ถ้ามีการกรอกรหัสผ่านใหม่เข้ามา ให้ทำการเข้ารหัสแล้วบันทึกทับ
        if (formObj.password && formObj.password.trim() !== "") {
          const securePassword = generateSecurePassword(formObj.password);
          sheet.getRange(rowNum, 2).setValue(securePassword); // คอลัมน์ B: Password
        }

        return {
          success: true,
          message: "อัปเดตข้อมูลคุณ " + formObj.name + " เรียบร้อยแล้ว",
        };
      }
    }
    return { success: false, message: "ไม่พบชื่อผู้ใช้งานในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
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
  try {
    // --- เริ่มด่านตรวจคนเข้าเมือง (Validation) ---
    // 1. เช็กว่าส่งวันที่มาไหม?
    if (!formObj.date) {
      return { success: false, message: "กรุณาระบุวันที่ทำรายการ" };
    }

    // 1.1 เช็กว่าเป็นวันที่ในอนาคตหรือไม่? (NEW!)
    const inputDate = new Date(formObj.date); // แปลงวันที่ที่กรอกมาเป็น Date Object
    const today = new Date(); // ดึงเวลาของวันนี้
    today.setHours(23, 59, 59, 999); // ปรับเวลาของวันนี้ให้เป็นเที่ยงคืนสุดๆ จะได้เทียบแค่วันที่

    if (inputDate > today) {
      return {
        success: false,
        message: "ไม่สามารถบันทึกรายการล่วงหน้า (วันที่ในอนาคต) ได้",
      };
    }

    // 2. เช็กว่ายอดเงินเป็นตัวเลขที่มากกว่า 0 หรือเปล่า? (ป้องกันส่งยอดเงิน 0 หรือยอดติดลบ)
    const amount = parseFloat(formObj.amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, message: "จำนวนเงินต้องมากกว่า 0 บาท" };
    }
    // --- จบด่านตรวจ ---

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const timestamp = new Date();

    // 1. เรียกใช้งานฟังก์ชันผู้ช่วย เพื่อสร้าง ID ใหม่ (กำหนดอักษรนำหน้าเป็น 'TX')
    const txId = getNextId("Transactions", "TX");

    // 2. บันทึกข้อมูลลง Sheet โดยเพิ่ม txId ไว้ที่ช่องสุดท้าย (คอลัมน์ที่ 8)
    sheet.appendRow([
      timestamp,
      formObj.date,
      formObj.type,
      formObj.description,
      parseFloat(formObj.amount),
      formObj.username,
      "", // คอลัมน์ที่ 7: รายการสินค้า (เว้นว่างไว้สำหรับ Transaction ปกติ)
      txId, // คอลัมน์ที่ 8: ID เอกสาร (เพิ่มเข้ามาใหม่!)
    ]);

    return {
      success: true,
      message: "บันทึกข้อมูลสำเร็จ (รหัสอ้างอิง: " + txId + ")",
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 4. บันทึกใบเสร็จ (พ่วง ID และ Validation)
function saveReceipt(formObj) {
  try {
    // --- ด่านตรวจของใบเสร็จ (Validation) ---
    if (!formObj.date || !formObj.shopName) {
      return {
        success: false,
        message: "กรุณาระบุวันที่และชื่อร้านค้าให้ครบถ้วน",
      };
    }

    // เช็กว่าเป็นวันที่ในอนาคตหรือไม่? (NEW!)
    const inputDate = new Date(formObj.date);
    const today = new Date();
    today.setHours(23, 59, 59, 999);

    if (inputDate > today) {
      return {
        success: false,
        message: "ไม่สามารถบันทึกใบเสร็จล่วงหน้า (วันที่ในอนาคต) ได้",
      };
    }

    // ประกาศ items ครั้งที่ 1 (และครั้งเดียว)
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

    // สร้าง ID ใหม่ 2 ตัว
    const rcId = getNextId("Receipts", "RC");
    const txId = getNextId("Transactions", "TX");

    // คำนวณยอดรวม (ใช้ตัวแปร items จากด้านบนได้เลย ไม่ต้องประกาศ const ซ้ำ)
    let total = 0;
    items.forEach((item) => {
      total += parseFloat(item.price) * parseInt(item.qty);
    });

    // บันทึกลงหน้า Receipts
    receiptSheet.appendRow([
      timestamp,
      formObj.date,
      formObj.shopName,
      formObj.items,
      total,
      formObj.username,
      rcId, // ID ใบเสร็จ
    ]);

    // บันทึกลงหน้า Transactions
    txSheet.appendRow([
      timestamp,
      formObj.date,
      "expense",
      "ใบเสร็จ: " + formObj.shopName + " (อ้างอิง " + rcId + ")",
      total,
      formObj.username,
      formObj.items,
      txId, // ID รายจ่าย
    ]);

    return {
      success: true,
      message: "บันทึกใบเสร็จสำเร็จ (รหัส: " + rcId + ")",
    };
  } catch (error) {
    return { success: false, message: error.toString() };
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
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const data = sheet.getDataRange().getValues();

    // วนลูปหา ID ที่ต้องการ (ID อยู่คอลัมน์ที่ 8 หรือ index 7)
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === txId) {
        // ตรวจสอบก่อนว่ารายการนี้เคยถูกยกเลิกไปแล้วหรือยัง
        if (data[i][3].toString().includes("[ยกเลิก]")) {
          return { success: false, message: "รายการนี้ถูกยกเลิกไปแล้ว" };
        }

        const rowNum = i + 1; // บรรทัดที่ต้องการแก้ในชีต

        // 1. ปรับยอดเงินเป็น 0 (คอลัมน์ที่ 5 หรือ index 4)
        sheet.getRange(rowNum, 5).setValue(0);

        // 2. เพิ่มคำว่า [ยกเลิก] ไว้หน้าคำอธิบาย (คอลัมน์ที่ 4 หรือ index 3)
        const oldDesc = data[i][3];
        sheet.getRange(rowNum, 4).setValue("[ยกเลิก] " + oldDesc);

        return {
          success: true,
          message: "ยกเลิกรายการ " + txId + " เรียบร้อยแล้ว",
        };
      }
    }
    return { success: false, message: "ไม่พบรหัสรายการในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ฟังก์ชันสำหรับอัปเดตข้อมูล (Edit)
function updateTransactionInSheet(formObj) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("Transactions");
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === formObj.id) {
        // หา ID ในคอลัมน์ที่ 8 (index 7)
        const rowNum = i + 1;

        // เขียนทับลงไปในช่องที่ต้องการ
        sheet.getRange(rowNum, 2).setValue(formObj.date); // คอลัมน์ B: วันที่
        sheet.getRange(rowNum, 4).setValue(formObj.description); // คอลัมน์ D: รายการ
        sheet.getRange(rowNum, 5).setValue(parseFloat(formObj.amount)); // คอลัมน์ E: จำนวนเงิน

        return {
          success: true,
          message: "อัปเดตรายการ " + formObj.id + " เรียบร้อยแล้ว",
        };
      }
    }
    return { success: false, message: "ไม่พบรหัสรายการในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
