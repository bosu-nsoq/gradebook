/**
 * ============================================================
 * แอปพลิเคชัน: Gradebook (ระบบบันทึกคะแนนนักเรียน)
 * สำหรับ: ครูบอส และคณะครู
 * สภาพแวดล้อม: Google Apps Script + Google Sheets Web App
 * รองรับ: Role-based Access (Admin / User), PIN Login, การพิมพ์คะแนน
 * ============================================================
 */

function doGet(e) {
  // ตรวจสอบหากมีพารามิเตอร์ action (สำหรับเรียกผ่าน API จาก GitHub Pages หรือภายนอก)
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function')
               ? globalThis[action]
               : (typeof this[action] === 'function' ? this[action] : null);

      if (!fn) {
        return ContentService.createTextOutput(JSON.stringify({ error: "ไม่พบฟังก์ชัน: " + action }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var result = fn.apply(null, args);
      return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // โหมดแสดงผลหน้าเว็บปกติ (Google Apps Script Web App)
  var html;
  try {
    html = HtmlService.createHtmlOutputFromFile('index');
  } catch(err) {
    html = HtmlService.createHtmlOutputFromFile('Index');
  }

  return html
      .setTitle('Gradebook')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0, viewport-fit=cover');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var args = data.args || [];
    var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function')
             ? globalThis[action]
             : (typeof this[action] === 'function' ? this[action] : null);

    if (!fn) {
      return ContentService.createTextOutput(JSON.stringify({ error: "ไม่พบฟังก์ชัน: " + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var result = fn.apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 1. ระบบจัดการข้อมูลผู้ใช้งาน (Users Sheet & Authentication)
// ============================================================

function getOrCreateUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    // หัวตารางข้อมูลผู้ใช้
    sheet.appendRow(["UserID", "Name", "Subject", "PIN", "Role", "CreatedAt"]);
    // ข้อมูลเริ่มต้น: ครูบอส (Admin) และผู้ใช้ตัวอย่าง (User)
    sheet.appendRow(["ADM01", "ครูบอส", "ภาษาไทย", "8055", "admin", new Date()]);
    sheet.appendRow(["TCH01", "ครูสมหมาย มีสุข", "คณิตศาสตร์", "1234", "user", new Date()]);
    
    // จัดรูปแบบหัวตาราง
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight("bold")
      .setBackground("#6c5ce7")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
  }
  return sheet;
}

// ตรวจสอบการเข้าสู่ระบบด้วย PIN
function loginUserByPin(pin) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var cleanPin = String(pin).trim();

  // ตรวจสอบผู้ใช้ในตาราง (ข้ามแถวที่ 1 ซึ่งเป็น Header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var userPin = String(row[3]).trim();
    if (userPin === cleanPin) {
      return {
        success: true,
        user: {
          userId: String(row[0]),
          name: String(row[1]),
          subject: String(row[2]),
          role: String(row[4]).toLowerCase() // 'admin' หรือ 'user'
        }
      };
    }
  }

  // รองรับกรณีฉุกเฉิน: หากไม่มีข้อมูลแต่ใส่ 8055 ให้สิทธิ์ Admin ครูบอส
  if (cleanPin === "8055") {
    return {
      success: true,
      user: {
        userId: "ADM01",
        name: "ครูบอส",
        subject: "ภาษาไทย",
        role: "admin"
      }
    };
  }

  return { success: false, message: "รหัส PIN ไม่ถูกต้อง" };
}

// ดึงรายชื่อผู้ใช้ทั้งหมด (สำหรับโหมด Admin)
function getUsers() {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var users = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;
    users.push({
      row: i + 1,
      userId: String(row[0]),
      name: String(row[1]),
      subject: String(row[2]),
      pin: String(row[3]),
      role: String(row[4] || 'user').toLowerCase()
    });
  }
  return users;
}

// บันทึก/แก้ไขข้อมูลผู้ใช้ (สำหรับโหมด Admin)
function saveUser(userData) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var userId = String(userData.userId).trim();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === userId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > 0) {
    // แก้ไขข้อมูลเดิม
    sheet.getRange(rowIndex, 2).setValue(userData.name);
    sheet.getRange(rowIndex, 3).setValue(userData.subject);
    sheet.getRange(rowIndex, 4).setValue(String(userData.pin));
    sheet.getRange(rowIndex, 5).setValue(userData.role || "user");
  } else {
    // เพิ่มผู้ใช้ใหม่
    sheet.appendRow([
      userId,
      userData.name,
      userData.subject,
      String(userData.pin),
      userData.role || "user",
      new Date()
    ]);
  }
  return "Success";
}

// ลบผู้ใช้ (สำหรับโหมด Admin)
function deleteUser(userId) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var cleanId = String(userId).trim();

  // ป้องกันการลบ Admin หลัก (ADM01)
  if (cleanId === "ADM01") return "CannotDeleteMasterAdmin";

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanId) {
      sheet.deleteRow(i + 1);
      return "Success";
    }
  }
  return "NotFound";
}

// ============================================================
// 2. ระบบจัดการห้องเรียนและคะแนน (Classrooms & Scores)
// ============================================================

// ดึงรายชื่อชั้นเรียนทั้งหมด (ยกเว้นชีตระบบ)
function getAllClassNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var systemSheets = ["Config", "Users", "System", "Settings", "Log"];
  
  var names = sheets.map(function(sheet) {
    return sheet.getName();
  }).filter(function(name) {
    return systemSheets.indexOf(name) === -1;
  });
  return names;
}

// สร้างชั้นเรียนใหม่
function createNewClass(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(className)) return "Duplicate"; // ชื่อซ้ำ

  var sheet = ss.insertSheet(className);
  // สร้างหัวตารางเริ่มต้น
  sheet.appendRow(["No", "Name", "คะแนนเก็บ (10)", "สอบกลางภาค (20)"]);
  // สร้างนักเรียนตัวอย่าง
  sheet.appendRow([1, "ด.ช. ตัวอย่าง รักเรียน", "", ""]);
  sheet.appendRow([2, "ด.ญ. มานี มีตา", "", ""]);
  
  return "Success";
}

// ดึงข้อมูลนักเรียนและคะแนนของห้องเรียน
function getClassData(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: [], students: [] };

  // อ่าน Header (เริ่ม Col C = index 2)
  var headers = [];
  for (var i = 2; i < data[0].length; i++) {
    var h = data[0][i];
    var max = 100;
    var match = String(h).match(/\((\d+)\)/);
    if (match) max = parseInt(match[1]);
    headers.push({ name: String(h), colIndex: i, max: max });
  }

  // อ่าน Students
  var students = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // ข้ามแถวที่ไม่มีชื่อหรือเลขที่
    if (row[0] === "" || row[1] === "") continue;

    var scores = {};
    headers.forEach(function(h) {
      scores[h.name] = row[h.colIndex];
    });

    students.push({
      row: i + 1,
      no: row[0],
      name: row[1],
      scores: scores
    });
  }
  
  // เรียงตามเลขที่
  students.sort(function(a, b) { 
    return Number(a.no) - Number(b.no); 
  });

  return { headers: headers, students: students };
}

// บันทึกคะแนน
function saveScore(className, row, colIndex, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return "SheetNotFound";
  sheet.getRange(row, colIndex + 1).setValue(value);
  return "Success";
}

// เพิ่มช่องคะแนนใหม่
function addHeader(className, headerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return "SheetNotFound";
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue(headerName);
  return "Success";
}

// เพิ่มนักเรียนใหม่เข้าห้องเรียน
function addStudent(className, no, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return "SheetNotFound";

  var lastCol = sheet.getLastColumn();
  if (lastCol < 2) lastCol = 4;
  
  var newRow = new Array(lastCol).fill("");
  newRow[0] = Number(no);
  newRow[1] = String(name).trim();

  sheet.appendRow(newRow);
  return "Success";
}

// ดึงข้อมูลภาพรวมเริ่มต้นแบบรวมศูนย์ (One-shot Background Sync)
function getInitialAppData(pin) {
  var auth = loginUserByPin(pin);
  if (!auth.success) return { success: false, message: auth.message };

  var classNames = getAllClassNames();
  var currentClassData = null;
  if (classNames.length > 0) {
    currentClassData = getClassData(classNames[0]);
  }

  var users = [];
  if (auth.user.role === 'admin') {
    users = getUsers();
  }

  return {
    success: true,
    user: auth.user,
    classNames: classNames,
    firstClass: classNames.length > 0 ? classNames[0] : "",
    currentClassData: currentClassData,
    users: users
  };
}
