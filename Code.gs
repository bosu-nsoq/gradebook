/**
 * ============================================================
 * แอปพลิเคชัน: Gradebook (ระบบบันทึกคะแนนนักเรียน)
 * พัฒนาสำหรับ: ครูบอส และคณะครู
 * สภาพแวดล้อม: Google Apps Script + Google Sheets Web App
 * คุณสมบัติ:
 * 1. ข้อมูลนักเรียนส่วนกลางในชีต Students (ชื่อ, ชั้น, เลขที่)
 * 2. แยกช่องเก็บคะแนนของผู้ใช้แต่ละคนอย่างเด็ดขาด (Class_UserID)
 * 3. อัปเดตข้อมูลอัตโนมัติ และรองรับ API Mode
 * ============================================================
 */

function doGet(e) {
  // ตรวจสอบพารามิเตอร์ action สำหรับเรียกผ่าน API จาก GitHub Pages
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

  // โหมด Web App ปกติ
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
// 1. ระบบจัดการนักเรียนส่วนกลาง (ชีต Students: ชื่อ, ชั้น, เลขที่)
// ============================================================

function getOrCreateStudentsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Students");
  if (!sheet) {
    sheet = ss.insertSheet("Students");
  }

  var data = sheet.getDataRange().getValues();
  var hasStudents = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && String(data[i][1]).trim() !== "") {
      hasStudents = true;
      break;
    }
  }

  // หากไม่มีข้อมูลนักเรียน ให้สร้างตัวอย่าง
  if (!hasStudents) {
    sheet.clear();
    // หัวตาราง: เลขที่, ชื่อ, ชั้น
    sheet.appendRow(["เลขที่", "ชื่อ", "ชั้น"]);
    
    // ข้อมูลเริ่มต้นตัวอย่างหลายชั้นเรียน
    var sampleStudents = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ป.5/1"],
      [2, "ด.ญ. มานี มีตา", "ป.5/1"],
      [3, "ด.ช. ชูใจ ใจดี", "ป.5/1"],
      [4, "ด.ญ. ปิติ สุขใจ", "ป.5/1"],
      [5, "ด.ช. วีระ มานะ", "ป.5/1"],
      [1, "ด.ช. สมศักดิ์ สดใส", "ป.5/2"],
      [2, "ด.ญ. สมใจ มั่นคง", "ป.5/2"],
      [1, "ด.ช. เก่งกล้า ท้าทาย", "ป.6/1"],
      [2, "ด.ญ. สวยสม นิยมไทย", "ป.6/1"]
    ];
    sampleStudents.forEach(function(row) {
      sheet.appendRow(row);
    });

    var headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight("bold")
      .setBackground("#6c5ce7")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    SpreadsheetApp.flush();
  }
  return sheet;
}

// ดึงรายชื่อชั้นเรียนทั้งหมดจากคอลัมน์ "ชั้น" ในชีต Students
function getAllClassNames() {
  var sheet = getOrCreateStudentsSheet();
  var data = sheet.getDataRange().getValues();
  var classesSet = {};

  for (var i = 1; i < data.length; i++) {
    var cls = String(data[i][2]).trim();
    if (cls) {
      classesSet[cls] = true;
    }
  }

  var classes = Object.keys(classesSet);
  classes.sort();
  if (classes.length === 0) {
    classes = ["ป.5/1", "ป.5/2", "ป.6/1"];
  }
  return classes;
}

// ดึงรายชื่อนักเรียนเฉพาะชั้นเรียนที่เลือก
function getStudentsByClass(className) {
  var sheet = getOrCreateStudentsSheet();
  var data = sheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var no = data[i][0];
    var name = String(data[i][1]).trim();
    var cls = String(data[i][2]).trim();

    if (cls === className && name) {
      students.push({
        no: Number(no),
        name: name,
        className: cls
      });
    }
  }

  // เรียงตามเลขที่
  students.sort(function(a, b) {
    return a.no - b.no;
  });
  return students;
}

// เพิ่มนักเรียนใหม่เข้าสู่ชีต Students (ชื่อ, ชั้น, เลขที่)
function addStudent(className, no, name) {
  var sheet = getOrCreateStudentsSheet();
  sheet.appendRow([Number(no), String(name).trim(), String(className).trim()]);
  return "Success";
}

// เพิ่มรายชื่อนักเรียนทีละหลายคนพร้อมกัน (Bulk Import)
function addMultipleStudents(className, studentList) {
  var sheet = getOrCreateStudentsSheet();
  var cls = String(className).trim();
  
  if (Array.isArray(studentList)) {
    studentList.forEach(function(s) {
      if (s.name && String(s.name).trim()) {
        sheet.appendRow([Number(s.no), String(s.name).trim(), cls]);
      }
    });
  }
  return "Success";
}

// ============================================================
// 2. ระบบจัดการผู้ใช้ (ชีต Users)
// ============================================================

function getOrCreateUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
  }
  if (sheet.getLastRow() < 2) {
    sheet.clear();
    sheet.appendRow(["UserID", "Name", "Subject", "PIN", "Role", "CreatedAt"]);
    sheet.appendRow(["ADM01", "ครูบอส", "ภาษาไทย", "8055", "admin", new Date()]);
    sheet.appendRow(["TCH01", "ครูสมหมาย มีสุข", "คณิตศาสตร์", "1234", "user", new Date()]);
    sheet.appendRow(["TCH02", "ครูสมศรี มณีฉาย", "วิทยาศาสตร์", "5678", "user", new Date()]);
    sheet.appendRow(["TCH03", "ครูวีระ รักชาติ", "สังคมศึกษา", "9999", "user", new Date()]);
  }
  formatUsersSheet(sheet);
  return sheet;
}

function formatUsersSheet(sheet) {
  if (!sheet) return;
  try {
    var headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight("bold")
      .setBackground("#6c5ce7")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");

    // กำหนดฟอร์แมตคอลัมน์ D (PIN) ให้เป็น Plain Text เพื่อไม่ให้ตัดเลขศูนย์นำหน้า
    sheet.getRange("D2:D100").setNumberFormat("@");

    // กำหนด Data Validation Dropdown ให้คอลัมน์ E (Role) เลือกเฉพาะ 'admin' หรือ 'user'
    var roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["admin", "user"], true)
      .setAllowInvalid(false)
      .setHelpText("กรุณาระบุ Role เป็น 'admin' หรือ 'user' เท่านั้น")
      .build();
    sheet.getRange("E2:E100").setDataValidation(roleRule);

    sheet.setColumnWidth(1, 95);  // UserID
    sheet.setColumnWidth(2, 180); // Name
    sheet.setColumnWidth(3, 140); // Subject
    sheet.setColumnWidth(4, 90);  // PIN
    sheet.setColumnWidth(5, 105); // Role
    sheet.setColumnWidth(6, 160); // CreatedAt
    SpreadsheetApp.flush();
  } catch(e) {}
}

function loginUserByPin(pin) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var cleanPin = String(pin).trim();

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
          role: String(row[4]).toLowerCase()
        }
      };
    }
  }

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
    sheet.getRange(rowIndex, 2).setValue(userData.name);
    sheet.getRange(rowIndex, 3).setValue(userData.subject);
    sheet.getRange(rowIndex, 4).setValue(String(userData.pin));
    sheet.getRange(rowIndex, 5).setValue(userData.role || "user");
  } else {
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

function deleteUser(userId) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var cleanId = String(userId).trim();

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
// 3. ระบบแยกคะแนนรายบุคคลของผู้ใช้แต่ละคน (User-isolated Score Sheets)
//    ชื่อชีตคะแนน: [ClassName]_[UserID] เช่น ป.5/1_ADM01
// ============================================================

function getUserScoreSheetName(className, userId) {
  var uId = (userId && String(userId).trim()) ? String(userId).trim() : "ADM01";
  return className + "_" + uId;
}

// ดึงข้อมูลคะแนนเฉพาะของ User คนนั้นในห้องเรียนนั้น
function getClassData(className, userId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = getUserScoreSheetName(className, userId);
  var sheet = ss.getSheetByName(sheetName);

  // ดึงรายชื่อนักเรียนปัจจุบันในชั้นเรียนนี้จากชีต Students
  var studentRoster = getStudentsByClass(className);

  // หากยังไม่มีชีตคะแนนของผู้ใช้นี้ ให้สร้างชีตใหม่พร้อมหัวข้อเริ่มต้น
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var defaultHeader = ["No", "Name", "Class", "คะแนนเก็บ (10)", "สอบกลางภาค (20)"];
    sheet.appendRow(defaultHeader);
    
    // ใส่นักเรียนจาก Roster
    studentRoster.forEach(function(s) {
      sheet.appendRow([s.no, s.name, s.className, "", ""]);
    });

    sheet.getRange(1, 1, 1, defaultHeader.length)
      .setFontWeight("bold")
      .setBackground("#e8eaed");
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.flush();
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return { headers: [], students: [] };

  // อ่าน Header (เริ่ม Col D = index 3)
  var headers = [];
  for (var c = 3; c < data[0].length; c++) {
    var h = String(data[0][c]);
    var max = 100;
    var match = h.match(/\((\d+)\)/);
    if (match) max = parseInt(match[1]);
    headers.push({ name: h, colIndex: c, max: max });
  }

  // อ่านคะแนนนักเรียนจากชีต
  var scoreMap = {};
  for (var r = 1; r < data.length; r++) {
    var rowNo = Number(data[r][0]);
    var scoresObj = {};
    headers.forEach(function(h) {
      scoresObj[h.name] = data[r][h.colIndex];
    });
    scoreMap[rowNo] = {
      row: r + 1,
      scores: scoresObj
    };
  }

  // รวมรายชื่อนักเรียนจาก Roster กับคะแนนที่มี
  var finalStudents = [];
  if (studentRoster && studentRoster.length > 0) {
    studentRoster.forEach(function(s, idx) {
      var existing = scoreMap[s.no];
      if (existing) {
        finalStudents.push({
          row: existing.row,
          no: s.no,
          name: s.name,
          scores: existing.scores
        });
      } else {
        // นักเรียนใหม่ที่ยังไม่มีแถวในชีตคะแนนของผู้ใช้นี้ ให้เติมต่อท้าย
        var newRow = [s.no, s.name, s.className];
        headers.forEach(function() { newRow.push(""); });
        sheet.appendRow(newRow);
        
        finalStudents.push({
          row: sheet.getLastRow(),
          no: s.no,
          name: s.name,
          scores: {}
        });
      }
    });
  } else {
    // หากชีต Students ว่าง ให้ดึงรายชื่อจากชีตคะแนนโดยตรง
    for (var r = 1; r < data.length; r++) {
      var rowNo = Number(data[r][0]);
      var studentName = String(data[r][1] || "").trim();
      if (!studentName) continue;
      var scoresObj = {};
      headers.forEach(function(h) {
        scoresObj[h.name] = data[r][h.colIndex];
      });
      finalStudents.push({
        row: r + 1,
        no: rowNo || r,
        name: studentName,
        scores: scoresObj
      });
    }
  }

  finalStudents.sort(function(a, b) {
    return a.no - b.no;
  });

  return { headers: headers, students: finalStudents };
}

// บันทึกคะแนนลงในชีตคะแนนของผู้ใช้นั้นโดยเฉพาะ
function saveScore(className, userId, row, colIndex, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = getUserScoreSheetName(className, userId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "SheetNotFound";

  sheet.getRange(row, colIndex + 1).setValue(value);
  return "Success";
}

// เพิ่มช่องคะแนนใหม่ในชีตคะแนนของผู้ใช้นั้นโดยเฉพาะ
function addHeader(className, userId, headerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = getUserScoreSheetName(className, userId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "SheetNotFound";

  var nextCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, nextCol).setValue(headerName);
  return "Success";
}

// สร้างชั้นเรียนใหม่
function createNewClass(className) {
  var sheet = getOrCreateStudentsSheet();
  var data = sheet.getDataRange().getValues();
  var cls = String(className).trim();
  if (!cls) return "InvalidClassName";

  var exists = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cls) {
      exists = true;
      break;
    }
  }

  if (exists) return "Duplicate";

  // เพิ่มนักเรียนตัวอย่างสำหรับห้องเรียนใหม่ในชีต Students
  sheet.appendRow([1, "ด.ช. ตัวอย่าง รักเรียน", cls]);
  sheet.appendRow([2, "ด.ญ. มานี มีตา", cls]);
  SpreadsheetApp.flush();

  return "Success";
}

// ลบชั้นเรียน (เฉพาะ Admin)
function deleteClass(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cls = String(className).trim();
  if (!cls) return "InvalidClassName";

  // 1. ลบนักเรียนของชั้นนี้ออกจากชีต Students
  var studentSheet = ss.getSheetByName("Students");
  if (studentSheet) {
    var data = studentSheet.getDataRange().getValues();
    // ลบจากแถวล่างสุดขึ้นบน เพื่อป้องกัน row index คลาดเคลื่อน
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][2]).trim() === cls) {
        studentSheet.deleteRow(r + 1);
      }
    }
  }

  // 2. ลบชีตคะแนนของห้องนี้ทั้งหมด (เช่น cls + "_ADM01", cls + "_TCH01")
  var sheets = ss.getSheets();
  sheets.forEach(function(sh) {
    var shName = sh.getName();
    if (shName === cls || shName.indexOf(cls + "_") === 0) {
      if (ss.getSheets().length > 1) {
        try {
          ss.deleteSheet(sh);
        } catch(e) {}
      }
    }
  });

  SpreadsheetApp.flush();
  return "Success";
}

// ============================================================
// 6. เมนูและฟังก์ชันสร้างชีตเริ่มต้นอัตโนมัติ
// ============================================================

function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Gradebook 🚀')
      .addItem('🚀 เริ่มต้นสร้างชีตทั้งหมด (Students, Users, ห้องเรียน)', 'setupInitialSheets')
      .addToUi();
  } catch(err) {}
}

function setupInitialSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ชีต Students (ชื่อ, ชั้น, เลขที่)
  var studentSheet = getOrCreateStudentsSheet();
  
  // 2. ชีต Users (รหัสผู้ใช้, ชื่อ, วิชา, PIN, Role)
  var userSheet = getOrCreateUsersSheet();
  
  // 3. ชีตคะแนน ป.5/1_ADM01 (ภาษาไทย ครูบอส)
  var s1 = ss.getSheetByName("ป.5/1_ADM01");
  if (!s1) {
    s1 = ss.insertSheet("ป.5/1_ADM01");
    s1.appendRow(["No", "Name", "Class", "คะแนนเก็บภาษาไทย (10)", "การอ่านคิดวิเคราะห์ (10)", "สอบกลางภาคไทย (20)"]);
    var h1 = s1.getRange(1, 1, 1, 6);
    h1.setFontWeight("bold").setBackground("#6c5ce7").setFontColor("#ffffff").setHorizontalAlignment("center");
    
    var st5_1 = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ป.5/1", 9, 8, 18],
      [2, "ด.ญ. มานี มีตา", "ป.5/1", 10, 10, 19],
      [3, "ด.ช. ชูใจ ใจดี", "ป.5/1", 8, 7, 15],
      [4, "ด.ญ. ปิติ สุขใจ", "ป.5/1", 7, 9, 16],
      [5, "ด.ช. วีระ มานะ", "ป.5/1", 10, 9, 17]
    ];
    st5_1.forEach(function(r) { s1.appendRow(r); });
  }

  // 4. ชีตคะแนน ป.5/1_TCH01 (คณิตศาสตร์ ครูสมหมาย)
  var s2 = ss.getSheetByName("ป.5/1_TCH01");
  if (!s2) {
    s2 = ss.insertSheet("ป.5/1_TCH01");
    s2.appendRow(["No", "Name", "Class", "การบ้านคณิต (10)", "สอบเก็บคะแนนคณิต (15)"]);
    var h2 = s2.getRange(1, 1, 1, 5);
    h2.setFontWeight("bold").setBackground("#a29bfe").setFontColor("#ffffff").setHorizontalAlignment("center");
    
    var st5_2 = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ป.5/1", 8, 14],
      [2, "ด.ญ. มานี มีตา", "ป.5/1", 10, 15],
      [3, "ด.ช. ชูใจ ใจดี", "ป.5/1", 7, 11],
      [4, "ด.ญ. ปิติ สุขใจ", "ป.5/1", 9, 12],
      [5, "ด.ช. วีระ มานะ", "ป.5/1", 8, 13]
    ];
    st5_2.forEach(function(r) { s2.appendRow(r); });
  }
  
  // 5. ลบแผ่นงานเริ่มต้น (Sheet1 หรือ แผ่นงาน1) ที่ว่างเปล่า
  var defaultSheet = ss.getSheetByName("Sheet1") || ss.getSheetByName("แผ่นงาน1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch(e) {}
  }
  
  return "Success";
}
