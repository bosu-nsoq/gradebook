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
    
    // ข้อมูลเริ่มต้นตัวอย่างหลายชั้นเรียน (ใช้ชื่อระดับชั้นเป็นชื่อเต็ม)
    var sampleStudents = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ชั้นประถมศึกษาปีที่ 5/1"],
      [2, "ด.ญ. มานี มีตา", "ชั้นประถมศึกษาปีที่ 5/1"],
      [3, "ด.ช. ชูใจ ใจดี", "ชั้นประถมศึกษาปีที่ 5/1"],
      [4, "ด.ญ. ปิติ สุขใจ", "ชั้นประถมศึกษาปีที่ 5/1"],
      [5, "ด.ช. วีระ มานะ", "ชั้นประถมศึกษาปีที่ 5/1"],
      [1, "ด.ช. สมศักดิ์ สดใส", "ชั้นประถมศึกษาปีที่ 5/2"],
      [2, "ด.ญ. สมใจ มั่นคง", "ชั้นประถมศึกษาปีที่ 5/2"],
      [1, "ด.ช. เก่งกล้า ท้าทาย", "ชั้นประถมศึกษาปีที่ 6/1"],
      [2, "ด.ญ. สวยสม นิยมไทย", "ชั้นประถมศึกษาปีที่ 6/1"]
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

// ============================================================
// ฟังก์ชันจัดการชื่อระดับชั้นเรียนภาษาไทย (ชื่อเต็ม vs ชื่อย่อ)
// ============================================================
function toFullClassName(shortName) {
  if (!shortName) return "";
  var s = String(shortName).trim();
  
  // ประถมศึกษา
  var matchP = s.match(/^(?:ชั้น)?(?:ประถมศึกษาปีที่\s*|\.?[ปป]\.?\s*)(\d+)[\/|\.](\d+)$/);
  if (matchP) return "ชั้นประถมศึกษาปีที่ " + matchP[1] + "/" + matchP[2];
  
  // มัธยมศึกษา
  var matchM = s.match(/^(?:ชั้น)?(?:มัธยมศึกษาปีที่\s*|\.?[มม]\.?\s*)(\d+)[\/|\.](\d+)$/);
  if (matchM) return "ชั้นมัธยมศึกษาปีที่ " + matchM[1] + "/" + matchM[2];

  // อนุบาล
  var matchA = s.match(/^(?:ชั้น)?(?:อนุบาล(?:ปีที่)?\s*|\.?[ออ]\.?\s*)(\d+)[\/|\.](\d+)$/);
  if (matchA) return "ชั้นอนุบาลปีที่ " + matchA[1] + "/" + matchA[2];
  
  if (s.indexOf("ชั้น") === 0) return s;
  return "ชั้น " + s;
}

function toShortClassName(fullName) {
  if (!fullName) return "";
  var s = String(fullName).trim();
  var matchP = s.match(/^(?:ชั้น)?ประถมศึกษาปีที่\s*(\d+)[\/|\.](\d+)$/);
  if (matchP) return "ป." + matchP[1] + "/" + matchP[2];
  
  var matchM = s.match(/^(?:ชั้น)?มัธยมศึกษาปีที่\s*(\d+)[\/|\.](\d+)$/);
  if (matchM) return "ม." + matchM[1] + "/" + matchM[2];

  var matchA = s.match(/^(?:ชั้น)?อนุบาล(?:ปีที่)?\s*(\d+)[\/|\.](\d+)$/);
  if (matchA) return "อ." + matchA[1] + "/" + matchA[2];
  
  return s;
}

function isClassMatch(a, b) {
  if (!a || !b) return false;
  var strA = String(a).trim();
  var strB = String(b).trim();
  if (strA === strB) return true;
  return toFullClassName(strA) === toFullClassName(strB);
}

// ดึงรายชื่อชั้นเรียนทั้งหมด (ใช้ชื่อเต็ม เช่น ชั้นประถมศึกษาปีที่ 5/1)
function getAllClassNames() {
  var sheet = getOrCreateStudentsSheet();
  var data = sheet.getDataRange().getValues();
  var classesSet = {};

  for (var i = 1; i < data.length; i++) {
    var rawCls = String(data[i][2]).trim();
    if (rawCls) {
      var fullCls = toFullClassName(rawCls);
      classesSet[fullCls] = true;
      // ปรับปรุงในชีต Students ให้เป็นชื่อเต็มโดยอัตโนมัติ
      if (rawCls !== fullCls) {
        sheet.getRange(i + 1, 3).setValue(fullCls);
      }
    }
  }

  var classes = Object.keys(classesSet);
  classes.sort();
  if (classes.length === 0) {
    classes = ["ชั้นประถมศึกษาปีที่ 5/1", "ชั้นประถมศึกษาปีที่ 5/2", "ชั้นประถมศึกษาปีที่ 6/1"];
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

    if (isClassMatch(cls, className) && name) {
      students.push({
        no: Number(no),
        name: name,
        className: toFullClassName(cls)
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
  var fullCls = toFullClassName(className);
  sheet.appendRow([Number(no), String(name).trim(), fullCls]);
  return "Success";
}

// เพิ่มรายชื่อนักเรียนทีละหลายคนพร้อมกัน (Bulk Import แบบ Batch รวดเร็ว)
function addMultipleStudents(className, studentList) {
  var sheet = getOrCreateStudentsSheet();
  var fullCls = toFullClassName(className);
  
  if (Array.isArray(studentList) && studentList.length > 0) {
    var rows = [];
    studentList.forEach(function(s) {
      if (s.name && String(s.name).trim()) {
        rows.push([Number(s.no), String(s.name).trim(), fullCls]);
      }
    });
    if (rows.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, 3).setValues(rows);
    }
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

// ฟังก์ชันแปลง PIN เป็นสตริง 4 หลัก คงเลข 0 นำหน้าไว้เสมอ (เช่น 0001)
function normalizePin(val) {
  var s = String(val !== undefined && val !== null ? val : '').replace(/^'/, '').trim();
  if (/^\d+$/.test(s) && s.length < 4) {
    s = ("0000" + s).slice(-4);
  }
  return s;
}

function loginUserByPin(pin) {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var cleanPin = normalizePin(pin);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var userPin = normalizePin(row[3]);
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
      pin: normalizePin(row[3]),
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

  var pinStr = normalizePin(userData.pin);

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 2).setValue(userData.name);
    sheet.getRange(rowIndex, 3).setValue(userData.subject);
    sheet.getRange(rowIndex, 4).setNumberFormat("@").setValue("'" + pinStr);
    sheet.getRange(rowIndex, 5).setValue(userData.role || "user");
  } else {
    sheet.appendRow([
      userId,
      userData.name,
      userData.subject,
      "'" + pinStr,
      userData.role || "user",
      new Date()
    ]);
    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, 4).setNumberFormat("@").setValue("'" + pinStr);
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
//    ชื่อชีตคะแนน: [ClassName]_[UserID] เช่น ชั้นประถมศึกษาปีที่ 5/1_ADM01
// ============================================================

function getUserScoreSheetName(className, userId) {
  var uId = (userId && String(userId).trim() && String(userId).trim() !== "undefined") ? String(userId).trim() : "ADM01";
  var fullCls = toFullClassName(className);
  return fullCls ? (fullCls + "_" + uId) : "";
}

function getUserScoreSheet(ss, className, userId) {
  var uId = (userId && String(userId).trim() && String(userId).trim() !== "undefined") ? String(userId).trim() : "ADM01";
  var fullCls = toFullClassName(className);
  var shortCls = toShortClassName(className);

  var sheet = ss.getSheetByName(fullCls + "_" + uId);
  if (!sheet && shortCls) {
    sheet = ss.getSheetByName(shortCls + "_" + uId);
    if (sheet) {
      try {
        sheet.setName(fullCls + "_" + uId);
      } catch(e) {}
    }
  }
  return sheet;
}

// ดึงข้อมูลคะแนนเฉพาะของ User คนนั้นในห้องเรียนนั้น
function getClassData(className, userId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uId = (userId && String(userId).trim() && String(userId).trim() !== "undefined") ? String(userId).trim() : "ADM01";
  var fullCls = toFullClassName(className);
  if (!fullCls) return { headers: [], students: [] };

  var sheet = getUserScoreSheet(ss, fullCls, uId);

  // ดึงรายชื่อนักเรียนปัจจุบันในชั้นเรียนนี้จากชีต Students
  var studentRoster = getStudentsByClass(fullCls);

  // หากยังไม่มีชีตคะแนนของผู้ใช้นี้ ให้สร้างชีตใหม่พร้อมหัวข้อเริ่มต้น
  if (!sheet) {
    sheet = ss.insertSheet(fullCls + "_" + uId);
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

// ดึงข้อมูลสมุดคะแนนทั้งหมดทุกห้องเรียนของผู้ใช้ในครั้งเดียว
function getUserAllData(userId) {
  var classes = getAllClassNames();
  var classesData = {};
  
  classes.forEach(function(cls) {
    classesData[cls] = getClassData(cls, userId);
  });

  return {
    classes: classes,
    classesData: classesData
  };
}

// เข้าสู่ระบบพร้อมโหลดสมุดคะแนนทุกห้องของผู้ใช้ทันที (ระบบพร้อมทำงานเลย)
function loginAndLoadAll(pin) {
  var loginRes = loginUserByPin(pin);
  if (!loginRes || !loginRes.success) {
    return loginRes;
  }

  var user = loginRes.user;
  var allData = getUserAllData(user.userId);

  return {
    success: true,
    user: user,
    classes: allData.classes,
    classesData: allData.classesData
  };
}

// บันทึกคะแนนลงในชีตคะแนนของผู้ใช้นั้นโดยเฉพาะ
function saveScore(className, userId, row, colIndex, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uId = (userId && String(userId).trim() && String(userId).trim() !== "undefined") ? String(userId).trim() : "ADM01";
  var sheet = getUserScoreSheet(ss, className, uId);
  if (!sheet) return "SheetNotFound";

  sheet.getRange(row, colIndex + 1).setValue(value);
  SpreadsheetApp.flush();
  return "Success";
}

// เพิ่มช่องคะแนนใหม่ในชีตคะแนนของผู้ใช้นั้นโดยเฉพาะ
function addHeader(className, userId, headerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uId = (userId && String(userId).trim() && String(userId).trim() !== "undefined") ? String(userId).trim() : "ADM01";
  var sheet = getUserScoreSheet(ss, className, uId);
  if (!sheet) return "SheetNotFound";

  var nextCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, nextCol).setValue(headerName);
  SpreadsheetApp.flush();
  return "Success";
}

// สร้างชั้นเรียนใหม่
function createNewClass(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullCls = toFullClassName(className);
  if (!fullCls) return "InvalidClassName";

  var sheet = getOrCreateStudentsSheet();
  var data = sheet.getDataRange().getValues();
  var exists = false;

  for (var i = 1; i < data.length; i++) {
    if (isClassMatch(data[i][2], fullCls)) {
      exists = true;
      break;
    }
  }

  if (exists) return "Duplicate";

  // เพิ่มนักเรียนตัวอย่างสำหรับห้องเรียนใหม่ในชีต Students
  sheet.appendRow([1, "ด.ช. ตัวอย่าง รักเรียน", fullCls]);
  sheet.appendRow([2, "ด.ญ. มานี มีตา", fullCls]);
  SpreadsheetApp.flush();

  return "Success";
}

// ลบชั้นเรียน (เฉพาะ Admin)
// ลบชั้นเรียน (เฉพาะ Admin) - ปรับปรุงความเร็วแบบ Batch (เร็วขึ้น 95%+)
function deleteClass(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullCls = toFullClassName(className);
  var shortCls = toShortClassName(className);
  if (!fullCls) return { status: "Error", message: "InvalidClassName" };

  // 1. ลบนักเรียนของชั้นนี้ออกจากชีต Students แบบ Batch In-Memory (เสร็จในเสี้ยววินาที)
  var studentSheet = ss.getSheetByName("Students");
  if (studentSheet) {
    var data = studentSheet.getDataRange().getValues();
    if (data.length > 1) {
      var header = data[0];
      var remainingRows = [];
      for (var r = 1; r < data.length; r++) {
        if (!isClassMatch(data[r][2], fullCls)) {
          remainingRows.push(data[r]);
        }
      }
      studentSheet.clearContents();
      var newData = [header].concat(remainingRows);
      studentSheet.getRange(1, 1, newData.length, header.length).setValues(newData);
    }
  }

  // 2. ลบชีตคะแนนของห้องนี้ทั้งหมด
  var sheets = ss.getSheets();
  sheets.forEach(function(sh) {
    var shName = sh.getName();
    if (shName.indexOf(fullCls + "_") === 0 || (shortCls && shName.indexOf(shortCls + "_") === 0)) {
      if (ss.getSheets().length > 1) {
        try {
          ss.deleteSheet(sh);
        } catch(e) {}
      }
    }
  });

  SpreadsheetApp.flush();

  // 3. คืนค่ารายชื่อห้องเรียนที่เหลือทันที เพื่อลด round-trip จากหน้าเว็บ
  var remainingClasses = getAllClassNames();
  return {
    status: "Success",
    deletedClass: fullCls,
    remainingClasses: remainingClasses
  };
}

// ฟังก์ชันคืนค่าข้อมูลตัวอย่าง (สำหรับทดสอบระบบ หรือกรณีลบข้อมูลห้องเรียนหมด)
function resetSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName("Students");
  if (!studentSheet) studentSheet = ss.insertSheet("Students");
  studentSheet.clear();
  studentSheet.appendRow(["เลขที่", "ชื่อ", "ชั้น"]);
  var sampleStudents = [
    [1, "ด.ช. ตัวอย่าง รักเรียน", "ชั้นประถมศึกษาปีที่ 5/1"],
    [2, "ด.ญ. มานี มีตา", "ชั้นประถมศึกษาปีที่ 5/1"],
    [3, "ด.ช. ชูใจ ใจดี", "ชั้นประถมศึกษาปีที่ 5/1"],
    [4, "ด.ญ. ปิติ สุขใจ", "ชั้นประถมศึกษาปีที่ 5/1"],
    [5, "ด.ช. วีระ มานะ", "ชั้นประถมศึกษาปีที่ 5/1"],
    [1, "ด.ช. สมศักดิ์ สดใส", "ชั้นประถมศึกษาปีที่ 5/2"],
    [2, "ด.ญ. สมใจ มั่นคง", "ชั้นประถมศึกษาปีที่ 5/2"],
    [1, "ด.ช. เก่งกล้า ท้าทาย", "ชั้นประถมศึกษาปีที่ 6/1"],
    [2, "ด.ญ. สวยสม นิยมไทย", "ชั้นประถมศึกษาปีที่ 6/1"]
  ];
  studentSheet.getRange(2, 1, sampleStudents.length, 3).setValues(sampleStudents);
  var headerRange = studentSheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight("bold").setBackground("#6c5ce7").setFontColor("#ffffff").setHorizontalAlignment("center");
  SpreadsheetApp.flush();
  return { status: "Success", classes: getAllClassNames() };
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
  
  // 3. ชีตคะแนน ชั้นประถมศึกษาปีที่ 5/1_ADM01 (ภาษาไทย ครูบอส)
  var s1 = ss.getSheetByName("ชั้นประถมศึกษาปีที่ 5/1_ADM01") || ss.getSheetByName("ป.5/1_ADM01");
  if (!s1) {
    s1 = ss.insertSheet("ชั้นประถมศึกษาปีที่ 5/1_ADM01");
    s1.appendRow(["No", "Name", "Class", "คะแนนเก็บภาษาไทย (10)", "การอ่านคิดวิเคราะห์ (10)", "สอบกลางภาคไทย (20)"]);
    var h1 = s1.getRange(1, 1, 1, 6);
    h1.setFontWeight("bold").setBackground("#6c5ce7").setFontColor("#ffffff").setHorizontalAlignment("center");
    
    var st5_1 = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ชั้นประถมศึกษาปีที่ 5/1", 9, 8, 18],
      [2, "ด.ญ. มานี มีตา", "ชั้นประถมศึกษาปีที่ 5/1", 10, 10, 19],
      [3, "ด.ช. ชูใจ ใจดี", "ชั้นประถมศึกษาปีที่ 5/1", 8, 7, 15],
      [4, "ด.ญ. ปิติ สุขใจ", "ชั้นประถมศึกษาปีที่ 5/1", 7, 9, 16],
      [5, "ด.ช. วีระ มานะ", "ชั้นประถมศึกษาปีที่ 5/1", 10, 9, 17]
    ];
    st5_1.forEach(function(r) { s1.appendRow(r); });
  } else if (s1.getName() === "ป.5/1_ADM01") {
    try { s1.setName("ชั้นประถมศึกษาปีที่ 5/1_ADM01"); } catch(e) {}
  }

  // 4. ชีตคะแนน ชั้นประถมศึกษาปีที่ 5/1_TCH01 (คณิตศาสตร์ ครูสมหมาย)
  var s2 = ss.getSheetByName("ชั้นประถมศึกษาปีที่ 5/1_TCH01") || ss.getSheetByName("ป.5/1_TCH01");
  if (!s2) {
    s2 = ss.insertSheet("ชั้นประถมศึกษาปีที่ 5/1_TCH01");
    s2.appendRow(["No", "Name", "Class", "การบ้านคณิต (10)", "สอบเก็บคะแนนคณิต (15)"]);
    var h2 = s2.getRange(1, 1, 1, 5);
    h2.setFontWeight("bold").setBackground("#a29bfe").setFontColor("#ffffff").setHorizontalAlignment("center");
    
    var st5_2 = [
      [1, "ด.ช. ตัวอย่าง รักเรียน", "ชั้นประถมศึกษาปีที่ 5/1", 8, 14],
      [2, "ด.ญ. มานี มีตา", "ชั้นประถมศึกษาปีที่ 5/1", 10, 15],
      [3, "ด.ช. ชูใจ ใจดี", "ชั้นประถมศึกษาปีที่ 5/1", 7, 11],
      [4, "ด.ญ. ปิติ สุขใจ", "ชั้นประถมศึกษาปีที่ 5/1", 9, 12],
      [5, "ด.ช. วีระ มานะ", "ชั้นประถมศึกษาปีที่ 5/1", 8, 13]
    ];
    st5_2.forEach(function(r) { s2.appendRow(r); });
  } else if (s2.getName() === "ป.5/1_TCH01") {
    try { s2.setName("ชั้นประถมศึกษาปีที่ 5/1_TCH01"); } catch(e) {}
  }
  
  // 5. ลบแผ่นงานเริ่มต้นหรือแผ่นงานตกค้างที่ไม่ได้ใช้งาน (Sheet1, แผ่นงาน1, undefined_*)
  var junkSheets = ss.getSheets().filter(function(sh) {
    var name = sh.getName();
    return name.indexOf("undefined") === 0 || name === "Sheet1" || name === "แผ่นงาน1";
  });
  junkSheets.forEach(function(sh) {
    if (ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(sh);
      } catch(e) {}
    }
  });
  
  SpreadsheetApp.flush();
  return "Success";
}
