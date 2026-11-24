/**
 * ============================================================
 * แอปพลิเคชัน: Gradebook (ระบบบันทึกคะแนนนักเรียน)
 * พัฒนาสำหรับ: ครูบอส
 * สภาพแวดล้อม: Google Apps Script + Google Sheets Web App
 * รองรับการใช้งาน: ทั้งใน Apps Script Web App และ GitHub Pages (Dual Mode)
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

// --- 1. ดึงรายชื่อชั้นเรียนทั้งหมด (Sheet ทั้งหมด) ---
function getAllClassNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // ดึงชื่อ Sheet ทุกอันออกมา (ไม่รวม sheet Config หากมี)
  var names = sheets.map(function(sheet) {
    return sheet.getName();
  }).filter(function(name) {
    return name !== "Config";
  });
  return names;
}

// --- 2. สร้างชั้นเรียนใหม่ ---
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

// --- 3. ดึงข้อมูลนักเรียนและคะแนน ---
function getClassData(className) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // อ่าน Header (เริ่ม Col C = index 2)
  var headers = [];
  for (var i = 2; i < data[0].length; i++) {
    var h = data[0][i];
    var max = 100;
    var match = h.match(/\((\d+)\)/);
    if (match) max = parseInt(match[1]);
    headers.push({ name: h, colIndex: i, max: max });
  }

  // อ่าน Students
  var students = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // ข้ามแถวที่ไม่มีชื่อหรือเลขที่
    if(row[0] === "" || row[1] === "") continue;

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

// --- 4. บันทึกคะแนน ---
function saveScore(className, row, colIndex, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return "SheetNotFound";
  sheet.getRange(row, colIndex + 1).setValue(value);
  return "Success";
}

// --- 5. เพิ่มช่องคะแนนใหม่ ---
function addHeader(className, headerName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(className);
  if (!sheet) return "SheetNotFound";
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue(headerName);
  return "Success";
}
