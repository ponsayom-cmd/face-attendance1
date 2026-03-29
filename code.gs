/**
 * ============================================================
 * ระบบเช็คชื่อรายวิชาด้วยใบหน้า (เวอร์ชันสมบูรณ์ - No GPS)
 * วิธีใช้: Deploy > New deployment > Web App
 * Execute as: Me | Who has access: Anyone
 * ============================================================
 */

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getConfig') { 
      result = getConfig(); 
    } else if (action === 'getKnownFaces') { 
      result = getKnownFaces(); 
    } else if (action === 'getSubjects') { 
      result = getSubjects(); 
    } else if (action === 'getAttendanceReport') { 
      result = getAttendanceReport(e.parameter.date, e.parameter.subject); 
    } else { 
      result = { error: 'Unknown action: ' + action }; 
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid JSON' })).setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result;
  try {
    if (action === 'login') { 
      result = checkLogin(data.username, data.password); 
    } else if (action === 'registerUser') { 
      result = registerStudent(data.name, data.faceDescriptor); 
    } else if (action === 'logAttendance') { 
      result = logAttendance(data.name, data.subject); 
    } else if (action === 'saveConfig') { 
      result = saveConfig(data.subjects); 
    } else if (action === 'deleteStudent') { 
      result = deleteStudent(data.name); 
    } else { 
      result = { error: 'Unknown action' }; 
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- 1. ระบบจัดการรายวิชา ---
function getSubjects() {
  const config = getConfig();
  return config.subjects || [];
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Parameter', 'Value']);
    sheet.appendRow(['Subjects', '']);
  }
  const data = sheet.getDataRange().getValues();
  let config = { subjects: [] };
  data.forEach(row => {
    if (row[0] === 'Subjects') {
      config.subjects = String(row[1]).split(',').map(s => s.trim()).filter(s => s !== "");
    }
  });
  return config;
}

function saveConfig(subjects) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config') || ss.insertSheet('Config');
  sheet.clear().appendRow(['Parameter', 'Value']);
  sheet.appendRow(['Subjects', Array.isArray(subjects) ? subjects.join(',') : subjects]);
  return { success: true, message: 'บันทึกรายวิชาเรียบร้อยแล้ว' };
}

// --- 2. ระบบจัดการนักเรียนและตรวจสอบข้อมูลซ้ำ ---
function registerStudent(name, descriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'FaceDescriptor', 'RegDate']);
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === name.trim().toLowerCase()) {
      return { success: false, message: 'ชื่อนี้มีอยู่ในระบบแล้ว' };
    }
  }
  
  sheet.appendRow([name.trim(), JSON.stringify(descriptor), new Date()]);
  return { success: true, message: 'ลงทะเบียนสำเร็จ' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  if (sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({ 
    name: row[0], 
    descriptor: JSON.parse(row[1]),
    regDate: row[2] ? Utilities.formatDate(new Date(row[2]), "GMT+7", "dd/MM/yyyy HH:mm") : '-'
  }));
}

function deleteStudent(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === name.trim()) {
      sheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

// --- 3. ระบบเช็คชื่อ ---
function logAttendance(name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'Time', 'Date', 'Subject']);
  
  const now = new Date();
  const dateStr = Utilities.formatDate(now, "GMT+7", 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, "GMT+7", 'HH:mm:ss');
  
  sheet.appendRow([name, timeStr, dateStr, subject]);
  return { success: true };
}

function getAttendanceReport(targetDate, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  
  const allStudents = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => r[0]) : [];
  const attendanceData = attSheet ? attSheet.getDataRange().getValues().slice(1) : [];
  
  const filtered = attendanceData.filter(row => {
    const rowDate = row[2] instanceof Date ? Utilities.formatDate(row[2], "GMT+7", 'yyyy-MM-dd') : String(row[2]);
    return rowDate === targetDate && row[3] === subject;
  });
  
  return {
    attendance: filtered.map(row => ({ name: row[0], time: row[1] })),
    stats: { total: allStudents.length, present: filtered.length, absent: Math.max(0, allStudents.length - filtered.length) }
  };
}

// --- 4. ระบบ Login สำหรับอาจารย์ ---
function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Admins') || ss.insertSheet('Admins');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Username', 'Password']);
    sheet.appendRow(['admin', '1234']); 
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && String(data[i][1]) === String(password)) return { success: true };
  }
  return { success: false };
}
