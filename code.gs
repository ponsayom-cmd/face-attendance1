/**
 * ============================================================
 * ระบบเช็คชื่อเวอร์ชันสมบูรณ์ (จัดการ Admin + นักเรียน + ความแม่นยำสูง)
 * ============================================================
 */

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getKnownFaces') return response(getKnownFaces());
    if (action === 'getSubjects') return response(getSubjects());
    if (action === 'getAttendanceReport') return response(getAttendanceReport(e.parameter.date, e.parameter.subject));
    if (action === 'getConfig') return response(getConfig());
    if (action === 'getAdmins') return response(getAdmins()); // เพิ่มการดึงรายชื่อ Admin
    return response({ error: 'Invalid GET action' });
  } catch (err) { return response({ error: err.message }); }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  try {
    if (data.action === 'login') return response(checkLogin(data.username, data.password));
    if (data.action === 'logAttendance') return response(logAttendance(data.name, data.subject));
    if (data.action === 'registerUser') return response(registerStudent(data.name, data.faceDescriptor));
    if (data.action === 'deleteStudent') return response(deleteStudent(data.name));
    if (data.action === 'saveConfig') return response(saveConfig(data.subjects));
    if (data.action === 'addAdmin') return response(addAdmin(data.username, data.password)); // เพิ่ม Admin
    if (data.action === 'deleteAdmin') return response(deleteAdmin(data.username)); // ลบ Admin
    return response({ error: 'Invalid POST action' });
  } catch (err) { return response({ error: err.message }); }
}

function response(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// --- 1. ระบบจัดการ Admin ---
function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Admins') || ss.insertSheet('Admins');
  if (sheet.getLastRow() <= 1) {
    sheet.clear().appendRow(['Username', 'Password']);
    sheet.appendRow(['admin', '1234']); 
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === username.trim() && 
        data[i][1].toString().trim() === password.toString().trim()) {
      return { success: true };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function getAdmins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Admins');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({ username: row[0] }));
}

function addAdmin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Admins') || ss.insertSheet('Admins');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === username.trim()) return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
  }
  sheet.appendRow([username.trim(), password.trim()]);
  return { success: true };
}

function deleteAdmin(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Admins');
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  if (data.length <= 2) return { success: false, message: 'ต้องมี Admin อย่างน้อย 1 คน' };
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === username.trim()) {
      sheet.deleteRow(i + 1);
    }
  }
  return { success: true };
}

// --- 2. ระบบจัดการนักเรียน (คงเดิม) ---
function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    name: row[0].toString().trim(),
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

// ... ส่วนอื่นๆ (logAttendance, getAttendanceReport, etc.) ให้ใช้โค้ดเดิมที่เคยให้ไว้
function logAttendance(name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'Time', 'Date', 'Subject']);
  const now = new Date();
  const timeStr = Utilities.formatDate(now, "GMT+7", 'HH:mm:ss');
  const dateStr = Utilities.formatDate(now, "GMT+7", 'yyyy-MM-dd');
  sheet.appendRow([name.trim(), timeStr, dateStr, subject.trim()]);
  return { success: true };
}

function getAttendanceReport(targetDate, targetSubject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  const allStudents = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => r[0].toString().trim()) : [];
  if (!attSheet) return { attendance: [], stats: { total: allStudents.length, present: 0, absent: allStudents.length }, absentList: allStudents };
  const attRows = attSheet.getDataRange().getValues().slice(1);
  const presentData = attRows.filter(row => {
    let rowDate = row[2] instanceof Date ? Utilities.formatDate(row[2], "GMT+7", "yyyy-MM-dd") : String(row[2]).trim();
    return rowDate === targetDate && String(row[3]).trim() === targetSubject.trim();
  });
  const presentNames = presentData.map(row => row[0].toString().trim().toLowerCase());
  const absentList = allStudents.filter(name => !presentNames.includes(name.toLowerCase()));
  return {
    attendance: presentData.map(row => ({ name: row[0], time: row[1] })),
    absentList: absentList,
    stats: { total: allStudents.length, present: presentData.length, absent: absentList.length }
  };
}

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
  return { success: true };
}

function registerStudent(name, descriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === name.trim().toLowerCase()) {
      return { success: false, message: 'ชื่อนี้ถูกลงทะเบียนไปแล้ว' };
    }
  }
  sheet.appendRow([name.trim(), JSON.stringify(descriptor), new Date()]);
  return { success: true };
}
