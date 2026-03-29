/**
 * แก้ไขระบบรายงานสถิติ: ปรับปรุงการเปรียบเทียบวันที่ให้แม่นยำ
 */

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getConfig') { result = getConfig(); }
    else if (action === 'getKnownFaces') { result = getKnownFaces(); }
    else if (action === 'getSubjects') { result = getSubjects(); }
    else if (action === 'getAttendanceReport') { 
      // แก้ไข: ส่งค่า date และ subject ไปตรวจสอบ
      result = getAttendanceReport(e.parameter.date, e.parameter.subject); 
    }
    else { result = { error: 'Unknown action' }; }
  } catch (err) { result = { error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data = JSON.parse(e.postData.contents);
  const action = data.action;
  let result;
  try {
    if (action === 'login') { result = checkLogin(data.username, data.password); }
    else if (action === 'registerUser') { result = registerStudent(data.name, data.faceDescriptor); }
    else if (action === 'logAttendance') { result = logAttendance(data.name, data.subject); }
    else if (action === 'saveConfig') { result = saveConfig(data.subjects); }
    else if (action === 'deleteStudent') { result = deleteStudent(data.name); }
  } catch (err) { result = { success: false, error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// --- แก้ไขการบันทึกเวลา ---
function logAttendance(name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'Time', 'Date', 'Subject']);
  
  const now = new Date();
  const timeStr = Utilities.formatDate(now, "GMT+7", 'HH:mm:ss');
  // บันทึกวันที่ในรูปแบบ String yyyy-MM-dd เพื่อป้องกัน Format เพี้ยน
  const dateStr = Utilities.formatDate(now, "GMT+7", 'yyyy-MM-dd');
  
  sheet.appendRow([name, timeStr, dateStr, subject]);
  return { success: true, message: 'บันทึกสำเร็จ' };
}

// --- แก้ไขการดึงรายงาน (จุดสำคัญที่ทำให้ขึ้นว่าขาดเรียน) ---
function getAttendanceReport(targetDate, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  
  // 1. รายชื่อนักเรียนทั้งหมด
  const allStudents = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => r[0].toString().trim()) : [];
  
  // 2. ข้อมูลการเข้าเรียน
  if (!attSheet) return { attendance: [], stats: { total: allStudents.length, present: 0, absent: allStudents.length } };
  
  const attendanceData = attSheet.getDataRange().getValues().slice(1);
  
  // กรองเฉพาะวันที่และวิชาที่เลือก
  const filtered = attendanceData.filter(row => {
    let rowDate = row[2];
    // แปลงวันที่จาก Sheet ให้เป็น yyyy-MM-dd เสมอไม่ว่าจะมาเป็น Date object หรือ String
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, "GMT+7", "yyyy-MM-dd");
    } else {
      rowDate = String(rowDate).trim();
    }
    
    const rowSubject = String(row[3]).trim();
    return rowDate === targetDate && rowSubject === subject;
  });

  const presentNames = filtered.map(row => row[0].toString().trim());
  
  return {
    attendance: filtered.map(row => ({ name: row[0], time: row[1] })),
    stats: { 
      total: allStudents.length, 
      present: filtered.length, 
      absent: Math.max(0, allStudents.length - filtered.length) 
    }
  };
}

// ฟังก์ชันอื่นๆ (getKnownFaces, deleteStudent, etc.) ให้ใช้ตามเวอร์ชันสมบูรณ์ก่อนหน้า
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

function getSubjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let subjects = [];
  data.forEach(row => {
    if (row[0] === 'Subjects') subjects = String(row[1]).split(',').map(s => s.trim()).filter(s => s !== "");
  });
  return subjects;
}

function registerStudent(name, descriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === name.trim().toLowerCase()) return { success: false, message: 'ชื่อซ้ำ' };
  }
  sheet.appendRow([name.trim(), JSON.stringify(descriptor), new Date()]);
  return { success: true };
}

function checkLogin(u, p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Admins') || ss.insertSheet('Admins');
  if (sheet.getLastRow() === 0) { sheet.appendRow(['Username', 'Password']); sheet.appendRow(['admin', '1234']); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === u && String(data[i][1]) === String(p)) return { success: true }; }
  return { success: false };
}
