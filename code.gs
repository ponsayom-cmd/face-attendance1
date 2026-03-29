/**
 * ระบบเช็คชื่อเวอร์ชันปรับปรุงความแม่นยำสูง
 */

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getKnownFaces') return response(getKnownFaces());
    if (action === 'getSubjects') return response(getSubjects());
    if (action === 'getAttendanceReport') return response(getAttendanceReport(e.parameter.date, e.parameter.subject));
    return response({ error: 'Invalid action' });
  } catch (err) { return response({ error: err.message }); }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  try {
    if (data.action === 'logAttendance') return response(logAttendance(data.name, data.subject));
    if (data.action === 'registerUser') return response(registerStudent(data.name, data.faceDescriptor));
    if (data.action === 'deleteStudent') return response(deleteStudent(data.name));
    return response({ error: 'Invalid action' });
  } catch (err) { return response({ error: err.message }); }
}

function response(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// --- การบันทึกเข้าเรียน (เน้นความเป๊ะของชื่อ) ---
function logAttendance(name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'Time', 'Date', 'Subject']);
  
  const now = new Date();
  const timeStr = Utilities.formatDate(now, "GMT+7", 'HH:mm:ss');
  const dateStr = Utilities.formatDate(now, "GMT+7", 'yyyy-MM-dd');
  
  // บันทึกชื่อแบบ Trim เสมอ
  sheet.appendRow([name.trim(), timeStr, dateStr, subject.trim()]);
  return { success: true };
}

// --- รายงานสถิติ (เปรียบเทียบชื่อแบบ Case-Insensitive) ---
function getAttendanceReport(targetDate, targetSubject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  
  const allStudents = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => r[0].toString().trim()) : [];
  if (!attSheet) return { attendance: [], stats: { total: allStudents.length, present: 0, absent: allStudents.length }, absentList: allStudents };
  
  const attRows = attSheet.getDataRange().getValues().slice(1);
  
  // กรองคนมาเรียน
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

// ฟังก์ชันอื่นๆ คงเดิมแต่เพิ่มการ .trim() ข้อมูล
function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    name: row[0].toString().trim(),
    descriptor: JSON.parse(row[1])
  }));
}

function getSubjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return [];
  const row = sheet.getDataRange().getValues().find(r => r[0] === 'Subjects');
  return row ? row[1].split(',').map(s => s.trim()).filter(s => s !== "") : [];
}

function registerStudent(name, descriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  sheet.appendRow([name.trim(), JSON.stringify(descriptor), new Date()]);
  return { success: true };
}
