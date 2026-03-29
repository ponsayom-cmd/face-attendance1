/**
 * แก้ไขระบบรายงานสถิติ: ปรับปรุงการเปรียบเทียบชื่อและวันที่ให้แม่นยำ 100%
 */

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getConfig') { result = getConfig(); }
    else if (action === 'getKnownFaces') { result = getKnownFaces(); }
    else if (action === 'getSubjects') { result = getSubjects(); }
    else if (action === 'getAttendanceReport') { 
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

// --- บันทึกการเข้าเรียน (เน้น Clean Data) ---
function logAttendance(name, subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance') || ss.insertSheet('Attendance');
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'Time', 'Date', 'Subject']);
  
  const now = new Date();
  const timeStr = Utilities.formatDate(now, "GMT+7", 'HH:mm:ss');
  const dateStr = Utilities.formatDate(now, "GMT+7", 'yyyy-MM-dd');
  
  // บันทึกแบบตัดช่องว่างออกให้สะอาด
  sheet.appendRow([name.toString().trim(), timeStr, dateStr, subject.toString().trim()]);
  return { success: true };
}

// --- ฟังก์ชันดึงรายงาน (แก้ไขจุดที่ทำให้ขึ้นว่าขาดเรียน) ---
function getAttendanceReport(targetDate, targetSubject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  
  // 1. ดึงรายชื่อนักเรียนทั้งหมดจากฐานข้อมูล (ทำความสะอาดชื่อ)
  const allStudents = studentSheet ? studentSheet.getDataRange().getValues().slice(1).map(r => r[0].toString().trim()) : [];
  
  // 2. ดึงข้อมูลการเข้าเรียน
  if (!attSheet) return { attendance: [], stats: { total: allStudents.length, present: 0, absent: allStudents.length } };
  
  const attendanceRows = attSheet.getDataRange().getValues().slice(1);
  
  // 3. กรองเฉพาะวันที่และวิชาที่เลือก (ใช้ String Comparison ที่ปลอดภัย)
  const presentData = attendanceRows.filter(row => {
    let rowDate = row[2];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, "GMT+7", "yyyy-MM-dd");
    } else {
      rowDate = String(rowDate).trim();
    }
    
    const rowSubject = String(row[3]).trim();
    // เทียบวันที่ตรงกัน และวิชาตรงกัน
    return rowDate === targetDate && rowSubject === targetSubject.trim();
  });

  // 4. สร้างรายการชื่อคนที่มาเรียน (ตัดช่องว่าง)
  const presentNames = presentData.map(row => row[0].toString().trim());

  // 5. คำนวณรายชื่อคนที่ขาด (คนที่มีชื่อใน Students แต่ไม่มีชื่อใน presentNames ของวัน/วิชานั้น)
  const absentStudents = allStudents.filter(name => !presentNames.includes(name));

  return {
    attendance: presentData.map(row => ({ name: row[0], time: row[1] })),
    absentList: absentStudents, // ส่งรายชื่อคนขาดกลับไปด้วย
    stats: { 
      total: allStudents.length, 
      present: presentNames.length, 
      absent: absentStudents.length 
    }
  };
}

// --- ฟังก์ชันเสริมอื่นๆ ---
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
  if (sheet.getLastRow() === 0) sheet.appendRow(['Name', 'FaceDescriptor', 'RegDate']);
  const data = sheet.getDataRange().getValues();
  const newName = name.trim();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === newName.toLowerCase()) return { success: false, message: 'ชื่อนี้มีอยู่แล้ว' };
  }
  sheet.appendRow([newName, JSON.stringify(descriptor), new Date()]);
  return { success: true };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students') || ss.insertSheet('Students');
  if (sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({ 
    name: row[0].toString().trim(), 
    descriptor: JSON.parse(row[1])
  }));
}
