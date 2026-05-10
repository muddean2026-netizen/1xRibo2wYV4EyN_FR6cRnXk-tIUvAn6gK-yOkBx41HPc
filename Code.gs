// ============ CONFIGURATION ============
const SHEET_ID = "1xRibo2wYV4EyN_FR6cRnXk-tIUvAn6gK-yOkBx41HPc";
const FOLDER_ID = "1bRh-AMIzc56dU_tKfD6NoefPUw9MlkUq";
const TELEGRAM_TOKEN = "8040742079:AAHz1I5a9Bi2QbiFXnaRMGGl5Yro8dqD2k8";
const TELEGRAM_CHAT_ID = "-4689717366";
const ADMIN_USERNAME = "admin";
const ADMIN_PASSWORD = "90030022";

// Sheet Names
const SHEET_NAMES = {
  REPORT: "รายงาน",
  CLASS: "ชั้นเรียน",
  STUDENTS: "ข้อมูลนักเรียน",
  REASONS: "สาเหตุ",
  RECORDERS: "ผู้บันทึก"
};

// ============ MAIN HANDLER ============
function doPost(e) {
  try {
    const action = e.parameter.action;
    
    switch(action) {
      case 'getDropdownData':
        return getDropdownData();
      case 'saveAttendance':
        return saveAttendance(e.parameter);
      case 'getReportData':
        return getReportData(e.parameter);
      case 'verifyAdmin':
        return verifyAdmin(e.parameter);
      case 'searchRecords':
        return searchRecords(e.parameter);
      case 'deleteRecord':
        return deleteRecord(e.parameter);
      case 'initializeSheets':
        return initializeSheets();
      case 'sendTelegramReport':
        return sendTelegramReport(e.parameter);
      default:
        return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Invalid action'}));
    }
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: 'Error: ' + error.toString()}));
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if(action === 'getPdfReport') {
      return getPdfReport(e.parameter);
    }
    
    return ContentService.createTextOutput('Method Not Allowed');
  } catch(error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// ============ INITIALIZE SHEETS ============
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // ลบชีตเก่า (ยกเว้นชีตที่มีข้อมูล)
  const sheetsToKeep = ['รายงาน', 'ชั้นเรียน', 'ข้อมูลนักเรียน', 'สาเหตุ', 'ผู้บันทึก'];
  ss.getSheets().forEach(sheet => {
    if (!sheetsToKeep.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
    }
  });

  // สร้างชีต "รายงาน"
  createReportSheet();
  
  // สร้างชีต "ชั้นเรียน"
  createClassSheet();
  
  // สร้างชีต "ข้อมูลนักเรียน"
  createStudentSheet();
  
  // สร้างชีต "สาเหตุ"
  createReasonSheet();
  
  // สร้างชีต "ผู้บันทึก"
  createRecorderSheet();

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    message: 'สร้างชีตแล้ว'
  })).setMimeType(ContentService.MimeType.JSON);
}

function createReportSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('รายงาน');
  
  if (!sheet) {
    sheet = ss.insertSheet('รายงาน');
  }
  
  // หัวตาราง
  const headers = ['วันที่', 'ชั้นเรียน', 'ชื่อ-สกุล', 'เวลา', 'สาเหตุ', 'หมายเหตุ', 'รูปภาพ', 'ผู้บันทึก', 'Link ไฟล์'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ฟอร์แมต
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#7c3aed')
    .setFontColor('white')
    .setFontWeight('bold');
  
  return sheet;
}

function createClassSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('ชั้นเรียน');
  
  if (!sheet) {
    sheet = ss.insertSheet('ชั้นเรียน');
  }
  
  // ข้อมูลชั้นเรียน
  const classes = [
    ['ม.1'],
    ['ม.2'],
    ['ม.3'],
    ['ม.4'],
    ['ม.5'],
    ['ม.6']
  ];
  
  sheet.getRange(1, 1, classes.length, 1).setValues(classes);
  
  return sheet;
}

function createStudentSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('ข้อมูลนักเรียน');
  
  if (!sheet) {
    sheet = ss.insertSheet('ข้อมูลนักเรียน');
  }
  
  // หัวตาราง
  sheet.getRange(1, 1).setValue('ชั้นเรียน');
  sheet.getRange(1, 2).setValue('ชื่อ-สกุล');
  
  // ข้อมูลตัวอย่าง
  const data = [
    ['ม.1', 'นายธิหัส สมบัติศรี'],
    ['ม.1', 'นางสาวชนาภา ศรีสวัสดิ์'],
    ['ม.2', 'นายวิทยา กมารกร'],
    ['ม.2', 'นางสาวกมลา หลวงพ่อนา'],
    ['ม.3', 'นายชัยพร วิริยศาสตร์'],
    ['ม.3', 'นางสาวลภัส เวทาลี']
  ];
  
  sheet.getRange(2, 1, data.length, 2).setValues(data);
  
  return sheet;
}

function createReasonSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('สาเหตุ');
  
  if (!sheet) {
    sheet = ss.insertSheet('สาเหตุ');
  }
  
  const reasons = [
    ['ป่วย'],
    ['ติดต่อไปสำนักงาน'],
    ['ไปแข่งขัน'],
    ['ไปรับรองการศึกษา'],
    ['หนาะที่อื่น'],
    ['อื่น ๆ']
  ];
  
  sheet.getRange(1, 1, reasons.length, 1).setValues(reasons);
  
  return sheet;
}

function createRecorderSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('ผู้บันทึก');
  
  if (!sheet) {
    sheet = ss.insertSheet('ผู้บันทึก');
  }
  
  const recorders = [
    ['ครูอัสมุดดีน เซ็งโซะ'],
    ['ครูประนอม นายหนา'],
    ['ครูสมใจ มหาชน'],
    ['นักเรียนอาสา']
  ];
  
  sheet.getRange(1, 1, recorders.length, 1).setValues(recorders);
  
  return sheet;
}

// ============ GET DROPDOWN DATA ============
function getDropdownData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // ชั้นเรียน
  const classSheet = ss.getSheetByName('ชั้นเรียน');
  const classData = classSheet.getRange(1, 1, classSheet.getLastRow(), 1).getValues();
  const classes = classData.map(row => row[0]).filter(v => v);
  
  // สาเหตุ
  const reasonSheet = ss.getSheetByName('สาเหตุ');
  const reasonData = reasonSheet.getRange(1, 1, reasonSheet.getLastRow(), 1).getValues();
  const reasons = reasonData.map(row => row[0]).filter(v => v);
  
  // ผู้บันทึก
  const recorderSheet = ss.getSheetByName('ผู้บันทึก');
  const recorderData = recorderSheet.getRange(1, 1, recorderSheet.getLastRow(), 1).getValues();
  const recorders = recorderData.map(row => row[0]).filter(v => v);
  
  // นักเรียนแยกตามชั้นเรียน
  const studentSheet = ss.getSheetByName('ข้อมูลนักเรียน');
  const studentData = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 2).getValues();
  
  const students = {};
  classes.forEach(cls => {
    students[cls] = studentData
      .filter(row => row[0] === cls)
      .map(row => row[1]);
  });
  
  return ContentService.createTextOutput(JSON.stringify({
    classes: classes,
    reasons: reasons,
    recorders: recorders,
    students: students
  })).setMimeType(ContentService.MimeType.JSON);
}

// ============ SAVE ATTENDANCE ============
function saveAttendance(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงาน');
    
    // เตรียมข้อมูล
    let imageLink = '';
    let fileName = '';
    
    // อัปโหลดรูปภาพ
    if (params.fileData) {
      const fileData = params.fileData;
      const mimeType = 'image/jpeg';
      
      // สร้างชื่อไฟล์
      fileName = `${params.date}_${params.class}_${params.studentName}`.replace(/[^a-zA-Z0-9_-]/g, '_') + '.jpg';
      
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
        const file = folder.createFile(blob);
        imageLink = file.getUrl();
      } catch(error) {
        Logger.log('Error uploading file: ' + error);
      }
    }
    
    // บันทึกข้อมูล
    const newRow = [
      params.date,
      params.class,
      params.studentName,
      params.time,
      params.reason,
      params.additionalNote || '',
      imageLink ? '✓' : '',
      params.recorder,
      imageLink || ''
    ];
    
    sheet.appendRow(newRow);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว ✅'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============ GET REPORT DATA ============
function getReportData(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงาน');
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    
    // กรองข้อมูล
    let filteredData = data.filter(row => {
      const date = row[0] ? new Date(row[0]).toISOString().split('T')[0] : '';
      const cls = row[1];
      
      let matchDate = true;
      let matchClass = true;
      
      if (params.date) {
        matchDate = date === params.date;
      }
      
      if (params.class) {
        matchClass = cls === params.class;
      }
      
      return matchDate && matchClass;
    });
    
    // สรุปตามชั้นเรียน
    const summary = {
      byClass: {}
    };
    
    filteredData.forEach(row => {
      const cls = row[1];
      summary.byClass[cls] = (summary.byClass[cls] || 0) + 1;
    });
    
    // สร้างข้อมูลสำหรับแสดง
    const reportData = filteredData.map((row, idx) => ({
      rowIndex: idx + 2,
      date: row[0] instanceof Date ? row[0].toLocaleDateString('th-TH') : row[0],
      class: row[1],
      name: row[2],
      time: row[3],
      reason: row[4],
      note: row[5],
      imageUrl: row[8],
      recorder: row[7]
    }));
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: reportData,
      summary: summary
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============ ADMIN FUNCTIONS ============
function verifyAdmin(params) {
  const isValid = params.username === ADMIN_USERNAME && params.password === ADMIN_PASSWORD;
  
  return ContentService.createTextOutput(JSON.stringify({
    isAdmin: isValid,
    message: isValid ? 'เข้าสู่ระบบสำเร็จ' : 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
  })).setMimeType(ContentService.MimeType.JSON);
}

function searchRecords(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงาน');
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const searchTerm = params.searchTerm.toLowerCase();
    
    // ค้นหา
    const results = data.filter(row => {
      return row[1].toLowerCase().includes(searchTerm) ||
             row[2].toLowerCase().includes(searchTerm) ||
             row[4].toLowerCase().includes(searchTerm);
    });
    
    const reportData = results.map((row, idx) => ({
      rowIndex: idx + 2,
      date: row[0] instanceof Date ? row[0].toLocaleDateString('th-TH') : row[0],
      class: row[1],
      name: row[2],
      time: row[3],
      reason: row[4]
    }));
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: reportData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function deleteRecord(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('รายงาน');
    
    const rowIndex = parseInt(params.rowIndex);
    sheet.deleteRow(rowIndex);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'ลบข้อมูลเรียบร้อยแล้ว'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============ TELEGRAM ============
function sendTelegramReport(params) {
  try {
    const text = params.reportText;
    const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
    
    const options = {
      method: 'post',
      payload: {
        chat_id: TELEGRAM_CHAT_ID,
        text: text,
        parse_mode: 'HTML'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'ส่งรายงานไปยัง Telegram แล้ว'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============ PDF REPORT ============
function getPdfReport(params) {
  try {
    const reportType = params.reportType || 'daily';
    const date = params.date || new Date().toISOString().split('T')[0];
    
    const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Kanit', Arial; }
    h1 { color: #7c3aed; text-align: center; }
    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
    th { background: #7c3aed; color: white; padding: 10px; }
    td { padding: 10px; border: 1px solid #ddd; }
    .summary { background: #f3e8ff; padding: 15px; margin: 20px 0; }
  </style>
</head>
<body>
  <h1>📋 รายงานการมาสายของนักเรียน</h1>
  <p style="text-align: center;">วันที่: ${date}</p>
  
  <div class="summary">
    <h2>📊 สรุปข้อมูล</h2>
    <p>รายงานผู้ติดต่อ (มาสาย) ตามประเภท: ${reportType}</p>
  </div>
  
  <table>
    <thead>
      <tr>
        <th>ลำดับ</th>
        <th>ชั้นเรียน</th>
        <th>ชื่อ-สกุล</th>
        <th>เวลา</th>
        <th>สาเหตุ</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td colspan="5" style="text-align: center; color: #999;">รายงานจะแสดงข้อมูลจากระบบ</td>
      </tr>
    </tbody>
  </table>
  
  <p style="margin-top: 30px; text-align: center; color: #999;">
    ระบบสายสิริ (Sai-Siri) | โรงเรียนบ้านสะพานหัก
  </p>
</body>
</html>
    `;
    
    return HtmlService.createHtmlOutput(html);
    
  } catch(error) {
    return HtmlService.createHtmlOutput('Error: ' + error);
  }
}

// ============ UTILITIES ============
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    date = new Date(date);
  }
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  return new Intl.DateTimeFormat('th-TH', options).format(date);
}