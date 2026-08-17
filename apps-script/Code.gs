var SPREADSHEET_ID = '1a_etzeh5hdUt0fh-CUjG9mPBTK_kDMkXhRaykrgDQ_Q';
var DRIVE_FOLDER_ID = '13E4rOei0_qvPRh0w3LE19GxiIt87BKDA';
var ADMIN_PASSWORD = 'admin';

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'test') {
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({error:'Use POST'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;
    switch(action) {
      case 'login': result = handleLogin(data); break;
      case 'getDashboard': result = handleGetDashboard(data); break;
      case 'getBookings': result = handleGetBookings(data); break;
      case 'addBooking': result = handleAddBooking(data); break;
      case 'updateBookingStatus': result = handleUpdateBookingStatus(data); break;
      case 'deleteBooking': result = handleDeleteBooking(data); break;
      case 'uploadFile': result = handleUploadFile(data); break;
      case 'getFiles': result = handleGetFiles(data); break;
      case 'updateFileStatus': result = handleUpdateFileStatus(data); break;
      case 'addEvaluation': result = handleAddEvaluation(data); break;
      case 'getEvaluations': result = handleGetEvaluations(data); break;
      case 'getCalendarData': result = handleGetCalendarData(data); break;
      case 'getTeacherReport': result = handleGetTeacherReport(data); break;
      case 'getDepartmentReport': result = handleGetDepartmentReport(data); break;
      case 'getTeachers': result = handleGetTeachers(data); break;
      case 'initSheets': result = handleInitSheets(data); break;
      default: result = {success:false, error:'Unknown action: ' + action};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleLogin(data) {
  if (data.password === ADMIN_PASSWORD) {
    return {success:true, role:'admin'};
  }
  return {success:true, role:'teacher', name: data.name || ''};
}

function handleInitSheets(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ['Booking','Files','Supervision','Users'];
  var headers = {
    'Booking': ['Timestamp','Date','Time','TeacherName','Department','Period','SubjectName','SubjectCode','ClassLevel','Room','Status','Note'],
    'Files': ['Timestamp','TeacherName','FileType','FileName','FileURL','DriveFileID','Status','Note','BookingRef'],
    'Supervision': ['Timestamp','TeacherName','SupervisionDate','Strengths','Improvements','Suggestions','Summary','QualityLevel','EvaluatorName','BookingRef'],
    'Users': ['Name','Department','Role']
  };
  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
      sheet.setFrozenRows(1);
    }
  });
  return {success:true, message:'Sheets initialized'};
}

function handleGetDashboard(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {success:true};
  var bookingSheet = ss.getSheetByName('Booking');
  if (bookingSheet && bookingSheet.getLastRow() > 1) {
    var bookings = bookingSheet.getRange(2,1,bookingSheet.getLastRow()-1,12).getValues();
    result.totalBookings = bookings.length;
    result.pendingBookings = bookings.filter(function(r){return r[10]==='รอดำเนินการ';}).length;
    result.confirmedBookings = bookings.filter(function(r){return r[10]==='ยืนยันแล้ว';}).length;
    result.completedBookings = bookings.filter(function(r){return r[10]==='นิเทศแล้ว';}).length;
    result.rejectedBookings = bookings.filter(function(r){return r[10]==='ปฏิเสธ';}).length;
    result.recentBookings = bookings.slice(-10).reverse().map(function(r){
      return {timestamp:r[0],date:r[1],time:r[2],teacher:r[3],department:r[4],period:r[5],subject:r[6],code:r[7],level:r[8],room:r[9],status:r[10],note:r[11]};
    });
  } else {
    result.totalBookings = 0;
    result.pendingBookings = 0;
    result.confirmedBookings = 0;
    result.completedBookings = 0;
    result.rejectedBookings = 0;
    result.recentBookings = [];
  }
  var fileSheet = ss.getSheetByName('Files');
  if (fileSheet && fileSheet.getLastRow() > 1) {
    var files = fileSheet.getRange(2,1,fileSheet.getLastRow()-1,9).getValues();
    result.totalFiles = files.length;
    result.pendingFiles = files.filter(function(r){return r[6]==='รอตรวจสอบ';}).length;
    result.approvedFiles = files.filter(function(r){return r[6]==='ผ่าน';}).length;
    result.revisionFiles = files.filter(function(r){return r[6]==='ปรับปรุง';}).length;
    result.recentFiles = files.slice(-10).reverse().map(function(r){
      return {timestamp:r[0],teacher:r[1],fileType:r[2],fileName:r[3],fileURL:r[4],driveFileID:r[5],status:r[6],note:r[7],bookingRef:r[8]};
    });
  } else {
    result.totalFiles = 0;
    result.pendingFiles = 0;
    result.approvedFiles = 0;
    result.revisionFiles = 0;
    result.recentFiles = [];
  }
  var supSheet = ss.getSheetByName('Supervision');
  if (supSheet && supSheet.getLastRow() > 1) {
    var sups = supSheet.getRange(2,1,supSheet.getLastRow()-1,10).getValues();
    result.totalEvaluations = sups.length;
  } else {
    result.totalEvaluations = 0;
  }
  return result;
}

function handleGetBookings(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Booking');
  if (!sheet || sheet.getLastRow() <= 1) return {success:true, bookings:[]};
  var allData = sheet.getRange(2,1,sheet.getLastRow()-1,12).getValues();
  var bookings = allData.map(function(r,i){
    return {row:i+2, timestamp:r[0], date:r[1], time:r[2], teacher:r[3], department:r[4], period:r[5], subject:r[6], code:r[7], level:r[8], room:r[9], status:r[10], note:r[11]};
  });
  if (data.teacherName) {
    bookings = bookings.filter(function(b){return b.teacher === data.teacherName;});
  }
  if (data.status) {
    bookings = bookings.filter(function(b){return b.status === data.status;});
  }
  if (data.dateFrom && data.dateTo) {
    bookings = bookings.filter(function(b){return b.date >= data.dateFrom && b.date <= data.dateTo;});
  }
  return {success:true, bookings:bookings.reverse()};
}

function handleAddBooking(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Booking');
  if (!sheet) return {success:false, error:'Sheet not found'};
  var existing = [];
  if (sheet.getLastRow() > 1) {
    existing = sheet.getRange(2,1,sheet.getLastRow()-1,12).getValues();
  }
  var conflict = existing.find(function(r){
    return r[1] === data.date && r[2] === data.time && r[5] === data.period && r[10] !== 'ปฏิเสธ';
  });
  if (conflict) {
    return {success:false, error:'วัน-เวลา-คาบนี้มีการจองแล้ว (' + conflict[3] + ' - ' + conflict[6] + ')'};
  }
  var row = [
    new Date().toISOString(),
    data.date,
    data.time,
    data.teacherName,
    data.department,
    data.period,
    data.subjectName,
    data.subjectCode,
    data.classLevel,
    data.room,
    'รอดำเนินการ',
    data.note || ''
  ];
  sheet.appendRow(row);
  return {success:true, message:'จองวันนิเทศสำเร็จ'};
}

function handleUpdateBookingStatus(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Booking');
  if (!sheet) return {success:false, error:'Sheet not found'};
  sheet.getRange(data.row, 11).setValue(data.status);
  if (data.note !== undefined) {
    sheet.getRange(data.row, 12).setValue(data.note);
  }
  return {success:true, message:'อัพเดทสถานะสำเร็จ'};
}

function handleDeleteBooking(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Booking');
  if (!sheet) return {success:false, error:'Sheet not found'};
  sheet.deleteRow(data.row);
  return {success:true, message:'ลบการจองสำเร็จ'};
}

function handleUploadFile(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Files');
  if (!sheet) return {success:false, error:'Sheet not found'};
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var subFolders = {};
  var iter = folder.getFolders();
  while(iter.hasNext()) {
    var f = iter.next();
    subFolders[f.getName()] = f.getId();
  }
  var typeFolders = {'แผนการสอน':'Plans','สื่อการสอน':'Media','ภาพกิจกรรม':'Photos','คลิปวิดีโอ':'Clips'};
  var folderName = typeFolders[data.fileType] || 'Other';
  var targetFolderId = subFolders[folderName];
  if (!targetFolderId) {
    var newFolder = folder.createFolder(folderName);
    targetFolderId = newFolder.getId();
  }
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  var driveFileId = '';
  var fileURL = '';
  var fileName = data.fileName || 'file';
  if (data.fileData && data.fileType !== 'คลิปวิดีโอ') {
    var blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), data.mimeType || 'application/octet-stream', fileName);
    var file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    driveFileId = file.getId();
    fileURL = file.getUrl();
  } else {
    fileURL = data.fileURL || '';
    driveFileId = data.driveFileID || '';
  }
  var row = [
    new Date().toISOString(),
    data.teacherName,
    data.fileType,
    fileName,
    fileURL,
    driveFileId,
    'รอตรวจสอบ',
    data.note || '',
    data.bookingRef || ''
  ];
  sheet.appendRow(row);
  return {success:true, message:'อัพโหลดไฟล์สำเร็จ', fileURL:fileURL};
}

function handleGetFiles(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Files');
  if (!sheet || sheet.getLastRow() <= 1) return {success:true, files:[]};
  var allData = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  var files = allData.map(function(r,i){
    return {row:i+2, timestamp:r[0], teacher:r[1], fileType:r[2], fileName:r[3], fileURL:r[4], driveFileID:r[5], status:r[6], note:r[7], bookingRef:r[8]};
  });
  if (data.teacherName) {
    files = files.filter(function(f){return f.teacher === data.teacherName;});
  }
  if (data.fileType) {
    files = files.filter(function(f){return f.fileType === data.fileType;});
  }
  if (data.status) {
    files = files.filter(function(f){return f.status === data.status;});
  }
  return {success:true, files:files.reverse()};
}

function handleUpdateFileStatus(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Files');
  if (!sheet) return {success:false, error:'Sheet not found'};
  sheet.getRange(data.row, 7).setValue(data.status);
  if (data.note !== undefined) {
    sheet.getRange(data.row, 8).setValue(data.note);
  }
  return {success:true, message:'อัพเดทสถานะไฟล์สำเร็จ'};
}

function handleAddEvaluation(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Supervision');
  if (!sheet) return {success:false, error:'Sheet not found'};
  var row = [
    new Date().toISOString(),
    data.teacherName,
    data.supervisionDate,
    data.strengths,
    data.improvements,
    data.suggestions,
    data.summary,
    data.qualityLevel,
    data.evaluatorName,
    data.bookingRef || ''
  ];
  sheet.appendRow(row);
  return {success:true, message:'บันทึกผลการประเมินสำเร็จ'};
}

function handleGetEvaluations(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Supervision');
  if (!sheet || sheet.getLastRow() <= 1) return {success:true, evaluations:[]};
  var allData = sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
  var evals = allData.map(function(r,i){
    return {row:i+2, timestamp:r[0], teacher:r[1], date:r[2], strengths:r[3], improvements:r[4], suggestions:r[5], summary:r[6], qualityLevel:r[7], evaluator:r[8], bookingRef:r[9]};
  });
  if (data.teacherName) {
    evals = evals.filter(function(e){return e.teacher === data.teacherName;});
  }
  return {success:true, evaluations:evals.reverse()};
}

function handleGetCalendarData(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Booking');
  if (!sheet || sheet.getLastRow() <= 1) return {success:true, events:[]};
  var allData = sheet.getRange(2,1,sheet.getLastRow()-1,12).getValues();
  var events = allData.map(function(r){
    return {date:r[1], time:r[2], teacher:r[3], department:r[4], period:r[5], subject:r[6], status:r[10]};
  });
  if (data.month && data.year) {
    var m = parseInt(data.month);
    var y = parseInt(data.year);
    events = events.filter(function(e){
      var d = new Date(e.date);
      return d.getMonth() === m && d.getFullYear() === y;
    });
  }
  return {success:true, events:events};
}

function handleGetTeacherReport(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var teacherName = data.teacherName;
  var report = {success:true, teacher:teacherName, bookings:[], files:[], evaluations:[]};
  var bSheet = ss.getSheetByName('Booking');
  if (bSheet && bSheet.getLastRow() > 1) {
    var bData = bSheet.getRange(2,1,bSheet.getLastRow()-1,12).getValues();
    report.bookings = bData.filter(function(r){return r[3]===teacherName;}).map(function(r){
      return {date:r[1],time:r[2],department:r[4],period:r[5],subject:r[6],status:r[10]};
    });
    report.totalBookings = report.bookings.length;
    report.completedBookings = report.bookings.filter(function(b){return b.status==='นิเทศแล้ว';}).length;
  }
  var fSheet = ss.getSheetByName('Files');
  if (fSheet && fSheet.getLastRow() > 1) {
    var fData = fSheet.getRange(2,1,fSheet.getLastRow()-1,9).getValues();
    report.files = fData.filter(function(r){return r[1]===teacherName;}).map(function(r){
      return {fileType:r[2],fileName:r[3],status:r[6]};
    });
    report.totalFiles = report.files.length;
    report.approvedFiles = report.files.filter(function(f){return f.status==='ผ่าน';}).length;
  }
  var sSheet = ss.getSheetByName('Supervision');
  if (sSheet && sSheet.getLastRow() > 1) {
    var sData = sSheet.getRange(2,1,sSheet.getLastRow()-1,10).getValues();
    report.evaluations = sData.filter(function(r){return r[1]===teacherName;}).map(function(r){
      return {date:r[2],strengths:r[3],improvements:r[4],suggestions:r[5],summary:r[6],qualityLevel:r[7],evaluator:r[8]};
    });
    report.totalEvaluations = report.evaluations.length;
  }
  return report;
}

function handleGetDepartmentReport(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var department = data.department;
  var report = {success:true, department:department, teachers:[]};
  var bSheet = ss.getSheetByName('Booking');
  if (!bSheet || bSheet.getLastRow() <= 1) return {success:true, department:department, teachers:[]};
  var bData = bSheet.getRange(2,1,bSheet.getLastRow()-1,12).getValues();
  var teacherMap = {};
  bData.forEach(function(r){
    if (r[4] === department) {
      var name = r[3];
      if (!teacherMap[name]) teacherMap[name] = {name:name, bookings:0, completed:0, files:0, evaluated:0};
      teacherMap[name].bookings++;
      if (r[10] === 'นิเทศแล้ว') teacherMap[name].completed++;
    }
  });
  var fSheet = ss.getSheetByName('Files');
  if (fSheet && fSheet.getLastRow() > 1) {
    var fData = fSheet.getRange(2,1,fSheet.getLastRow()-1,9).getValues();
    fData.forEach(function(r){
      if (teacherMap[r[1]]) teacherMap[r[1]].files++;
    });
  }
  var sSheet = ss.getSheetByName('Supervision');
  if (sSheet && sSheet.getLastRow() > 1) {
    var sData = sSheet.getRange(2,1,sSheet.getLastRow()-1,10).getValues();
    sData.forEach(function(r){
      if (teacherMap[r[1]]) teacherMap[r[1]].evaluated++;
    });
  }
  report.teachers = Object.values(teacherMap);
  return report;
}

function handleGetTeachers(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) {
    return {success:true, teachers:[]};
  }
  var allData = sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
  var teachers = allData.filter(function(r){return r[2]==='teacher';}).map(function(r){
    return {name:r[0], department:r[1]};
  });
  return {success:true, teachers:teachers};
}
