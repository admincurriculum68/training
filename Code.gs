// =========================================================================
// Code.gs [ฉบับอัปเกรด TMS - Single Input, Multi-Trainer & Attendance]
// =========================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('ระบบบริหารจัดการการอบรม (TMS)');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// -------------------------------------------------------------------------
// Helper Functions
// -------------------------------------------------------------------------
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetData(sheetName) {
  var sheet = getSS().getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  if (data.length <= 1) return result;
  
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function findRowIndex(sheet, taskId, traineeId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(taskId).trim() && String(data[i][2]).trim() === String(traineeId).trim()) {
      return i + 1;
    }
  }
  return -1;
}

// -------------------------------------------------------------------------
// 1. Authentication (Single Input)
// -------------------------------------------------------------------------
function loginUser(personalId) {
  try {
    var users = getSheetData('Users'); 
    
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      
      if (String(u.personal_id).trim() === String(personalId).trim() && String(u.status).trim() === 'ACTIVE') {
        return {
          status: 'SUCCESS',
          personal_id: u.personal_id,
          role: u.role,
          full_name: u.full_name,
          position: u.Position || '', 
          office: u.Office || '',     
          user_type: u.user_type || '', 
          trainer_id: u.trainer_id || '' 
        };
      }
    }
    return { status: 'FAILED', message: 'รหัสประจำตัวไม่ถูกต้อง หรือบัญชีถูกระงับการใช้งาน' };
  } catch (e) {
    return { status: 'FAILED', message: 'Database Error: ' + e.toString() };
  }
}

// -------------------------------------------------------------------------
// 2. Trainee Logic (ฝั่งผู้อบรม)
// -------------------------------------------------------------------------
function getTasksForTrainee(traineeId, userType) {
  try {
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');
    var evaluations = getSheetData('Evaluations');
    
    var openTasks = tasks.filter(function(t) { 
      var isTarget = (String(t.target_group).trim() === String(userType).trim() || String(t.target_group).trim() === 'ALL');
      return t.status === 'OPEN' && isTarget; 
    });

    var taskList = openTasks.map(function(task) {
      var mySub = submissions.find(function(s) {
        return String(s.task_id).trim() === String(task.task_id).trim() && String(s.trainee_id).trim() === String(traineeId).trim();
      });

      var myEval = null;
      if (mySub) {
        var subEvals = evaluations.filter(function(e) { return String(e.submission_id) === String(mySub.submission_id); });
        if (subEvals.length > 0) {
          myEval = subEvals[subEvals.length - 1];
        }
      }

      return {
        task_id: task.task_id,
        task_name: task.task_name,
        task_desc: task.description,
        task_type: task.task_type,
        deadline: task.close_date ? new Date(task.close_date).toLocaleDateString("th-TH") : "-",
        is_submitted: (mySub ? true : false),
        file_url: (mySub ? mySub.file_url : ''),
        eval_result: (myEval ? myEval.evaluation_result : null),
        eval_level: (myEval ? myEval.evaluation_level : null),
        eval_feedback: (myEval ? myEval.feedback : null)
      };
    });
    return taskList;
  } catch (e) {
    return [];
  }
}

// -------------------------------------------------------------------------
// 3. Trainer Grading Logic (ฝั่งวิทยากร)
// -------------------------------------------------------------------------
function getTrainerTrainees(trainerId) {
  var users = getSheetData('Users');
  var subms = getSheetData('Submissions');
  
  var myTrainees = users.filter(function(u) { 
    if (u.role !== 'TRAINEE' || !u.trainer_id) return false;
    var trainerArray = String(u.trainer_id).split(',').map(function(id) { return id.trim(); });
    return trainerArray.indexOf(String(trainerId).trim()) !== -1; 
  });

  return myTrainees.map(function(t) {
    var count = subms.filter(function(sub) { return String(sub.trainee_id).trim() === String(t.personal_id).trim(); }).length;
    return {
      trainee_id: t.personal_id,
      trainee_name: t.full_name,
      user_type: t.user_type,
      submitted_count: count
    };
  });
}

function getTraineeWorksForGrading(traineeId, userType) {
  var tasks = getSheetData('Tasks');
  var submissions = getSheetData('Submissions');
  var evaluations = getSheetData('Evaluations');
  
  var targetTasks = tasks.filter(function(t) { 
    var isTarget = (String(t.target_group).trim() === String(userType).trim() || String(t.target_group).trim() === 'ALL');
    return t.status === 'OPEN' && isTarget; 
  });

  return targetTasks.map(function(task) {
    var sub = submissions.find(function(s) {
      return String(s.task_id).trim() === String(task.task_id).trim() && String(s.trainee_id).trim() === String(traineeId).trim();
    });
    
    var eval = null;
    if (sub) {
      var evals = evaluations.filter(function(e) { return String(e.submission_id) === String(sub.submission_id); });
      if (evals.length > 0) eval = evals[evals.length - 1];
    }

    return {
      task_id: task.task_id,
      task_name: task.task_name,
      submission: sub ? {
        id: sub.submission_id,
        file_url: sub.file_url,
        type: "FILE", 
        date: new Date(sub.timestamp).toLocaleDateString("th-TH")
      } : null,
      evaluation: eval ? {
        result: eval.evaluation_result,
        feedback: eval.feedback
      } : null
    };
  });
}

function saveEvaluation(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSS().getSheetByName('Evaluations');
    var timestamp = new Date();
    sheet.appendRow([
      'EVAL-' + Utilities.getUuid().slice(0,8),
      data.submission_id,
      data.trainer_id, 
      data.level,
      data.feedback,
      timestamp,
      data.result
    ]);
    return { status: 'SUCCESS', message: 'บันทึกผลการประเมินแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// -------------------------------------------------------------------------
// 4. Admin Logic (ฝั่งผู้ดูแลระบบ)
// -------------------------------------------------------------------------
function getAllTasksAdmin() {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var result = [];
    var colMap = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      result.push({
        task_id: row[colMap['task_id']] || '',
        task_name: row[colMap['task_name']] || '',
        task_type: row[colMap['task_type']] || 'FILE',
        target_group: row[colMap['target_group']] || 'ALL',
        status: row[colMap['status']] || 'CLOSED',
        folder_id: row[colMap['folder_id']] || '',
        description: row[colMap['description']] || '',
        fiscal_year: row[colMap['fiscal_year']] || '',
        round: row[colMap['round']] || '',
        open_date: row[colMap['open_date']] || '-',
        close_date: row[colMap['close_date']] || '-'
      });
    }
    return result.reverse(); 
  } catch(e) { return []; }
}

function getAdminStats() {
  try {
    var users = getSheetData('Users');
    var tasks = getSheetData('Tasks');
    var submissions = getSheetData('Submissions');

    var trainees = users.filter(function(u) { return u.role === 'TRAINEE'; });
    var totalTrainees = trainees.length;
    var activeTasks = tasks.filter(function(t) { return t.status === 'OPEN'; });
    var totalActiveTasks = activeTasks.length;

    var traineeProgress = trainees.map(function(t) {
      var myTasks = activeTasks.filter(function(task) { 
        return (String(task.target_group).trim() === String(t.user_type).trim() || String(task.target_group).trim() === 'ALL'); 
      });
      
      var myTotalTasks = myTasks.length;
      var submittedCount = 0;
      
      if (myTotalTasks > 0) {
        submittedCount = myTasks.filter(function(task) {
          return submissions.some(function(sub) {
            return String(sub.task_id).trim() === String(task.task_id).trim() && String(sub.trainee_id).trim() === String(t.personal_id).trim();
          });
        }).length;
      }

      var percent = (myTotalTasks > 0) ? (submittedCount / myTotalTasks) * 100 : 0;

      return {
        trainee_id: t.personal_id,
        trainee_name: t.full_name,
        user_type: t.user_type || '-',
        submitted: submittedCount,
        total: myTotalTasks,
        percent: Math.round(percent)
      };
    });
    
    traineeProgress.sort(function(a, b) { return a.percent - b.percent; });

    return {
      totalTrainees: totalTrainees,
      totalActiveTasks: totalActiveTasks,
      traineeProgress: traineeProgress
    };
  } catch (e) { return null; }
}

function createNewTask(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSS().getSheetByName('Tasks');
    var lastRow = sheet.getLastRow();
    var newId = 'T001';
    
    if (lastRow > 1) {
      var lastId = sheet.getRange(lastRow, 1).getValue();
      var num = parseInt(lastId.replace('T', '')) + 1;
      newId = 'T' + String(num).padStart(3, '0');
    }

    var folderId = "";
    if (form.taskType !== 'LINK') {
      var parentFolderId = "1HHQgpS3CXvZLbP0e1QS9ORG2al89GbSd"; // ⚠️ เปลี่ยนเป็น ID ของ Folder หลัก
      var parentFolder;
      try { parentFolder = DriveApp.getFolderById(parentFolderId); } 
      catch (e) { parentFolder = DriveApp.getRootFolder(); }

      var folderName = newId + "_" + form.taskName;
      var newFolder = parentFolder.createFolder(folderName);
      newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      folderId = newFolder.getId();
    }

    sheet.appendRow([
      newId, form.taskName, form.description, form.taskType, form.targetGroup || 'ALL', form.round, form.fiscalYear,
      "'" + form.openDate, "'" + form.closeDate, folderId, 'OPEN'
    ]);
    return { status: 'SUCCESS', message: 'สร้างงานใหม่เรียบร้อย! (ID: ' + newId + ')' };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function toggleTaskStatus(taskId, currentStatus) {
  var sheet = getSS().getSheetByName('Tasks');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(taskId).trim()) {
      var newStatus = (currentStatus === 'OPEN') ? 'CLOSED' : 'OPEN';
      sheet.getRange(i + 1, 11).setValue(newStatus); 
      return { status: 'SUCCESS', message: 'เปลี่ยนสถานะเป็น ' + newStatus };
    }
  }
  return { status: 'FAILED', message: 'ไม่พบงาน' };
}

function updateTask(data) {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.taskId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ระบุ' };

    sheet.getRange(rowIndex, 2).setValue(data.taskName);
    sheet.getRange(rowIndex, 3).setValue(data.description);
    sheet.getRange(rowIndex, 4).setValue(data.taskType);
    sheet.getRange(rowIndex, 5).setValue(data.targetGroup);
    sheet.getRange(rowIndex, 6).setValue(data.round);
    sheet.getRange(rowIndex, 7).setValue(data.fiscalYear);
    sheet.getRange(rowIndex, 8).setValue("'" + data.openDate);
    sheet.getRange(rowIndex, 9).setValue("'" + data.closeDate);

    return { status: 'SUCCESS', message: 'บันทึกการแก้ไขเรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}

function handleFileUpload(fileData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var contentType = fileData.mimeType || "application/pdf";
    var blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), contentType, fileData.fileName);
    
    var tasks = getSheetData('Tasks');
    var targetTask = tasks.find(function(t) { return String(t.task_id).trim() === String(fileData.taskId).trim(); });
    var rawFolderId = (targetTask && targetTask.folder_id) ? String(targetTask.folder_id).trim() : "";
    var folder;
    try {
      if (rawFolderId && rawFolderId.length > 10) folder = DriveApp.getFolderById(rawFolderId);
      else folder = DriveApp.getRootFolder();
    } catch (err) { folder = DriveApp.getRootFolder(); }

    var newFileName = fileData.traineeId + "_" + fileData.taskId + "_" + fileData.fileName;
    var file = folder.createFile(blob).setName(newFileName);
    
    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, fileData.taskId, fileData.traineeId);
    var submitInfo = JSON.stringify({ type: 'FILE', originalName: fileData.fileName, mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });
    
    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(file.getUrl());
      sheet.getRange(existingRow, 5).setValue(file.getId());
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), fileData.taskId, fileData.traineeId, file.getUrl(), file.getId(), submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'ส่งไฟล์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Upload Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function handleLinkSubmission(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSS().getSheetByName('Submissions');
    var timestamp = new Date();
    var existingRow = findRowIndex(sheet, data.taskId, data.traineeId);
    var submitInfo = JSON.stringify({ type: 'LINK', mode: (existingRow > 0 ? 'RESUBMIT' : 'FIRST') });
    
    if (existingRow > 0) {
      sheet.getRange(existingRow, 4).setValue(data.url);
      sheet.getRange(existingRow, 5).setValue('LINK_SUBMISSION');
      sheet.getRange(existingRow, 6).setValue(submitInfo);
      sheet.getRange(existingRow, 7).setValue(timestamp);
    } else {
      sheet.appendRow(['SUB-' + Utilities.getUuid().slice(0,8), data.taskId, data.traineeId, data.url, 'LINK_SUBMISSION', submitInfo, timestamp, 'ON_TIME']);
    }
    return { status: 'SUCCESS', message: 'บันทึกลิงก์เรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: 'Link Error: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask(taskId) {
  try {
    var sheet = getSS().getSheetByName('Tasks');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(taskId).trim()) {
        var folderId = data[i][9];
        if (folderId && folderId.toString().length > 5) {
          try { DriveApp.getFolderById(folderId).setTrashed(true); } 
          catch (err) { console.log("ไม่สามารถลบ Folder ได้: " + err); }
        }
        sheet.deleteRow(i + 1);
        return { status: 'SUCCESS', message: 'ลบภาระงานและโฟลเดอร์เรียบร้อยแล้ว' };
      }
    }
    return { status: 'FAILED', message: 'ไม่พบรหัสภาระงานที่ต้องการลบ' };
  } catch (e) {
    return { status: 'FAILED', message: 'Error: ' + e.toString() };
  }
}

// -------------------------------------------------------------------------
// 5. Attendance Logic (ระบบลงเวลา Check-In / Check-Out)
// -------------------------------------------------------------------------

function getAttendanceStatus(traineeId) {
  try {
    var sheet = getSS().getSheetByName('Attendance');
    if (!sheet) return { status: 'FAILED', message: 'ไม่พบแผ่นงาน Attendance' };
    
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // ค้นหาจากล่างขึ้นบน (ข้อมูลล่าสุดของวันนี้)
    for (var i = data.length - 1; i >= 1; i--) {
      var rowDate = data[i][2];
      if (!rowDate) continue;
      
      var rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (String(data[i][1]).trim() === String(traineeId).trim() && rowDateStr === todayStr) {
        return {
          status: 'SUCCESS',
          checkedIn: data[i][3] ? true : false,
          checkedOut: data[i][5] ? true : false,
          checkInTime: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "HH:mm") : null,
          checkOutTime: data[i][5] ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "HH:mm") : null
        };
      }
    }
    return { status: 'SUCCESS', checkedIn: false, checkedOut: false };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  }
}

function recordCheckIn(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSS().getSheetByName('Attendance');
    if (!sheet) return { status: 'FAILED', message: 'ไม่พบแผ่นงาน Attendance' };
    
    var timestamp = new Date();
    sheet.appendRow([
      'ATT-' + Utilities.getUuid().slice(0,8),
      data.traineeId,
      timestamp, // date
      timestamp, // check_in_time
      data.goal, // learning_goal
      '',        // check_out_time
      ''         // reflection
    ]);
    return { status: 'SUCCESS', message: 'บันทึกเวลาเข้าอบรมเรียบร้อยแล้ว' };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function recordCheckOut(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = getSS().getSheetByName('Attendance');
    if (!sheet) return { status: 'FAILED', message: 'ไม่พบแผ่นงาน Attendance' };
    
    var rows = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (var i = rows.length - 1; i >= 1; i--) {
      var rowDate = rows[i][2];
      if (!rowDate) continue;
      
      var rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (String(rows[i][1]).trim() === String(data.traineeId).trim() && rowDateStr === todayStr) {
        // อัปเดตข้อมูล Check-Out ในแถวของวันนี้
        sheet.getRange(i + 1, 6).setValue(new Date()); // check_out_time
        sheet.getRange(i + 1, 7).setValue(data.reflection); // reflection
        return { status: 'SUCCESS', message: 'บันทึกเวลาออกและผลการสะท้อนเรียบร้อยแล้ว' };
      }
    }
    return { status: 'FAILED', message: 'ไม่พบประวัติการ Check-In ของวันนี้' };
  } catch (e) {
    return { status: 'FAILED', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
