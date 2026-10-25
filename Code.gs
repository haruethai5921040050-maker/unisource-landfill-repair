var SS_NAME = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('แจ้งแผนงานซ่อมรายวัน - Unisource Land Fill')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันเช็คชื่อและรหัสผ่านจากชีส Users แบบยืดหยุ่น (รองรับการเว้นวรรคและชื่อในชีส)
function loginUser(role, password) {
  var sheet = SS_NAME.getSheetByName("Users");
  if (!sheet) {
    return { success: false, message: "ไม่พบชีสเก็บข้อมูลผู้ใช้งาน (Users) ใน Google Sheets" };
  }
  
  var data = sheet.getDataRange().getValues();
  
  // แปลงค่าที่ส่งมาจากหน้าเว็บให้เป็นตัวพิมพ์เล็กและตัดช่องว่าง เพื่อเทียบเคียงได้ง่าย
  var inputRole = String(role || "").trim().toLowerCase();
  var inputPass = String(password || "").trim();
  
  for (var i = 1; i < data.length; i++) {
    var dbUser = String(data[i][0] || "").trim();     // Col A: Username
    var dbPass = String(data[i][1] || "").trim();     // Col B: Password
    var dbFullName = String(data[i][2] || "").trim(); // Col C: FullName
    var dbRole = String(data[i][3] || "").trim();     // Col D: Role
    var dbStatus = String(data[i][4] || "").trim();   // Col E: Status
    
    // เทียบเคียง Username (แบบตรงกัน หรือแบบไม่สนตัวพิมพ์ใหญ่-เล็ก/เว้นวรรค)
    var isUserMatch = (dbUser === role || dbUser.toLowerCase() === inputRole || dbUser.replace(/\s+/g, '') === inputRole.replace(/\s+/g, ''));
    var isPassMatch = (dbPass === inputPass);
    
    if (isUserMatch && isPassMatch) {
      if (dbStatus.toLowerCase() !== "active" && dbStatus !== "") {
        return { success: false, message: "บัญชีนี้ถูกระงับการใช้งาน (Inactive)" };
      }
      
      writeLog("LOGIN", "SYSTEM", dbUser, "SUCCESS", "Login successful as " + dbRole);
      return { 
        success: true, 
        role: dbRole || "Admin",          // สิทธิ์การใช้งาน
        fullName: dbFullName || dbUser    // ชื่อเต็มที่จะแสดงบนมุมขวาบน
      };
    }
  }
  
  return { success: false, message: "รหัสผ่านไม่ถูกต้อง หรือชื่อผู้ใช้งานไม่ถูกต้อง" };
}

function getVehicles() {
  var sheet = SS_NAME.getSheetByName("Vehicles");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var vehicles = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      vehicles.push({
        vehicleNumber: String(data[i][0]),
        plateNumber: String(data[i][1] || ''),
        driverName: String(data[i][2] || ''),
        status: String(data[i][4] || 'พร้อมใช้งาน')
      });
    }
  }
  return vehicles;
}

function getRepairs() {
  var sheet = SS_NAME.getSheetByName("Repairs");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var repairs = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var rawStatus = String(data[i][4] || '');
      var statusDisplay = (rawStatus === 'RDY' || rawStatus === 'พร้อมใช้งาน') ? 'พร้อมใช้งาน' : rawStatus;
      
      var startDate = data[i][5] ? new Date(data[i][5]) : null;
      var startDateISO = startDate ? startDate.getFullYear() + '-' + String(startDate.getMonth() + 1).padStart(2, '0') + '-' + String(startDate.getDate()).padStart(2, '0') : '';
      var startDateText = startDate ? startDate.toLocaleDateString('th-TH') + ' ' + String(startDate.getHours()).padStart(2, '0') + ':' + String(startDate.getMinutes()).padStart(2, '0') : '-';

      var endDate = data[i][6] ? new Date(data[i][6]) : null;
      var endDateText = endDate ? endDate.toLocaleDateString('th-TH') + ' ' + String(endDate.getHours()).padStart(2, '0') + ':' + String(endDate.getMinutes()).padStart(2, '0') : '';

      repairs.push({
        repairID: String(data[i][0]),
        vehicleNumber: String(data[i][1]),
        plateNumber: String(data[i][2] || ''),
        driverName: String(data[i][3] || ''),
        repairStatus: statusDisplay,
        startTime: data[i][5] ? new Date(data[i][5]).toISOString() : '',
        endTime: data[i][6] ? new Date(data[i][6]).toISOString() : '',
        endDateText: endDateText,
        duration: data[i][7] || 0,
        repairCause: String(data[i][8] || '-'),
        garage: String(data[i][9] || '-'),
        approvalStatus: String(data[i][10] || 'โปรดเลือก'),
        postRepairStatus: String(data[i][11] || 'กำลังซ่อม'),
        remark: String(data[i][12] || '-'),
        transportAck: String(data[i][13] || 'FALSE'),
        recorder: String(data[i][16] || '-'),
        startDateText: startDateText,
        startDateISO: startDateISO
      });
    }
  }
  return repairs;
}

function generateRepairID() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SS_NAME.getSheetByName("Repairs");
    if (!sheet) return "UNS-00001";
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return "UNS-00001";
    
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxNum = 0;
    
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i][0]);
      if (idStr.indexOf("UNS-") === 0) {
        var num = parseInt(idStr.replace("UNS-", ""), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
    
    return "UNS-" + ("00000" + (maxNum + 1)).slice(-5);
  } finally {
    lock.releaseLock();
  }
}

function saveRepairRow(data, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var repairID = generateRepairID();
  var now = new Date();
  
  var start = data.startTime ? new Date(data.startTime) : now;
  var end = data.endTime ? new Date(data.endTime) : now;
  var duration = Math.max(0, Math.floor((end - start) / (1000 * 60)));
  
  var rowData = [
    repairID,
    data.vehicleNumber,
    data.plateNumber,
    data.driverName,
    data.repairStatus,
    start,
    "",
    duration,
    data.repairCause || "-",
    data.garage || "-",
    "โปรดเลือก",
    "กำลังซ่อม",
    data.remark || "-",
    "FALSE",
    "",
    "",
    data.recorder || userName
  ];
  
  sheet.appendRow(rowData);
  writeLog("SAVE_REPAIR", userName, userRole, repairID, "Saved repair row for " + data.vehicleNumber);
  return { success: true, message: "บันทึกเรียบร้อย" };
}

function saveMultipleRepairs(items, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var now = new Date();
  
  items.forEach(function(data) {
    var repairID = generateRepairID();
    var start = data.startTime ? new Date(data.startTime) : now;
    var end = data.endTime ? new Date(data.endTime) : now;
    var duration = Math.max(0, Math.floor((end - start) / (1000 * 60)));
    
    var rowData = [
      repairID,
      data.vehicleNumber,
      data.plateNumber,
      data.driverName,
      data.repairStatus,
      start,
      "",
      duration,
      data.repairCause || "-",
      data.garage || "-",
      "โปรดเลือก",
      "กำลังซ่อม",
      data.remark || "-",
      "FALSE",
      "",
      "",
      data.recorder || userName
    ];
    sheet.appendRow(rowData);
  });

  writeLog("SAVE_BULK_REPAIR", userName, userRole, "BULK", "Saved " + items.length + " vehicle records");
  return { success: true, message: "บันทึกเรียบร้อย" };
}

function updatePostRepair(repairID, postStatus, endTime, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(repairID)) {
      sheet.getRange(i + 1, 12).setValue(postStatus);
      
      if (postStatus === "เสร็จพร้อมรับงาน" && endTime) {
        var start = new Date(data[i][5]);
        var end = new Date(endTime);
        var duration = Math.max(0, Math.floor((end - start) / (1000 * 60)));
        
        sheet.getRange(i + 1, 7).setValue(end);
        sheet.getRange(i + 1, 8).setValue(duration);
      }
      
      writeLog("UPDATE_POST_REPAIR", userName, userRole, repairID, "Updated post status to " + postStatus);
      return { success: true, message: "บันทึกเรียบร้อย" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

function approveRepair(repairID, status, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(repairID)) {
      sheet.getRange(i + 1, 11).setValue(status);
      sheet.getRange(i + 1, 13).setValue(userName);
      writeLog("APPROVE_REPAIR", userName, userRole, repairID, "Set approval to " + status);
      return { success: true, message: "บันทึกเรียบร้อย" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

function acknowledgeRepair(repairID, isChecked, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(repairID)) {
      sheet.getRange(i + 1, 14).setValue(isChecked ? "TRUE" : "FALSE");
      sheet.getRange(i + 1, 15).setValue(userName);
      sheet.getRange(i + 1, 16).setValue(new Date());
      writeLog("ACK_REPAIR", userName, userRole, repairID, "Transport ack updated");
      return { success: true, message: "บันทึกเรียบร้อย" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

function deleteRepairRecord(repairID, userName, userRole) {
  var sheet = SS_NAME.getSheetByName("Repairs");
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(repairID)) {
      sheet.getRange(i + 1, 11).setValue("ยกเลิก/ลบแล้ว");
      writeLog("DELETE_REPAIR", userName, userRole, repairID, "Marked record as deleted/archived");
      return { success: true, message: "ลบเรียบร้อย" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

function writeLog(action, user, role, targetID, details) {
  try {
    var sheet = SS_NAME.getSheetByName("SystemLogs");
    if (!sheet) return;
    var now = new Date();
    sheet.appendRow(["LOG-" + now.getTime(), now, user, role, action, targetID, details, ""]);
  } catch (e) {
    console.error("Log error: " + e.toString());
  }
}