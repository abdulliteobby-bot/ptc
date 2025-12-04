let cachedEmployeeCodes = null;


function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('บันทึกวันลา')
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3012/3012434.png');
  } catch (e) {
    return HtmlService.createHtmlOutput(`
      <h3 style="color:red">Error ใน doGet():</h3>
      <pre>${e.toString()}</pre>
      <pre>${e.stack}</pre>
    `);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function getEmployeeCodes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employees");
    if (!cachedEmployeeCodes) {
      const data = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 1).getValues();
      cachedEmployeeCodes = data.flat().filter(code => code).map(code => String(code).trim());
    }
    return cachedEmployeeCodes;
  } catch (error) {
    console.error("Error in getEmployeeCodes:", error);
    return [];
  }
}


function getEmployeesWithUsed() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employees");
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 11).getValues();
    const leaveData = leaveSheet.getRange(2, 1, leaveSheet.getLastRow() - 1, 10).getValues();
    const allUsed = {};
    for (let row of leaveData) {
      const code = String(row[1]).trim();
      const type = String(row[5]).trim();
      const total = parseFloat(row[8]) || 0;
      if (!allUsed[code]) {
        allUsed[code] = {
          "สาย": 0,
          "ลาป่วย": 0,
          "ลากิจ": 0,
          "ลาพักผ่อน": 0,
          "ลาคลอด": 0,
          "ลาบวช": 0
        };
      }
      if (allUsed[code].hasOwnProperty(type)) {
        allUsed[code][type] += total;
      }
    }
    const employeesWithUsed = employeeData.map(emp => {
      const code = String(emp[0]).trim();
      const used = allUsed[code] || {
        "สาย": 0,
        "ลาป่วย": 0,
        "ลากิจ": 0,
        "ลาพักผ่อน": 0,
        "ลาคลอด": 0,
        "ลาบวช": 0
      };
      return {
        code: code,
        name: emp[1],
        section: emp[2],
        position: emp[3],
        late: parseFloat(emp[4]) || 0,
        sick: parseFloat(emp[5]) || 0,
        business: parseFloat(emp[6]) || 0,
        vacation: parseFloat(emp[7]) || 0,
        maternity: parseFloat(emp[8]) || 0,
        ordination: parseFloat(emp[9]) || 0,
        used: used
      };
    });
    return employeesWithUsed;
  } catch (error) {
    console.error("Error in getEmployeesWithUsed:", error);
    return [];
  }
}


function getAdminPassword() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employees");
    return String(sheet.getRange("L1").getValue()).trim();
  } catch (error) {
    console.error("Error in getAdminPassword:", error);
    return "";
  }
}


function verifyPassword(userType, code, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employees");
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    if (userType === "admin") {
      const adminPassword = getAdminPassword();
      if (String(password).trim() === adminPassword) {
        return { success: true, isAdmin: true };
      } else {
        return { success: false, message: "รหัสผ่าน Admin ไม่ถูกต้อง" };
      }
    } else {
      const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 11).getValues();
      let employee = null;
      code = String(code).trim();
      password = String(password).trim();
      for (let i = 0; i < employeeData.length; i++) {
        if (String(employeeData[i][0]).trim() === code && String(employeeData[i][10]).trim() === password) {
          employee = {
            code: employeeData[i][0],
            name: employeeData[i][1],
            section: employeeData[i][2],
            position: employeeData[i][3],
            late: parseFloat(employeeData[i][4]) || 0,
            sick: parseFloat(employeeData[i][5]) || 0,
            business: parseFloat(employeeData[i][6]) || 0,
            vacation: parseFloat(employeeData[i][7]) || 0,
            maternity: parseFloat(employeeData[i][8]) || 0,
            ordination: parseFloat(employeeData[i][9]) || 0
          };
          break;
        }
      }
      if (!employee) {
        return { success: false, message: "รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง" };
      }
      const used = {
        "สาย": 0,
        "ลาป่วย": 0,
        "ลากิจ": 0,
        "ลาพักผ่อน": 0,
        "ลาคลอด": 0,
        "ลาบวช": 0
      };
      const lastRow = leaveSheet.getLastRow();
      if (lastRow > 1) {
        const leaveData = leaveSheet.getRange(2, 1, lastRow - 1, 10).getValues();
        for (let row of leaveData) {
          if (String(row[1]).trim() === code) {
            const type = String(row[5]).trim();
            const total = parseFloat(row[8]) || 0;
            if (used.hasOwnProperty(type)) {
              used[type] += total;
            }
          }
        }
      }
      return {
        success: true,
        isAdmin: false,
        employee: { ...employee, used: used }
      };
    }
  } catch (error) {
    console.error("Error in verifyPassword:", error);
    return { success: false, message: "เกิดข้อผิดพลาดในการตรวจสอบรหัสผ่าน" };
  }
}


function submitLeaveRecord(record, code, password) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {  // Wait up to 10 seconds
      return { success: false, message: "ระบบกำลัง忙 ลองใหม่ในไม่กี่วินาที" };
    }
    let verification;
    if (code && password) {
      verification = verifyPassword("user", code, password);
    } else {
      verification = { success: true };
    }
    if (!verification.success) {
      lock.releaseLock();
      return verification;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    const employeeSheet = ss.getSheetByName("Employees");
    const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow()-1, 11).getValues();
    let employeeRow = -1;
    let employee = null;
    for (let i = 0; i < employeeData.length; i++) {
      if (String(employeeData[i][0]).trim() === record.code) {
        employeeRow = i + 2;
        employee = {
          code: employeeData[i][0],
          name: employeeData[i][1],
          section: employeeData[i][2],
          position: employeeData[i][3],
          late: parseFloat(employeeData[i][4]) || 0,
          sick: parseFloat(employeeData[i][5]) || 0,
          business: parseFloat(employeeData[i][6]) || 0,
          vacation: parseFloat(employeeData[i][7]) || 0,
          maternity: parseFloat(employeeData[i][8]) || 0,
          ordination: parseFloat(employeeData[i][9]) || 0
        };
        break;
      }
    }
    if (!employee) {
      lock.releaseLock();
      return { success: false, message: "ไม่พบข้อมูลพนักงาน" };
    }
    const leaveType = record.type;
    const daysRequested = parseFloat(record.total);
    let remainingDays = 0;
    switch(leaveType) {
      case "สาย":
        remainingDays = employee.late;
        break;
      case "ลาป่วย":
        remainingDays = employee.sick;
        break;
      case "ลากิจ":
        remainingDays = employee.business;
        break;
      case "ลาพักผ่อน":
        remainingDays = employee.vacation;
        break;
      case "ลาคลอด":
        remainingDays = employee.maternity;
        break;
      case "ลาบวช":
        remainingDays = employee.ordination;
        break;
      default:
        lock.releaseLock();
        return { success: false, message: "ประเภทการลาไม่ถูกต้อง" };
    }
    if (daysRequested > remainingDays) {
      lock.releaseLock();
      return {
        success: false,
        message: `วันลาไม่เพียงพอ (ขอ ${daysRequested} วัน แต่เหลือ ${remainingDays} วัน)`
      };
    }
    const timestamp = new Date();
    // Generate reference number
    const lastRow = leaveSheet.getLastRow();
    let maxNum = 0;
    const prefix = `LR-${Utilities.formatDate(timestamp, "GMT+7", "yyyyMM")}-`;
    if (lastRow > 1) {
      const existingRefs = leaveSheet.getRange(2, 11, lastRow - 1, 1).getValues().flat();
      existingRefs.forEach(ref => {
        if (ref && ref.startsWith(prefix)) {
          const num = parseInt(ref.substring(prefix.length), 10);
          if (!isNaN(num) && num > maxNum) maxNum = num;
        }
      });
    }
    const refNum = `${prefix}${(maxNum + 1).toString().padStart(4, '0')}`;


    // Calculate remaining balances for all leave types after deduction
    const remainingBalances = {
      "สาย": employee.late - (leaveType === "สาย" ? daysRequested : 0),
      "ลาป่วย": employee.sick - (leaveType === "ลาป่วย" ? daysRequested : 0),
      "ลากิจ": employee.business - (leaveType === "ลากิจ" ? daysRequested : 0),
      "ลาพักผ่อน": employee.vacation - (leaveType === "ลาพักผ่อน" ? daysRequested : 0),
      "ลาคลอด": employee.maternity - (leaveType === "ลาคลอด" ? daysRequested : 0),
      "ลาบวช": employee.ordination - (leaveType === "ลาบวช" ? daysRequested : 0)
    };


    // Append the leave record with remaining balances in columns L to Q, and empty for S-V
    leaveSheet.appendRow([
      timestamp,
      employee.code,
      employee.name,
      employee.section,
      employee.position,
      record.type,
      record.start,
      record.finish,
      daysRequested,
      record.remark || "ไม่ระบุ",
      refNum,
      remainingBalances["สาย"],      // Column L
      remainingBalances["ลาป่วย"],  // Column M
      remainingBalances["ลากิจ"],    // Column N
      remainingBalances["ลาพักผ่อน"], // Column O
      remainingBalances["ลาคลอด"],   // Column P
      remainingBalances["ลาบวช"],     // Column Q
      "",                             // Placeholder for R if needed
      "",                             // Column S: Approver 1
      "",                             // Column T: Remark 1
      "",                             // Column U: Approver 2
      ""                              // Column V: Remark 2
    ]);


    // Update the employee's remaining leave balance in the Employees sheet
    const colIndex = getLeaveColumnIndex(leaveType);
    const currentValue = parseFloat(employeeSheet.getRange(employeeRow, colIndex).getValue()) || 0;
    employeeSheet.getRange(employeeRow, colIndex).setValue(currentValue - daysRequested);


    // Log the action
    let logSheet = ss.getSheetByName("Log");
    if (!logSheet) {
      logSheet = ss.insertSheet("Log");
      logSheet.appendRow(["Timestamp", "User Code", "Action", "Details"]);
    }
    logSheet.appendRow([
      timestamp,
      employee.code,
      "submit_leave",
      `Reference: ${refNum}, Record: ${JSON.stringify(record)}, Remaining: ${JSON.stringify(remainingBalances)}`
    ]);


    lock.releaseLock();
    return { success: true, message: "บันทึกการลาสำเร็จ", refNum: refNum };
  } catch (error) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Error in submitLeaveRecord:", error);
    return { success: false, message: "เกิดข้อผิดพลาดในการบันทึกการลา" };
  }
}


function getLeaveColumnIndex(leaveType) {
  switch(leaveType) {
    case "สาย": return 5;
    case "ลาป่วย": return 6;
    case "ลากิจ": return 7;
    case "ลาพักผ่อน": return 8;
    case "ลาคลอด": return 9;
    case "ลาบวช": return 10;
    default: return 0;
  }
}


function getLeaveHistory(code, leaveType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    const lastRow = leaveSheet.getLastRow();
    let history = [];
    if (lastRow > 1) {
      const leaveData = leaveSheet.getRange(2, 1, lastRow - 1, 10).getValues();
      history = leaveData
        .filter(row => String(row[1]).trim() === code && String(row[5]).trim() === leaveType)
        .map(row => ({
          start: row[6],
          finish: row[7],
          total: parseFloat(row[8]) || 0,
          remark: row[9] || "ไม่ระบุ",
          timestamp: row[0]
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Sort by latest date first
    }
    return history;
  } catch (error) {
    console.error("Error in getLeaveHistory:", error);
    return [];
  }
}


// New functions for approval system


function getApprovalData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let leaveSheet = ss.getSheetByName('LeaveRecords');
  let data = leaveSheet.getDataRange().getDisplayValues();
  let appSheet = ss.getSheetByName('app');
  let admins = appSheet.getDataRange().getValues().slice(1).map(r => r[0]);
  return {data, admins};
}


function checkApproverLogin(name, pass) {
  let appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('app');
  let data = appSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && data[i][1] === pass) {
      let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
      logSheet.appendRow([new Date(), name, 'Logged in']);
      return true;
    }
  }
  return false;
}


function setApprovalStatus(refNum, value, user) {
  let leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LeaveRecords');
  let data = leaveSheet.getDataRange().getDisplayValues();
  let ids = data.map(r => r[10]); // refNum in column K (index 10)
  let index = ids.indexOf(refNum);
  if (index === -1) return 'Error: RefNum not found';
  let rowNum = index + 1; // 1-based row
  let appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('app');
  let admins = appSheet.getDataRange().getValues().slice(1).map(r => r[0]);
  let userIndex = admins.indexOf(user);
  if (userIndex === -1 || userIndex > 1) return 'Error: User not authorized';
  let statusCol = 19 + userIndex * 2; // S=19 (1-based), U=21
  leaveSheet.getRange(rowNum, statusCol).setValue(value);
  let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  logSheet.appendRow([new Date(), user, `Set status for Ref ${refNum} to ${value}`]);
  return leaveSheet.getRange(rowNum, statusCol).getValue();
}


function setApprovalRemark(refNum, text, user) {
  let leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LeaveRecords');
  let data = leaveSheet.getDataRange().getDisplayValues();
  let ids = data.map(r => r[10]);
  let index = ids.indexOf(refNum);
  if (index === -1) return 'Error: RefNum not found';
  let rowNum = index + 1;
  let appSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('app');
  let admins = appSheet.getDataRange().getValues().slice(1).map(r => r[0]);
  let userIndex = admins.indexOf(user);
  if (userIndex === -1 || userIndex > 1) return 'Error: User not authorized';
  let remarkCol = 20 + userIndex * 2; // T=20, V=22
  leaveSheet.getRange(rowNum, remarkCol).setValue(text);
  let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  logSheet.appendRow([new Date(), user, `Set remark for Ref ${refNum} to ${text}`]);
  return leaveSheet.getRange(rowNum, remarkCol).getValue();
}  


function getLeaveRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    const data = leaveSheet.getRange(2, 1, leaveSheet.getLastRow() - 1, 22).getValues();
    return data;
  } catch (error) {
    console.error("Error in getLeaveRecords:", error);
    return [];
  }
}


function updateEmployee(code, updates) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: "ระบบกำลัง忙" };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employees");
    const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < employeeData.length; i++) {
      if (String(employeeData[i][0]).trim() === code) {
        row = i + 2;
        break;
      }
    }
    if (row === -1) {
      lock.releaseLock();
      return { success: false, message: "ไม่พบพนักงาน" };
    }
    const fields = {
      late: 5,
      sick: 6,
      business: 7,
      vacation: 8,
      maternity: 9,
      ordination: 10
    };
    for (let key in updates) {
      if (fields[key] && !isNaN(updates[key])) {
        employeeSheet.getRange(row, fields[key]).setValue(parseFloat(updates[key]));
      }
    }
    lock.releaseLock();
    return { success: true };
  } catch (error) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Error in updateEmployee:", error);
    return { success: false, message: "เกิดข้อผิดพลาด" };
  }
}


function updateLeaveRecord(refNum, updates) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: "ระบบกำลัง忙" };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leaveSheet = ss.getSheetByName("LeaveRecords");
    const refData = leaveSheet.getRange(2, 11, leaveSheet.getLastRow() - 1, 1).getValues();
    let row = -1;
    for (let i = 0; i < refData.length; i++) {
      if (String(refData[i][0]).trim() === refNum) {
        row = i + 2;
        break;
      }
    }
    if (row === -1) {
      lock.releaseLock();
      return { success: false, message: "ไม่พบ record" };
    }
    const fields = {
      type: 6, // F
      start: 7, // G
      finish: 8, // H
      total: 9, // I
      remark: 10, // J
      remaining_late: 12, // L
      remaining_sick: 13, // M
      remaining_business: 14, // N
      remaining_vacation: 15, // O
      remaining_maternity: 16, // P
      remaining_ordination: 17 // Q
    };
    for (let key in updates) {
      if (fields[key]) {
        let value = updates[key];
        if (['start', 'finish'].includes(key)) {
          value = new Date(value);
        } else if (['total', 'remaining_late', 'remaining_sick', 'remaining_business', 'remaining_vacation', 'remaining_maternity', 'remaining_ordination'].includes(key)) {
          value = parseFloat(value);
        }
        leaveSheet.getRange(row, fields[key]).setValue(value);
      }
    }
    lock.releaseLock();
    return { success: true };
  } catch (error) {
    if (lock.hasLock()) lock.releaseLock();
    console.error("Error in updateLeaveRecord:", error);
    return { success: false, message: "เกิดข้อผิดพลาด" };
  }
}
