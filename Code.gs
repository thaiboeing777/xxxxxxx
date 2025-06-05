/**
 * โค้ดสำหรับระบบลงเวลาทำงาน Time Attendance
 * @OnlyCurrentDoc
 */

// Global variables
const SHEET_NAME_ATTENDANCE = "TimeAttendance";
const SHEET_NAME_EMPLOYEES = "Employees";
const SHEET_NAME_SHIFTS = "Shifts";
const SHEET_NAME_SETTINGS = "Settings";
const TELEGRAM_TOKEN = "7919596148:AAFMnaADqVlH4Kz5PZ2dmATs9EnhEQQhQ4Q";
const TELEGRAM_CHAT_ID = "-4917185874";
const TIME_ZONES = {
  "UTC": "UTC",
  "Asia/Bangkok": "Asia/Bangkok",
  "Asia/Dhaka": "Asia/Dhaka"  // Bangladesh
};

/**
 * เปิดหน้า web app
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Time Attendance System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * แทรกโค้ด JavaScript และ CSS
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ดึงข้อมูล Time Zones ที่รองรับ
 */
function getTimeZones() {
  return TIME_ZONES;
}

/**
 * ดึงข้อมูลพนักงานทั้งหมด
 */
function getEmployeeList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
    
    if (!sheet) {
      Logger.log("Sheet not found: " + SHEET_NAME_EMPLOYEES);
      return [];
    }
    
    const range = sheet.getDataRange();
    const data = range.getValues();
    
    if (data.length <= 1) {
      Logger.log("No employee data found");
      return [];
    }
    
    // หา index ของคอลัมน์ต่างๆ
    const headers = data[0];
    const persNoIndex = headers.indexOf("PersNo");
    const fullNameIndex = headers.indexOf("FullName");
    const departmentIndex = headers.indexOf("Department");
    const positionIndex = headers.indexOf("Position");
    const profileImageUrlIndex = headers.indexOf("ProfileImageURL");
    
    if (persNoIndex === -1 || fullNameIndex === -1) {
      Logger.log("Required columns not found");
      return [];
    }
    
    // สร้าง array ของพนักงาน
    const employees = data.slice(1).map(row => {
      return {
        persNo: row[persNoIndex].toString(),
        fullName: row[fullNameIndex].toString(),
        department: departmentIndex !== -1 ? row[departmentIndex] || "" : "",
        position: positionIndex !== -1 ? row[positionIndex] || "" : "",
        profileImageUrl: profileImageUrlIndex !== -1 ? row[profileImageUrlIndex] || "" : ""
      };
    }).filter(emp => emp.persNo && emp.fullName);
    
    Logger.log("Found " + employees.length + " employees");
    return employees;
  } catch (error) {
    Logger.log("Error in getEmployeeList: " + error.message);
    return [];
  }
}

/**
 * ดึงข้อมูลแผนก
 */
function getDepartments() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
    
    if (!sheet) {
      Logger.log("Sheet not found: " + SHEET_NAME_EMPLOYEES);
      return [];
    }
    
    const range = sheet.getDataRange();
    const data = range.getValues();
    
    if (data.length <= 1) {
      Logger.log("No employee data found");
      return [];
    }
    
    // หา index ของคอลัมน์แผนก
    const headers = data[0];
    const departmentIndex = headers.indexOf("Department");
    
    if (departmentIndex === -1) {
      Logger.log("Department column not found");
      return [];
    }
    
    // ดึงข้อมูลแผนกจากทุก row และกรองเฉพาะค่าที่ไม่ซ้ำ
    const departments = data.slice(1)
      .map(row => row[departmentIndex])
      .filter(dept => dept !== null && dept !== "")
      .filter((dept, index, self) => self.indexOf(dept) === index)
      .sort();
    
    Logger.log("Found " + departments.length + " departments");
    return departments;
  } catch (error) {
    Logger.log("Error in getDepartments: " + error.message);
    return [];
  }
}

/**
 * ดึงข้อมูลกะการทำงานจาก Sheet Shifts
 */
function getShifts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shiftsSheet = ss.getSheetByName(SHEET_NAME_SHIFTS);
    
    if (!shiftsSheet) {
      Logger.log("Sheet 'Shifts' not found");
      return [];
    }
    
    const dataRange = shiftsSheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) {
      Logger.log("No shift data found in 'Shifts' sheet");
      return [];
    }
    
    const headers = data[0];
    const shiftNameIndex = headers.indexOf("ShiftName");
    const shiftIdIndex = headers.indexOf("ShiftId");
    const startTimeIndex = headers.indexOf("StartTime");
    const endTimeIndex = headers.indexOf("EndTime");
    
    if (shiftNameIndex === -1) {
      Logger.log("Column 'ShiftName' not found in 'Shifts' sheet");
      return [];
    }
    
    const shifts = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[shiftNameIndex] && row[shiftNameIndex].toString().trim() !== "") {
        const shift = {
          shiftName: row[shiftNameIndex].toString()
        };
        
        if (shiftIdIndex !== -1 && row[shiftIdIndex]) {
          shift.shiftId = row[shiftIdIndex].toString();
        }
        
        if (startTimeIndex !== -1 && row[startTimeIndex]) {
          // จัดรูปแบบ StartTime ให้เป็น HH:MM
          const startTime = row[startTimeIndex];
          if (typeof startTime === 'object' && startTime instanceof Date) {
            const hours = startTime.getHours().toString().padStart(2, '0');
            const minutes = startTime.getMinutes().toString().padStart(2, '0');
            shift.startTime = `${hours}:${minutes}`;
          } else {
            shift.startTime = row[startTimeIndex].toString();
          }
        }
        
        if (endTimeIndex !== -1 && row[endTimeIndex]) {
          // จัดรูปแบบ EndTime ให้เป็น HH:MM
          const endTime = row[endTimeIndex];
          if (typeof endTime === 'object' && endTime instanceof Date) {
            const hours = endTime.getHours().toString().padStart(2, '0');
            const minutes = endTime.getMinutes().toString().padStart(2, '0');
            shift.endTime = `${hours}:${minutes}`;
          } else {
            shift.endTime = row[endTimeIndex].toString();
          }
        }
        
        shifts.push(shift);
      }
    }
    
    Logger.log("Found " + shifts.length + " shifts");
    return shifts;
  } catch (error) {
    Logger.log("Error in getShifts: " + error.message);
    return [];
  }
}

/**
 * ตรวจสอบความถูกต้องของตำแหน่ง
 */
function checkLocationValidity(latitude, longitude) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
    
    if (!settingsSheet) {
      return { isValid: true, distance: 0, message: "No settings found, location check skipped." };
    }
    
    const settingsData = settingsSheet.getDataRange().getValues();
    
    let allowedLatitude = null;
    let allowedLongitude = null;
    let maxDistance = 100;
    let locationName = "Unknown";
    
    for (let i = 1; i < settingsData.length; i++) {
      const key = settingsData[i][0];
      const value = settingsData[i][1];
      
      if (key === "AllowedLatitude" && value) {
        allowedLatitude = parseFloat(value);
      } else if (key === "AllowedLongitude" && value) {
        allowedLongitude = parseFloat(value);
      } else if (key === "MaxDistance" && value) {
        maxDistance = parseFloat(value);
      } else if (key === "LocationName" && value) {
        locationName = value;
      }
    }
    
    if (allowedLatitude === null || allowedLongitude === null) {
      return { isValid: true, distance: 0, locationName: locationName, message: "No location constraints set." };
    }
    
    const distance = calculateDistance(latitude, longitude, allowedLatitude, allowedLongitude);
    const isValid = distance <= maxDistance;
    
    return {
      isValid: isValid,
      distance: Math.round(distance),
      locationName: locationName,
      message: isValid ? "Location is valid." : "Location is too far from the allowed location."
    };
  } catch (error) {
    Logger.log("Error in checkLocationValidity: " + error.message);
    return { isValid: false, distance: 0, message: "Error checking location: " + error.message };
  }
}

/**
 * คำนวณระยะห่างระหว่างสองพิกัด (ในหน่วยเมตร)
 */
function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lon2 - lon1) * Math.PI / 180;

  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  const d = R * c;
  return d;
}

/**
 * แปลงวันที่เป็นรูปแบบ DD-MM-YYYY
 */
function formatDateDDMMYYYY(dateStr) {
  const parts = dateStr.split('-');
  
  if (parts.length !== 3) {
    return dateStr;
  }
  
  const year = parts[0];
  const month = parts[1];
  const day = parts[2];
  
  return `${day}-${month}-${year}`;
}

/**
 * แปลงเวลาตาม Time Zone
 */
function adjustTimeForZone(timeStr, timeZone) {
  try {
    // แยกชั่วโมง นาที วินาที
    const timeParts = timeStr.split(':');
    const hours = parseInt(timeParts[0]);
    const minutes = timeParts[1];
    const seconds = timeParts[2];
    
    let adjustedHours = hours;
    
    // ปรับเวลาตาม Time Zone
    if (timeZone === "UTC") {
      // ไม่ต้องปรับ
      adjustedHours = hours;
    } else if (timeZone === "Asia/Bangkok") {
      // UTC+7
      adjustedHours = (hours + 7) % 24;
    } else if (timeZone === "Asia/Dhaka") {
      // UTC+6
      adjustedHours = (hours + 6) % 24;
    }
    
    // จัดรูปแบบให้เป็น HH:MM:SS
    return `${adjustedHours.toString().padStart(2, '0')}:${minutes}:${seconds}`;
  } catch (error) {
    Logger.log("Error adjusting time for zone: " + error.message);
    return timeStr;
  }
}

/**
 * ดึงเวลาปัจจุบันตาม Time Zone
 */
function getCurrentDateTime(timeZone) {
  try {
    // ใช้เวลาปัจจุบันจริง
    const now = new Date();
    const year = now.getUTCFullYear();
    const month = (now.getUTCMonth() + 1).toString().padStart(2, '0');
    const day = now.getUTCDate().toString().padStart(2, '0');
    const hours = now.getUTCHours().toString().padStart(2, '0');
    const minutes = now.getUTCMinutes().toString().padStart(2, '0');
    const seconds = now.getUTCSeconds().toString().padStart(2, '0');
    
    // สร้างรูปแบบวันที่และเวลาในรูปแบบ YYYY-MM-DD HH:MM:SS (UTC)
    const datePart = `${year}-${month}-${day}`;
    const timePart = `${hours}:${minutes}:${seconds}`;
    
    // ปรับเวลาตาม Time Zone
    const adjustedTimePart = adjustTimeForZone(timePart, timeZone);
    
    // แปลงวันที่เป็นรูปแบบ DD-MM-YYYY
    const formattedDate = formatDateDDMMYYYY(datePart);
    
    return {
      dateTimeStr: datePart + " " + adjustedTimePart,
      datePart: datePart,
      timePart: adjustedTimePart,
      formattedDate: formattedDate
    };
  } catch (error) {
    Logger.log("Error getting current date time: " + error.message);
    return {
      dateTimeStr: "Error",
      datePart: "Error",
      timePart: "Error",
      formattedDate: "Error"
    };
  }
}

/**
 * บันทึกข้อมูลการลงเวลา
 */
function processTimeAttendance(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_ATTENDANCE);
    
    if (!sheet) {
      return { success: false, message: "TimeAttendance sheet not found." };
    }
    
    // บันทึก log ข้อมูลที่ได้รับจาก client
    Logger.log("Received data from client:");
    Logger.log("PersNo: " + data.persNo);
    Logger.log("FullName: " + data.fullName);
    Logger.log("ShiftName: " + data.shiftName);
    Logger.log("StartTime: " + data.startTime);
    Logger.log("EndTime: " + data.endTime);
    Logger.log("ClockStatus: " + data.clockStatus);
    Logger.log("TimeZone: " + data.timeZone);
    
    // ใช้เวลาปัจจุบัน
    const dateTimeInfo = getCurrentDateTime(data.timeZone);
    Logger.log("Current Date/Time: " + dateTimeInfo.dateTimeStr + " (" + data.timeZone + ")");
    
    // สร้าง shiftData object จากข้อมูลที่ได้รับจาก client
    const shiftData = {
      startTime: data.startTime || "",
      endTime: data.endTime || ""
    };
    
    // แยกเวลาปัจจุบันเป็นชั่วโมงและนาที
    const currentHour = parseInt(dateTimeInfo.timePart.split(":")[0]);
    const currentMinute = parseInt(dateTimeInfo.timePart.split(":")[1]);
    const currentTimeInMinutes = currentHour * 60 + currentMinute;
    
    // วิเคราะห์สถานะการมาทำงานตามกฎที่กำหนด
    const statusDetails = analyzeAttendanceStatus(
      currentTimeInMinutes,
      data.clockStatus,
      shiftData
    );
    
    // สร้าง Attendance ID
    const lastRow = sheet.getLastRow();
    let attendanceId = 1;
    if (lastRow > 1) {
      const lastAttendanceId = sheet.getRange(lastRow, 1).getValue();
      if (typeof lastAttendanceId === 'number') {
        attendanceId = lastAttendanceId + 1;
      }
    }
    
    // บันทึกรูปภาพลง Google Drive และสร้าง URL
    const imageUrl = saveImageToDrive(data.photoBlob, data.persNo, data.clockStatus);
    
    // กำหนดข้อมูลที่จะบันทึก
    if (data.clockStatus === "ClockIn") {
      // บันทึกข้อมูลโดยตรง แทนการใช้ appendRow
      const newRowNum = sheet.getLastRow() + 1;
      
      // บันทึกข้อมูลทีละคอลัมน์
      sheet.getRange(newRowNum, 1).setValue(attendanceId);
      sheet.getRange(newRowNum, 2).setValue(data.persNo);
      sheet.getRange(newRowNum, 3).setValue(data.fullName);
      sheet.getRange(newRowNum, 4).setValue(dateTimeInfo.formattedDate); // วันที่ในรูปแบบ DD-MM-YYYY
      sheet.getRange(newRowNum, 5).setValue(data.shiftName); // ShiftName - บันทึกชื่อกะที่เลือก
      sheet.getRange(newRowNum, 6).setValue(dateTimeInfo.timePart);  // บันทึกเวลา
      sheet.getRange(newRowNum, 7).setValue("");
      sheet.getRange(newRowNum, 8).setValue(data.latitude);
      sheet.getRange(newRowNum, 9).setValue(data.longitude);
      sheet.getRange(newRowNum, 10).setValue(data.distance);
      sheet.getRange(newRowNum, 11).setValue(statusDetails.status); // ClockInStatus ตามกฎ
      sheet.getRange(newRowNum, 12).setValue("");
      sheet.getRange(newRowNum, 13).setValue(0);
      sheet.getRange(newRowNum, 14).setValue(0);
      sheet.getRange(newRowNum, 15).setValue(imageUrl);
      
      Logger.log("ClockIn record saved with ShiftName: " + data.shiftName + " at row: " + newRowNum);
      Logger.log("Time saved: " + dateTimeInfo.timePart);
      Logger.log("Status saved: " + statusDetails.status);
      
      // ส่งข้อความแจ้งเตือนผ่าน Telegram
      sendTelegramNotification(data, dateTimeInfo.formattedDate, dateTimeInfo.timePart, statusDetails, imageUrl, data.latitude, data.longitude, data.timeZone);
    } else {
      // ค้นหาข้อมูล ClockIn ล่าสุดของพนักงานคนนี้
      const data2D = sheet.getDataRange().getValues();
      let clockInRow = -1;
      
      for (let i = data2D.length - 1; i > 0; i--) {
        if (data2D[i][1] === data.persNo && data2D[i][6] === "") {
          clockInRow = i + 1;
          break;
        }
      }
      
      if (clockInRow > 0) {
        // อัปเดตข้อมูล ClockOut ในแถวที่มี ClockIn
        sheet.getRange(clockInRow, 7).setValue(dateTimeInfo.timePart); // ClockOutTime
        sheet.getRange(clockInRow, 12).setValue(statusDetails.status); // ClockOutStatus ตามกฎ
        
        // คำนวณชั่วโมงทำงาน
        const clockInTime = sheet.getRange(clockInRow, 6).getValue();
        const workHours = calculateWorkHours(clockInTime, dateTimeInfo.timePart);
        
        sheet.getRange(clockInRow, 13).setValue(workHours); // WorkHours
        
        // อัปเดตรูปภาพ ClockOut
        const currentImageUrl = sheet.getRange(clockInRow, 15).getValue();
        const newImageUrl = currentImageUrl + ", " + imageUrl;
        sheet.getRange(clockInRow, 15).setValue(newImageUrl);
        
        Logger.log("ClockOut updated for existing record at row: " + clockInRow);
        Logger.log("Time saved: " + dateTimeInfo.timePart);
        Logger.log("Status saved: " + statusDetails.status);
        
        // ส่งข้อความแจ้งเตือนผ่าน Telegram
        sendTelegramNotification(data, dateTimeInfo.formattedDate, dateTimeInfo.timePart, statusDetails, imageUrl, data.latitude, data.longitude, data.timeZone);
        
        return {
          success: true,
          dateTime: dateTimeInfo.dateTimeStr,
          formattedDate: dateTimeInfo.formattedDate,
          statusDetails: statusDetails,
          timeZone: data.timeZone
        };
      } else {
        // ไม่พบข้อมูล ClockIn สร้างแถวใหม่สำหรับ ClockOut
        const newRowNum = sheet.getLastRow() + 1;
        
        // บันทึกข้อมูลทีละคอลัมน์
        sheet.getRange(newRowNum, 1).setValue(attendanceId);
        sheet.getRange(newRowNum, 2).setValue(data.persNo);
        sheet.getRange(newRowNum, 3).setValue(data.fullName);
        sheet.getRange(newRowNum, 4).setValue(dateTimeInfo.formattedDate); // วันที่ในรูปแบบ DD-MM-YYYY
        sheet.getRange(newRowNum, 5).setValue(data.shiftName); // ShiftName - บันทึกชื่อกะที่เลือก
        sheet.getRange(newRowNum, 6).setValue("");
        sheet.getRange(newRowNum, 7).setValue(dateTimeInfo.timePart); // ClockOutTime
        sheet.getRange(newRowNum, 8).setValue(data.latitude);
        sheet.getRange(newRowNum, 9).setValue(data.longitude);
        sheet.getRange(newRowNum, 10).setValue(data.distance);
        sheet.getRange(newRowNum, 11).setValue("");
        sheet.getRange(newRowNum, 12).setValue(statusDetails.status); // ClockOutStatus ตามกฎ
        sheet.getRange(newRowNum, 13).setValue(0);
        sheet.getRange(newRowNum, 14).setValue(0);
        sheet.getRange(newRowNum, 15).setValue(imageUrl);
        
        Logger.log("New ClockOut record created with ShiftName: " + data.shiftName + " at row: " + newRowNum);
        Logger.log("Time saved: " + dateTimeInfo.timePart);
        Logger.log("Status saved: " + statusDetails.status);
        
        // ส่งข้อความแจ้งเตือนผ่าน Telegram
        sendTelegramNotification(data, dateTimeInfo.formattedDate, dateTimeInfo.timePart, statusDetails, imageUrl, data.latitude, data.longitude, data.timeZone);
      }
    }
    
    return {
      success: true,
      dateTime: dateTimeInfo.dateTimeStr,
      formattedDate: dateTimeInfo.formattedDate,
      statusDetails: statusDetails,
      timeZone: data.timeZone
    };
  } catch (error) {
    Logger.log("Error in processTimeAttendance: " + error.message);
    Logger.log("Error stack: " + error.stack);
    return { success: false, message: "Error processing attendance: " + error.message };
  }
}

/**
 * วิเคราะห์สถานะการมาทำงานตามกฎที่กำหนด
 */
function analyzeAttendanceStatus(currentTimeInMinutes, clockStatus, shiftData) {
  try {
    if (!shiftData.startTime && !shiftData.endTime) {
      Logger.log("Missing shift time data, returning N/A status");
      return { status: "N/A", difference: 0 };
    }
    
    Logger.log("Analyzing status - clockStatus: " + clockStatus);
    Logger.log("Shift times - start: " + shiftData.startTime + ", end: " + shiftData.endTime);
    Logger.log("Current time in minutes: " + currentTimeInMinutes);
    
    // แปลงเวลาเริ่มและจบกะเป็นนาที
    let startTimeInMinutes = 0;
    let endTimeInMinutes = 0;
    
    if (shiftData.startTime) {
      const startTimeParts = shiftData.startTime.split(":");
      if (startTimeParts.length >= 2) {
        startTimeInMinutes = parseInt(startTimeParts[0]) * 60 + parseInt(startTimeParts[1]);
        Logger.log("Start time in minutes: " + startTimeInMinutes);
      } else {
        Logger.log("Invalid start time format: " + shiftData.startTime);
      }
    }
    
    if (shiftData.endTime) {
      const endTimeParts = shiftData.endTime.split(":");
      if (endTimeParts.length >= 2) {
        endTimeInMinutes = parseInt(endTimeParts[0]) * 60 + parseInt(endTimeParts[1]);
        Logger.log("End time in minutes: " + endTimeInMinutes);
      } else {
        Logger.log("Invalid end time format: " + shiftData.endTime);
      }
    }
    
    // ตรวจสอบกรณีข้ามวัน (เช่น กะที่เริ่ม 22:00 และจบ 06:00)
    if (endTimeInMinutes < startTimeInMinutes) {
      endTimeInMinutes += 24 * 60; // เพิ่ม 24 ชั่วโมง
      if (currentTimeInMinutes < startTimeInMinutes) {
        currentTimeInMinutes += 24 * 60; // เพิ่ม 24 ชั่วโมง ถ้าเวลาปัจจุบันอยู่ในวันถัดไป
      }
      Logger.log("Adjusted for overnight shift - end time: " + endTimeInMinutes + ", current time: " + currentTimeInMinutes);
    }
    
    try {
      if (clockStatus === "ClockIn") {
        // คำนวณความต่างของเวลาปัจจุบันกับเวลาเริ่มกะ (นาที)
        const differenceFromStart = startTimeInMinutes - currentTimeInMinutes;
        Logger.log("Difference from start time: " + differenceFromStart + " minutes");
        
        if (differenceFromStart > 120) {
          // มาเร็วกว่าเวลาเริ่มกะมากกว่า 120 นาที
          return { status: "Early Clock In", difference: differenceFromStart };
        } else if (differenceFromStart >= 0) {
          // มาก่อนเวลาเริ่มกะไม่เกิน 120 นาที (ถูกต้องตามกฎ)
          return { status: "On Time", difference: 0 };
        } else {
          // มาหลังเวลาเริ่มกะ
          const lateMinutes = Math.abs(differenceFromStart);
          if (lateMinutes > 15) {
            // สาย (หลังเวลาเริ่มกะเกิน 15 นาที)
            return { status: "Late Clock In", difference: lateMinutes };
          } else {
            // สายไม่เกิน 15 นาที ถือว่า On Time
            return { status: "On Time", difference: 0 };
          }
        }
      } else if (clockStatus === "ClockOut") {
        // คำนวณความต่างของเวลาปัจจุบันกับเวลาจบกะ (นาที)
        const differenceFromEnd = currentTimeInMinutes - endTimeInMinutes;
        Logger.log("Difference from end time: " + differenceFromEnd + " minutes");
        
        if (differenceFromEnd > 120) {
          // ออกช้ากว่าเวลาจบกะมากกว่า 120 นาที
          return { status: "Late Clock Out", difference: differenceFromEnd };
        } else if (differenceFromEnd >= 0) {
          // ออกหลังเวลาจบกะไม่เกิน 120 นาที (ถูกต้องตามกฎ)
          return { status: "On Time", difference: 0 };
        } else {
          // ออกก่อนเวลาจบกะ
          const earlyMinutes = Math.abs(differenceFromEnd);
          if (earlyMinutes > 15) {
            // ออกก่อนเวลาจบกะเกิน 15 นาที
            return { status: "Early Clock Out", difference: earlyMinutes };
          } else {
            // ออกก่อนไม่เกิน 15 นาที ถือว่า On Time
            return { status: "On Time", difference: 0 };
          }
        }
      } else {
        Logger.log("Invalid clock status: " + clockStatus);
        return { status: "Error", difference: 0 };
      }
    } catch (innerError) {
      Logger.log("Error in calculation: " + innerError.message);
      return { status: "N/A", difference: 0 };
    }
  } catch (error) {
    Logger.log("Error in analyzeAttendanceStatus: " + error.message);
    Logger.log("Error stack: " + error.stack);
    return { status: "N/A", difference: 0 };
  }
}

/**
 * ส่งการแจ้งเตือนผ่าน Telegram
 */
function sendTelegramNotification(data, formattedDate, timePart, statusDetails, imageUrl, latitude, longitude, timeZone) {
  try {
    // สร้างข้อความแจ้งเตือน - ใช้รูปแบบเวลาที่ถูกต้อง
    // แสดง Work Time ในรูปแบบ StartTime(HH:MM) - EndTime(HH:MM)
    const workTime = `${data.startTime} - ${data.endTime}`;
    
    const message = `
📢 New ${data.clockStatus} Recorded 📢
👨‍✈️ Name         : ${data.fullName}
📂 Department   : ${data.department}
🧑‍💼 Position      : ${data.position}
🕘 Work Shift   : ${data.shiftName}
⏰ Work Time    : ${workTime}
✏️ Clock Status : ${data.clockStatus}
⏰ Date/Time    : ${formattedDate} ${timePart}
🌐 Time Zone    : ${timeZone || "UTC"}
📍 Location     : ${latitude}, ${longitude}
🚩 Distance     : ${data.distance} Meters
⏳ OnDuty Status : ${statusDetails.status}${statusDetails.difference > 0 ? 
`\n💼 ${statusDetails.status.includes("Early") || statusDetails.status.includes("Late") ? statusDetails.status + " " : ""}   : ${statusDetails.difference} Minutes` : ''}
`;
    
    // ส่งข้อความ
    const textUrl = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage?chat_id=${TELEGRAM_CHAT_ID}&text=${encodeURIComponent(message)}`;
    UrlFetchApp.fetch(textUrl);
    
    // ส่งรูปภาพ (ถ้ามี)
    if (imageUrl) {
      const photoUrl = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendPhoto?chat_id=${TELEGRAM_CHAT_ID}&photo=${encodeURIComponent(imageUrl)}&caption=${encodeURIComponent(data.fullName + " - " + data.clockStatus)}`;
      UrlFetchApp.fetch(photoUrl);
    }
    
    Logger.log("Telegram notification sent successfully");
    return true;
  } catch (error) {
    Logger.log("Error sending Telegram notification: " + error.message);
    return false;
  }
}

/**
 * บันทึกรูปภาพลง Google Drive และคืนค่า URL
 */
function saveImageToDrive(base64Image, persNo, clockStatus) {
  try {
    const imageData = base64Image.split('base64,')[1] || base64Image;
    const folder = DriveApp.getFolderById("1dqzNN5JJlxaGdaR6doLUYXPMzIIFk0jj");
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const filename = `${persNo}_${clockStatus}_${timestamp}.png`;
    
    const blob = Utilities.newBlob(Utilities.base64Decode(imageData), 'image/png', filename);
    const file = folder.createFile(blob);
    
    return `https://drive.google.com/uc?id=${file.getId()}`;
  } catch (error) {
    Logger.log("Error saving image to Drive: " + error.message);
    return "";
  }
}

/**
 * คำนวณชั่วโมงทำงาน
 */
function calculateWorkHours(clockInTime, clockOutTime) {
  try {
    // ตัดส่วนเวลาเหลือแค่ HH:MM ถ้าเป็นรูปแบบ HH:MM:SS
    const inTimeParts = clockInTime.split(':');
    const outTimeParts = clockOutTime.split(':');
    
    const inHour = parseInt(inTimeParts[0]);
    const inMin = parseInt(inTimeParts[1]);
    
    const outHour = parseInt(outTimeParts[0]);
    const outMin = parseInt(outTimeParts[1]);
    
    let totalMinutes = (outHour * 60 + outMin) - (inHour * 60 + inMin);
    
    if (totalMinutes < 0) {
      totalMinutes += 24 * 60;
    }
    
    return Math.round(totalMinutes / 60 * 100) / 100;
  } catch (error) {
    Logger.log("Error calculating work hours: " + error.message);
    return 0;
  }
}