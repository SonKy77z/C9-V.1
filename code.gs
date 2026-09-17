const SHEET_NAME = "Bookings";
// ใส่ Spreadsheet ID ของคุณลงในช่องนี้ (หากเป็นสคริปต์ที่สร้างจาก Sheet โดยตรง ปล่อยว่างไว้ได้)
const SPREADSHEET_ID = "";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getDbSheet() {
  const ss = getSpreadsheet();
  if (!ss) {
    throw new Error("ไม่พบ Google Sheet กรุณาใส่ SPREADSHEET_ID");
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "ID", "ห้อง", "วันที่", "เวลาเริ่ม", "เวลาสิ้นสุด", 
      "รหัสนักศึกษา", "ชื่อ-นามสกุล", "เบอร์โทร", "รายละเอียดกิจกรรม", 
      "รหัสผ่านยกเลิก", "วันที่ทำรายการ"
    ]);
  }
  return sheet;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบจองห้อง สำนักวิชาเภสัชศาสตร์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ดึงข้อมูลการจองทั้งหมด โดยอ่านแบบ DisplayValues เพื่อป้องกัน Format วันที่เพี้ยน
function getBookings() {
  try {
    const sheet = getDbSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    // getDisplayValues จะได้ข้อความ String เหมือนที่ตาเห็นบนหน้าจอชีทเป๊ะๆ
    const rows = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    
    const bookings = rows.map(r => {
      // จัดการฟอร์แมตวันที่ให้เป็น YYYY-MM-DD เสมอ
      let rawDate = String(r[2]).trim();
      let formattedDate = rawDate;
      if (rawDate.includes('/')) {
        const parts = rawDate.split('/');
        if (parts.length === 3) {
          // หาก Sheet แสดงเป็น DD/MM/YYYY
          if (parts[2].length === 4) {
            formattedDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
          }
        }
      }

      return {
        id: String(r[0]).trim(),
        room: String(r[1]).trim(),
        date: formattedDate,
        startTime: String(r[3]).trim().substring(0, 5), // ตัดเฉพาะ HH:mm
        endTime: String(r[4]).trim().substring(0, 5),
        studentId: String(r[5]).trim(),
        name: String(r[6]).trim(),
        phone: String(r[7]).trim(),
        reason: String(r[8]).trim(),
        hasPasscode: String(r[9]).trim().length > 0,
        createdAt: String(r[10]).trim()
      };
    }).filter(b => b.id && b.room && b.date);

    return bookings;
  } catch (err) {
    Logger.log("Error getBookings: " + err.message);
    return [];
  }
}

// บันทึกการจอง
function createBooking(data) {
  try {
    const sheet = getDbSheet();
    const rows = sheet.getLastRow() > 1 
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getDisplayValues() 
      : [];
    
    const targetDate = String(data.date).trim();
    const targetStdId = String(data.studentId).trim();
    const newStart = parseMinutes(data.startTime);
    const newEnd = parseMinutes(data.endTime);
    const newDuration = newEnd - newStart;

    if (newDuration <= 0) return { success: false, message: "เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น" };
    if (newDuration > 240) return { success: false, message: "จองได้ไม่เกินครั้งละ 4 ชั่วโมง" };

    let totalMinutesToday = 0;
    for (let i = 0; i < rows.length; i++) {
      let rDate = String(rows[i][2]).trim();
      if (rDate.includes('/')) {
        const p = rDate.split('/');
        if (p.length === 3 && p[2].length === 4) rDate = `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
      }

      const rRoom = String(rows[i][1]).trim();
      const rStart = parseMinutes(rows[i][3]);
      const rEnd = parseMinutes(rows[i][4]);
      const rStdId = String(rows[i][5]).trim();

      if (rDate === targetDate) {
        // ตรวจสอบห้องซ้ำและเวลาชน
        if (rRoom === String(data.room).trim()) {
          if (newStart < rEnd && newEnd > rStart) {
            return { success: false, message: `ห้อง ${data.room} ติดจองแล้วในช่วงเวลาดังกล่าว` };
          }
        }
        // ตรวจสอบโควตา 4 ชม./วัน (240 นาที)
        if (rStdId === targetStdId) {
          totalMinutesToday += (rEnd - rStart);
        }
      }
    }

    if (totalMinutesToday + newDuration > 240) {
      const remainMins = Math.max(0, 240 - totalMinutesToday);
      return { 
        success: false, 
        message: `คุณใช้โควตาไปแล้ว ${totalMinutesToday / 60} ชม. ในวันที่เลือก สามารถจองเพิ่มได้อีกไม่เกิน ${remainMins / 60} ชม. เท่านั้น (โควตา 4 ชม./วัน)` 
      };
    }

    const newId = "BK-" + new Date().getTime();
    const nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

    // บันทึกลง Sheet เป็น Text เสมอ เพื่อป้องกัน Sheet แปลงเป็น Time Format เพี้ยน
    sheet.appendRow([
      "'" + newId,
      "'" + String(data.room).trim(),
      "'" + targetDate,
      "'" + String(data.startTime).trim(),
      "'" + String(data.endTime).trim(),
      "'" + targetStdId,
      "'" + String(data.name).trim(),
      "'" + String(data.phone).trim(),
      "'" + String(data.reason).trim(),
      data.passcode ? "'" + String(data.passcode).trim() : "",
      "'" + nowStr
    ]);

    return { success: true, message: "จองห้องสำเร็จ!", id: newId };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

// ขอยกเลิกการจอง
function cancelBooking(bookingId, inputPasscode) {
  try {
    const sheet = getDbSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "ไม่พบข้อมูลการจอง" };

    const rows = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    let targetRowIndex = -1;
    let savedPasscode = "";

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(bookingId).trim()) {
        targetRowIndex = i + 2;
        savedPasscode = String(rows[i][9]).trim();
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: "ไม่พบรายการจองนี้ในระบบ" };
    }

    // ตรวจสอบรหัสผ่าน
    if (savedPasscode !== "") {
      if (!inputPasscode || String(inputPasscode).trim() !== savedPasscode) {
        return { success: false, message: "รหัสผ่านยกเลิกไม่ถูกต้อง!" };
      }
    }

    sheet.deleteRow(targetRowIndex);
    return { success: true, message: "ยกเลิกการจองสำเร็จ" };
  } catch (err) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + err.message };
  }
}

function parseMinutes(timeStr) {
  if (!timeStr) return 0;
  const parts = String(timeStr).split(":");
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1] || 0, 10);
}
