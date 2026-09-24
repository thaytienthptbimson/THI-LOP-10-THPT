/**
 * HỆ THỐNG KHẢO SÁT & THI TRỰC TUYẾN THPT - BACKEND
 * Mã nguồn này xử lý các request từ file index.html
 */

const ADMIN_SECRET_PIN = "28011981"; // Phải khớp với frontend
const SHEET_RESULTS = "KetQuaThi";
const SHEET_CONFIG = "CauHinhHeThong";

/**
 * Hàm khởi tạo (Chạy hàm này 1 lần đầu tiên để tạo các sheet cần thiết)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tạo sheet lưu kết quả
  let sheetResults = ss.getSheetByName(SHEET_RESULTS);
  if (!sheetResults) {
    sheetResults = ss.insertSheet(SHEET_RESULTS);
    sheetResults.appendRow([
      "Thời gian Server", "Mã Đề (ExamID)", "Khóa Học Sinh (Key)", 
      "Họ Tên", "Lớp", "SBD", "Điểm", "Số câu đúng", 
      "Thời gian làm (s)", "Chuyển tab (lần)", "Thời điểm nộp", "Chi tiết bài làm", "Trạng Thái"
    ]);
    sheetResults.setFrozenRows(1);
    sheetResults.getRange("1:1").setFontWeight("bold").setBackground("#f3f4f6");
  }

  // Tạo sheet lưu cấu hình
  let sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(SHEET_CONFIG);
    sheetConfig.appendRow(["ExamID", "ScheduleJSON"]);
    sheetConfig.setFrozenRows(1);
    sheetConfig.getRange("1:1").setFontWeight("bold").setBackground("#f3f4f6");
  }
}

/**
 * Xử lý các yêu cầu GET (Lấy cấu hình, Kiểm tra trạng thái)
 */
function doGet(e) {
  const p = e.parameter;
  const action = p.action;
  const exam = p.exam || 'vao10';

  try {
    // Trả về cấu hình lịch thi
    if (action === 'getConfig') {
      const schedule = getConfig(exam);
      return jsonResponse({ ok: true, now: new Date().getTime(), schedule: schedule });
    }

    // Kiểm tra học sinh đã thi chưa hoặc ghi nhận bắt đầu thi
    if (action === 'start' || action === 'check') {
      const key = p.key;
      const isUsed = checkStudentExists(exam, key);
      
      if (action === 'check') {
        return jsonResponse({ ok: true, used: isUsed });
      }
      
      // Xử lý action = 'start'
      if (!isUsed) {
        // Ghi nhận trạng thái đang thi để chống mở nhiều tab/máy
        logStudentStatus(exam, key, p.name, p.cls, p.sbd, "STARTED");
        return jsonResponse({ ok: true, allowed: true });
      } else {
        return jsonResponse({ ok: true, allowed: false });
      }
    }

    return jsonResponse({ ok: false, msg: 'Unknown action' });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

/**
 * Xử lý các yêu cầu POST (Nộp bài, Đặt lịch thi, Reset bài thi)
 */
function doPost(e) {
  const p = e.parameter;
  const action = p.action;

  try {
    // 1. CÁC TÁC VỤ QUẢN TRỊ (Yêu cầu mã PIN)
    if (action === 'setConfig' || action === 'reset') {
      if (p.pin !== ADMIN_SECRET_PIN) return jsonResponse({ ok: false, error: 'Mã PIN quản trị không chính xác' });

      // Lưu lịch thi
      if (action === 'setConfig') {
        setConfig(p.exam, p.schedule);
        return jsonResponse({ ok: true });
      }

      // Cho phép thi lại
      if (action === 'reset') {
        resetStudent(p.exam, p.key);
        return jsonResponse({ ok: true });
      }
    }

    // 2. TÁC VỤ NỘP BÀI THI CỦA HỌC SINH
    // Frontend sử dụng mode: 'no-cors' nên không đọc được response, nhưng dữ liệu vẫn được đẩy lên
    if (p.score !== undefined && p.name !== undefined) {
      const key = "k:" + [
        normString(p.name), 
        normString(p.cls), 
        normString(p.sbd)
      ].join('|');
      
      logStudentStatus(
        p.exam || 'vao10', 
        key, 
        p.name, 
        p.cls, 
        p.sbd, 
        "SUBMITTED", 
        p
      );
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, msg: 'Hành động không hợp lệ' });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

/* ==========================================================================
   CÁC HÀM XỬ LÝ DỮ LIỆU BÊN TRONG (INTERNAL FUNCTIONS)
   ========================================================================== */

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Hàm chuẩn hóa chuỗi tạo Khóa (giống y hệt frontend)
function normString(s) {
  if (!s) return "";
  const chars = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  return chars.replace(/đ/gi, 'd').replace(/\s+/g, ' ').trim();
}

function getConfig(examId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === examId) {
      try {
        return data[i][1] ? JSON.parse(data[i][1]) : null;
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

function setConfig(examId, scheduleJson) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === examId) {
      sheet.getRange(i + 1, 2).setValue(scheduleJson);
      return;
    }
  }
  sheet.appendRow([examId, scheduleJson]);
}

function checkStudentExists(examId, key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    // Chỉ kiểm tra cùng mã đề và cùng khóa học sinh
    if (data[i][1] === examId && data[i][2] === key) {
      const status = data[i][12];
      // Nếu trạng thái là STARTED hoặc SUBMITTED thì không cho thi nữa
      if (status === "STARTED" || status === "SUBMITTED") {
        return true; 
      }
    }
  }
  return false;
}

function resetStudent(examId, key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === examId && data[i][2] === key) {
      // Đổi trạng thái thành RESET để bỏ qua ở lần kiểm tra sau
      sheet.getRange(i + 1, 13).setValue("RESET");
    }
  }
}

function logStudentStatus(examId, key, name, cls, sbd, status, payload = null) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTS);
  if (!sheet) return;
  
  const timestamp = new Date();
  
  if (status === "STARTED") {
    sheet.appendRow([
      timestamp, examId, key, name, cls, sbd, 
      "", "", "", "", "", "", "STARTED"
    ]);
  } 
  else if (status === "SUBMITTED" && payload) {
    sheet.appendRow([
      timestamp, 
      examId, 
      key, 
      name, 
      cls, 
      sbd, 
      payload.score, 
      payload.correct, 
      payload.time, 
      payload.tabSwitches, 
      payload.submittedAt, 
      payload.detail, 
      "SUBMITTED"
    ]);
  }
}
