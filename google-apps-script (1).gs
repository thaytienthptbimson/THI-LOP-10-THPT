
const ADMIN_PIN = '28011981';   // đổi thành mã PIN quản trị của bạn (khớp với PIN đăng nhập)
const SHEET_NAME = 'KetQua';    // bảng điểm
const USED_SHEET = 'DaThi';     // danh sách thí sinh đã bắt đầu thi (chống thi lại)

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'getConfig') {
    const raw = PropertiesService.getScriptProperties().getProperty('schedule' + suffix(p));
    return out({ ok: true, now: Date.now(), schedule: raw ? JSON.parse(raw) : null });
  }

  if (p.action === 'start' || p.action === 'check') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = usedSheet();
      const keys = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat() : [];
      const used = keys.indexOf(ns(p)) >= 0;
      if (p.action === 'check') return out({ ok: true, used });
      if (used) return out({ ok: true, allowed: false });
      sh.appendRow([ns(p), p.name, p.cls, p.sbd, new Date()]);
      return out({ ok: true, allowed: true });
    } finally {
      lock.releaseLock();
    }
  }

  return out({ ok: true, now: Date.now() });
}

function doPost(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'setConfig') {
    if (p.pin !== ADMIN_PIN) return out({ ok: false, error: 'PIN sai' });
    const props = PropertiesService.getScriptProperties();
    if (p.schedule) {
      try { JSON.parse(p.schedule); } catch (err) { return out({ ok: false, error: 'Lịch không hợp lệ' }); }
      props.setProperty('schedule' + suffix(p), p.schedule);
    } else {
      props.deleteProperty('schedule' + suffix(p));
    }
    return out({ ok: true });
  }

  if (p.action === 'reset') {
    if (p.pin !== ADMIN_PIN) return out({ ok: false, error: 'PIN sai' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = usedSheet();
      if (sh.getLastRow()) {
        const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
        for (let i = keys.length - 1; i >= 0; i--) if (keys[i] === ns(p)) sh.deleteRow(i + 1);
      }
    } finally {
      lock.releaseLock();
    }
    return out({ ok: true });
  }

  // Ghi điểm
  const safe = v => (/^[=+\-@]/.test(String(v || '')) ? "'" + v : (v || ''));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nm = SHEET_NAME + suffix(p);
    const sh = ss.getSheetByName(nm) || ss.insertSheet(nm);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Thời điểm nộp', 'Họ và tên', 'Lớp', 'SBD', 'Điểm', 'Số câu đúng', 'Thời gian (giây)', 'Rời màn hình (lần)', 'Bài làm câu điền/viết lại']);
    }
    sh.appendRow([safe(p.submittedAt), safe(p.name), safe(p.cls), safe(p.sbd), p.score, p.correct, p.time, p.tabSwitches, safe(p.detail)]);
  } finally {
    lock.releaseLock();
  }
  return out({ ok: true });
}

// Mỗi mã đề (p.exam) có lịch thi, danh sách đã thi và bảng điểm riêng.
function suffix(p) { return p.exam ? '_' + p.exam : ''; }
function ns(p) { return p.exam ? p.exam + '|' + p.key : p.key; }

function usedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(USED_SHEET) || ss.insertSheet(USED_SHEET);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
