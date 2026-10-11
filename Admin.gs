/**
 * Admin.gs — Module quản trị cho bản Google Apps Script. Có thể thêm/gỡ độc lập.
 *
 * Lưu trữ (tự tạo ở lần dùng đầu tiên, trong 1 thư mục Drive):
 *   - Google Sheet "Quy ước rút gọn": tab KhachHang (4 cột), tab SanPham (2 cột)
 *   - File form báo cáo .xlsx (chỉ giữ 1 form hiện hành)
 *
 * Cài đặt mật khẩu admin: Cài đặt dự án -> Thuộc tính tập lệnh -> thêm ADMIN_PASSWORD.
 *
 * Hàm công khai (mọi người dùng): getRules, getForm
 * Hàm admin (cần token từ adminLogin): adminCheck, adminGetRules, adminSaveRules,
 *   adminAppendNew, adminFormInfo, adminUploadForm, adminLinks
 */
var ADM = {
  PW: 'ADMIN_PASSWORD', FOLDER: 'FOLDER_ID', SHEET: 'RULES_SHEET_ID', FORM: 'FORM_FILE_ID',
  TAB_KH: 'KhachHang', TAB_SP: 'SanPham',
  HEAD_KH: ['Tên khách hàng (gốc)', 'Điểm giao (gốc)', 'Khách hàng', 'Điểm giao'],
  HEAD_SP: ['Sản phẩm (gốc)', 'Sản phẩm rút gọn'],
  XLSX: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

/* ---------- tiện ích ---------- */
function props_() { return PropertiesService.getScriptProperties(); }
function norm_(s) { return String(s == null ? '' : s).normalize('NFC').replace(/\s+/g, ' ').trim(); }
function key_(s) { return norm_(s).toLowerCase(); }
function kind_(k) {
  if (k !== 'kh' && k !== 'sp') throw new Error('Loại quy ước không hợp lệ.');
  return k === 'kh'
    ? { tab: ADM.TAB_KH, n: 4, head: ADM.HEAD_KH, keyCols: [0, 1] }
    : { tab: ADM.TAB_SP, n: 2, head: ADM.HEAD_SP, keyCols: [0] };
}
function rowKey_(k, row) { return k === 'kh' ? key_(row[0]) + '||' + key_(row[1]) : key_(row[0]); }
function ver_(k) { return Number(props_().getProperty('VER_' + k) || 0); }
function bump_(k) { const v = ver_(k) + 1; props_().setProperty('VER_' + k, String(v)); return v; }
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- thư mục, Sheet, form ---------- */
function getFolder_() {
  const p = props_(); const id = p.getProperty(ADM.FOLDER);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* tạo lại */ } }
  const f = DriveApp.createFolder('Tra cứu Ship - quy ước và form');
  p.setProperty(ADM.FOLDER, f.getId());
  return f;
}
function tab_(ss, k) {
  const c = kind_(k);
  let sh = ss.getSheetByName(c.tab);
  if (!sh) {
    sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(c.tab) : ss.insertSheet(c.tab);
    sh.getRange(1, 1, 1, c.n).setValues([c.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), c.n).setNumberFormat('@');   // giữ nguyên dạng chữ (vd. "1:" không bị đổi thành giờ)
  }
  return sh;
}
function getSS_() {
  const p = props_(); const id = p.getProperty(ADM.SHEET);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* tạo lại */ } }
  const ss = SpreadsheetApp.create('Quy ước rút gọn');
  DriveApp.getFileById(ss.getId()).moveTo(getFolder_());
  p.setProperty(ADM.SHEET, ss.getId());
  tab_(ss, 'kh'); tab_(ss, 'sp');
  return ss;
}
function readRows_(k) {
  const c = kind_(k), sh = tab_(getSS_(), k), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, c.n).getDisplayValues()
    .map(function (r) { return r.map(norm_); })
    .filter(function (r) { return r.some(function (x) { return x !== ''; }); });
}
function writeAll_(k, rows) {
  const c = kind_(k), sh = tab_(getSS_(), k), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, c.n).clearContent();
  if (rows.length) { const rg = sh.getRange(2, 1, rows.length, c.n); rg.setNumberFormat('@'); rg.setValues(rows); }
}

/* ---------- công khai ---------- */
function getRules() {
  return {
    kh: readRows_('kh').filter(function (r) { return r[2] || r[3]; }),
    sp: readRows_('sp').filter(function (r) { return r[1]; })
  };
}
function formInfo_(withData) {
  const id = props_().getProperty(ADM.FORM);
  if (!id) return null;
  try {
    const f = DriveApp.getFileById(id);
    if (f.isTrashed()) return null;
    const bytes = f.getBlob().getBytes();
    const o = { name: f.getName(), size: bytes.length, updated: f.getLastUpdated().getTime() };
    if (withData) o.b64 = Utilities.base64Encode(bytes);
    return o;
  } catch (e) { return null; }
}
function getForm() { return formInfo_(true); }

/* ---------- xác thực admin (mật khẩu) ---------- */
function adminLogin(pw) {
  const real = props_().getProperty(ADM.PW);
  if (!real) throw new Error('Chưa đặt mật khẩu admin. Vào Cài đặt dự án > Thuộc tính tập lệnh, thêm ADMIN_PASSWORD.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adm_fail') || 0);
  if (fails >= 8) throw new Error('Nhập sai quá nhiều lần. Thử lại sau 10 phút.');
  if (String(pw) !== real) { cache.put('adm_fail', String(fails + 1), 600); throw new Error('Sai mật khẩu.'); }
  cache.remove('adm_fail');
  const token = Utilities.getUuid();
  cache.put('adm_' + token, '1', 21600);   // phiên admin 6 giờ
  return token;
}
function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('adm_' + token) !== '1') {
    throw new Error('Phiên admin đã hết hạn. Hãy đăng nhập lại.');
  }
}
function adminCheck(token) {
  try { requireAdmin_(token); return true; } catch (e) { return false; }
}

/* ---------- admin: quy ước ---------- */
function adminGetRules(token) {
  requireAdmin_(token);
  return { kh: readRows_('kh'), sp: readRows_('sp'), ver: { kh: ver_('kh'), sp: ver_('sp') } };
}

// Lưu toàn bộ bảng (thay thế). baseVer để phát hiện người khác đã sửa trước.
function adminSaveRules(token, k, rows, baseVer) {
  requireAdmin_(token);
  const c = kind_(k);
  return withLock_(function () {
    if (Number(baseVer) !== ver_(k)) throw new Error('Dữ liệu đã được thay đổi ở nơi khác. Bấm "Tải lại" rồi sửa lại.');
    const clean = [];
    (rows || []).forEach(function (r) {
      const row = []; for (let i = 0; i < c.n; i++) row.push(norm_(r && r[i]));
      if (c.keyCols.some(function (i) { return row[i]; })) clean.push(row);
    });
    writeAll_(k, clean);
    return { ver: bump_(k), count: clean.length };
  });
}

// Nút "Cập nhật": thêm các mục CHƯA có vào cuối danh sách (cột rút gọn để trống).
// items: kh -> [[khách, điểm giao], ...]; sp -> [tên sản phẩm, ...]
function adminAppendNew(token, k, items) {
  requireAdmin_(token);
  const c = kind_(k);
  return withLock_(function () {
    const have = {};
    readRows_(k).forEach(function (r) { have[rowKey_(k, r)] = true; });
    const add = [];
    (items || []).forEach(function (it) {
      const row = k === 'kh' ? [norm_(it && it[0]), norm_(it && it[1]), '', ''] : [norm_(Array.isArray(it) ? it[0] : it), ''];
      if (!c.keyCols.some(function (i) { return row[i]; })) return;
      const kk = rowKey_(k, row);
      if (have[kk]) return;
      have[kk] = true; add.push(row);
    });
    if (add.length) {
      const sh = tab_(getSS_(), k), at = Math.max(sh.getLastRow(), 1) + 1;
      const rg = sh.getRange(at, 1, add.length, c.n);
      rg.setNumberFormat('@'); rg.setValues(add);
      bump_(k);
    }
    return { added: add.length, total: readRows_(k).length };
  });
}

/* ---------- admin: form báo cáo ---------- */
function adminFormInfo(token) { requireAdmin_(token); return formInfo_(false); }

function adminUploadForm(token, b64, fileName) {
  requireAdmin_(token);
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4B) throw new Error('File không phải .xlsx hợp lệ.');
  const name = norm_(fileName).replace(/[\\/:*?"<>|]/g, '') || 'Form báo cáo.xlsx';
  return withLock_(function () {
    const old = props_().getProperty(ADM.FORM);
    const file = getFolder_().createFile(Utilities.newBlob(bytes, ADM.XLSX, name));
    props_().setProperty(ADM.FORM, file.getId());
    if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) { /* bỏ qua */ } }
    return formInfo_(false);
  });
}

function adminLinks(token) {
  requireAdmin_(token);
  return { sheet: getSS_().getUrl(), folder: getFolder_().getUrl() };
}
