/**
 * ============================================================================
 * ผลตรวจสุขภาพเบื้องต้น (Health Check Records) — Google Apps Script Backend
 * ============================================================================
 *
 * โครงการ: ระบบบันทึก/ประมวลผล/แสดงผลข้อมูลผลตรวจสุขภาพเบื้องต้น
 * ไฟล์นี้เป็น backend API (Web App) ที่เชื่อมต่อกับ Google Sheets
 * ทำงานคู่กับ index.html (Standalone Web App)
 *
 * ----------------------------------------------------------------------------
 * วิธีติดตั้งและ Deploy (ทำตามลำดับ)
 * ----------------------------------------------------------------------------
 * 1) สร้าง Google Sheet ใหม่ (หรือใช้ไฟล์ที่แปลงจาก Excel Template ที่แนบมา)
 *    - ต้องมีชีตชื่อ "Records" (ระบบจะสร้าง Header แถวแรกให้อัตโนมัติถ้ายังไม่มี)
 *    - ต้องมีชีตชื่อ "AuditLog" (ระบบจะสร้างให้อัตโนมัติถ้ายังไม่มี)
 *    - ต้องมีชีตชื่อ "Settings" (ระบบจะสร้างให้อัตโนมัติถ้ายังไม่มี พร้อมบัญชี admin/admin เริ่มต้น)
 *
 * 2) เปิด Google Sheet -> เมนู Extensions/ส่วนขยาย -> Apps Script
 *    - ลบโค้ดตัวอย่างเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ทั้งหมดแทน
 *
 * 3) ตั้งค่า Script Properties (Project Settings -> Script Properties) [ไม่บังคับ]
 *    - SPREADSHEET_ID : ค่าเริ่มต้นว่างเปล่า (สคริปต์จะใช้ SpreadsheetApp.getActiveSpreadsheet() คือสเปรดชีตที่ผูกสคริปต์นี้อยู่)
 *      หากต้องการให้สคริปต์รันจากไฟล์อื่นโดยไม่ผูกกับสเปรดชีต ให้ตั้งค่า Script Property ชื่อ SPREADSHEET_ID เป็น ID ของ Google Sheet ปลายทาง
 *      (คัดลอกจาก URL ของสเปรดชีต ส่วนระหว่าง /d/ กับ /edit) จะ override ค่าว่างในโค้ดทันทีโดยไม่ต้องแก้โค้ด
 *    - RECORDS_SHEET_NAME : ค่าเริ่มต้น "Records"
 *    - AUDIT_SHEET_NAME   : ค่าเริ่มต้น "AuditLog"
 *    - SETTINGS_SHEET_NAME: ค่าเริ่มต้น "Settings"
 *
 * 4) กด Deploy -> New deployment
 *    - เลือกประเภท "Web app"
 *    - Description: ใส่ชื่อเวอร์ชันที่ต้องการ
 *    - Execute as: "Me" (บัญชีของคุณ)
 *    - Who has access: "Anyone" (เพื่อให้หน้าเว็บสาธารณะและฟอร์มบันทึกข้อมูลเรียก API ได้)
 *    - กด Deploy แล้วอนุญาตสิทธิ์ (Authorize access) ตามที่ระบบขอ
 *    - คัดลอก "Web app URL" ที่ได้ (ลงท้ายด้วย /exec)
 *
 * 5) นำ URL ไปวางในไฟล์ index.html ที่ตัวแปร GAS_API_URL (บรรทัดเดียว จุดเดียวในไฟล์)
 *    ระบบจะเปลี่ยนจาก Demo Mode (Local Storage) ไปใช้ Google Sheets ทันที
 *
 * 6) ทุกครั้งที่แก้โค้ดไฟล์นี้ ต้องกด Deploy -> Manage deployments -> แก้ไข (ไอคอนดินสอ)
 *    -> เปลี่ยน Version เป็น "New version" -> Deploy ใหม่ เพื่อให้ URL เดิมใช้โค้ดล่าสุด
 *
 * ----------------------------------------------------------------------------
 * รูปแบบ Action ที่รองรับ
 * ----------------------------------------------------------------------------
 * GET  ?action=ping                              -> ตรวจสอบสถานะ API
 * GET  ?action=list&filters...                    -> ดึงรายการทั้งหมด (แบบ public-safe หรือเต็มถ้ามี token)
 * GET  ?action=get&record_id=REC-xxxx             -> ดึงรายการเดียว
 * GET  ?action=summary&filters...                 -> สรุปผลสำหรับ Dashboard (ไม่มีข้อมูลส่วนบุคคล)
 * POST { action: "create", payload: {...} }                         -> เพิ่มข้อมูลใหม่ (สาธารณะ)
 * POST { action: "update", payload: {...}, auth:{token} }            -> แก้ไข (ต้องมี token ผู้ดูแล)
 * POST { action: "delete", record_id, auth:{token} }                 -> ลบ (soft delete, ต้องมี token)
 * POST { action: "adminLogin", username, password }                  -> เข้าสู่ระบบผู้ดูแล คืน token
 * POST { action: "changeAdminCredential", auth:{token}, newUsername, newPassword } -> เปลี่ยนบัญชีผู้ดูแล
 *
 * ผลลัพธ์ทุก action จะอยู่ในรูปแบบ JSON:
 *   { success: true|false, message: "...", data: {...}|[...], error: null|"..." }
 * ============================================================================
 */

// ============================================================================
// ค่าคงที่ของระบบ (ปรับได้จากจุดเดียว)
// ============================================================================
var SPREADSHEET_ID      = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
var RECORDS_SHEET_NAME  = PropertiesService.getScriptProperties().getProperty('RECORDS_SHEET_NAME') || 'Records';
var AUDIT_SHEET_NAME    = PropertiesService.getScriptProperties().getProperty('AUDIT_SHEET_NAME') || 'AuditLog';
var SETTINGS_SHEET_NAME = PropertiesService.getScriptProperties().getProperty('SETTINGS_SHEET_NAME') || 'Settings';
var TOKEN_TTL_SECONDS   = 8 * 60 * 60; // อายุ token ผู้ดูแล 8 ชั่วโมง

// โครงสร้างคอลัมน์ทั้งหมด (ต้องตรงกับ index.html และ Excel Template ทุกตัวอักษร)
// type: text | number | boolean | select
var FIELD_SCHEMA = [
  // ฟิลด์ระบบ
  { key: 'record_id',       label: 'รหัสรายการ',                     type: 'text',    system: true, required: false },
  { key: 'created_at',      label: 'วันเวลาที่บันทึก',                 type: 'text',    system: true, required: false },
  { key: 'updated_at',      label: 'วันเวลาที่แก้ไขล่าสุด',            type: 'text',    system: true, required: false },
  { key: 'record_status',   label: 'สถานะรายการ',                    type: 'text',    system: true, required: false },
  { key: 'created_by',      label: 'ผู้บันทึก',                       type: 'text',    system: true, required: false },
  { key: 'updated_by',      label: 'ผู้แก้ไขล่าสุด',                   type: 'text',    system: true, required: false },
  // ข้อมูลส่วนบุคคล (จากส่วนหัวของแบบฟอร์มต้นฉบับ)
  { key: 'first_name',      label: 'ชื่อ',                            type: 'text',    required: true },
  { key: 'last_name',       label: 'นามสกุล',                         type: 'text',    required: true },
  { key: 'gender',          label: 'เพศ',                             type: 'select',  required: true, note: 'เพิ่มเติมจากต้นฉบับ: จำเป็นสำหรับแปลผลตามเกณฑ์ที่แยกชาย/หญิงในแบบฟอร์ม' },
  { key: 'age',             label: 'อายุ (ปี)',                       type: 'number',  required: true },
  { key: 'occupation',      label: 'อาชีพ',                           type: 'text',    required: false },
  { key: 'phone',           label: 'เบอร์โทรศัพท์',                    type: 'text',    required: false },
  { key: 'height_cm',       label: 'ส่วนสูง (ซม.)',                    type: 'number',  required: true },
  { key: 'waist_cm',        label: 'รอบเอว (ซม.)',                     type: 'number',  required: true },
  { key: 'disease_diabetes',     label: 'โรคประจำตัว: เบาหวาน',        type: 'boolean', required: false },
  { key: 'disease_hypertension', label: 'โรคประจำตัว: ความดัน',        type: 'boolean', required: false },
  { key: 'disease_lipid',        label: 'โรคประจำตัว: ไขมัน',          type: 'boolean', required: false },
  { key: 'disease_other',        label: 'โรคประจำตัว: อื่นๆ',          type: 'boolean', required: false },
  { key: 'disease_other_detail', label: 'โรคประจำตัวอื่นๆ (ระบุ)',     type: 'text',    required: false, conditionOf: 'disease_other' },
  // ผลตรวจองค์ประกอบร่างกาย (BIA)
  { key: 'weight_kg',              label: 'น้ำหนัก (กก.)',                          type: 'number', required: true },
  { key: 'body_fat_percent',       label: 'ไขมันทั่วร่างกาย (%)',                    type: 'number', required: true },
  { key: 'muscle_mass_kg',         label: 'มวลกล้ามเนื้อ (กก.)',                     type: 'number', required: true },
  { key: 'bmi',                    label: 'ค่าดัชนีมวลกาย (BMI)',                    type: 'number', required: false, note: 'คำนวณอัตโนมัติจากน้ำหนัก/ส่วนสูง แก้ไขได้' },
  { key: 'bmr',                    label: 'อัตราการเผาผลาญพลังงานต่อวัน (BMR)',       type: 'number', required: false },
  { key: 'body_age',               label: 'อายุการเผาผลาญของร่างกาย (ปี)',           type: 'number', required: false },
  { key: 'body_water_percent',     label: 'น้ำในร่างกาย (%)',                        type: 'number', required: false },
  { key: 'visceral_fat_level',     label: 'ไขมันในช่องท้อง (ระดับ)',                 type: 'number', required: false },
  // ผลตรวจความเครียด
  { key: 'physical_stress', label: 'ความเครียดทางกาย (0-100)', type: 'number', required: false },
  { key: 'mental_stress',   label: 'ความเครียดทางจิตใจ (0-100)', type: 'number', required: false },
  { key: 'stress_level',    label: 'ระดับความเครียด (0-100)',   type: 'number', required: false },
  // ผลตรวจความยืดหยุ่นของหลอดเลือด
  { key: 'peripheral_artery_elasticity', label: 'ความยืดหยุ่นของหลอดเลือดแดงส่วนปลาย', type: 'number', required: false },
  { key: 'artery_elasticity',            label: 'ความยืดหยุ่นของหลอดเลือดแดง',         type: 'number', required: false },
  // ผลตรวจสุขภาพหลอดเลือด (Type 1-7, หน่วยร้อยละ รวมกันไม่เกิน 100)
  { key: 'vascular_type1_percent', label: 'หลอดเลือด Type1 (%)', type: 'number', required: false },
  { key: 'vascular_type2_percent', label: 'หลอดเลือด Type2 (%)', type: 'number', required: false },
  { key: 'vascular_type3_percent', label: 'หลอดเลือด Type3 (%)', type: 'number', required: false },
  { key: 'vascular_type4_percent', label: 'หลอดเลือด Type4 (%)', type: 'number', required: false },
  { key: 'vascular_type5_percent', label: 'หลอดเลือด Type5 (%)', type: 'number', required: false },
  { key: 'vascular_type6_percent', label: 'หลอดเลือด Type6 (%)', type: 'number', required: false },
  { key: 'vascular_type7_percent', label: 'หลอดเลือด Type7 (%)', type: 'number', required: false }
];

var FIELD_KEYS = FIELD_SCHEMA.map(function (f) { return f.key; });
var PII_FIELDS = ['first_name', 'last_name', 'phone', 'occupation']; // ฟิลด์ที่ห้ามแสดงใน public dashboard

// ============================================================================
// Entry points: doGet / doPost
// ============================================================================
function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || 'ping';

    switch (action) {
      case 'ping':
        return jsonOut({ success: true, message: 'API is online', data: { time: new Date().toISOString() }, error: null });
      case 'list':
        return jsonOut(handleList(params));
      case 'get':
        return jsonOut(handleGet(params));
      case 'summary':
        return jsonOut(handleSummary(params));
      default:
        return jsonOut({ success: false, message: 'ไม่รู้จัก action: ' + action, data: null, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    return jsonOut(errorResponse(err));
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action;

    switch (action) {
      case 'create':
        return jsonOut(handleCreate(body.payload || {}));
      case 'update':
        return jsonOut(handleUpdate(body.payload || {}, body.auth || {}));
      case 'delete':
        return jsonOut(handleDelete(body.record_id, body.auth || {}));
      case 'adminLogin':
        return jsonOut(handleAdminLogin(body.username, body.password));
      case 'changeAdminCredential':
        return jsonOut(handleChangeAdminCredential(body.auth || {}, body.newUsername, body.newPassword));
      default:
        return jsonOut({ success: false, message: 'ไม่รู้จัก action: ' + action, data: null, error: 'UNKNOWN_ACTION' });
    }
  } catch (err) {
    return jsonOut(errorResponse(err));
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(err) {
  logError(err);
  return { success: false, message: 'เกิดข้อผิดพลาดในระบบ', data: null, error: String(err && err.message || err) };
}

// ============================================================================
// Sheet helpers
// ============================================================================
function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headerRow) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0 && headerRow && headerRow.length) {
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.setFrozenRows(1);
  } else if (headerRow && headerRow.length) {
    // ตรวจสอบและเรียง Header ให้ตรงกับโครงสร้างข้อมูลล่าสุด (เพิ่มคอลัมน์ที่ขาดต่อท้าย)
    var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var existingSet = {};
    existing.forEach(function (h) { if (h) existingSet[h] = true; });
    var missing = headerRow.filter(function (h) { return !existingSet[h]; });
    if (missing.length) {
      sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sheet;
}

function getRecordsSheet() {
  return getOrCreateSheet(RECORDS_SHEET_NAME, FIELD_KEYS);
}

function getAuditSheet() {
  return getOrCreateSheet(AUDIT_SHEET_NAME, ['timestamp', 'actor', 'action', 'record_id', 'detail']);
}

var ADMIN_HASH_PATTERN = /^[0-9a-f]{64}$/i; // รูปแบบ SHA-256 hex ที่ถูกต้อง

function getSettingsSheet() {
  var sheet = getOrCreateSheet(SETTINGS_SHEET_NAME, ['key', 'value']);
  ensureValidAdminSeed(sheet);
  return sheet;
}

// ตรวจสอบและซ่อมแซมบัญชี admin เริ่มต้นให้ใช้งานได้เสมอ ไม่ใช่แค่เช็คจำนวนแถว
// (กรณีนำเข้าจาก Excel Template ที่มีข้อความ placeholder อยู่ในคอลัมน์ value แทนที่จะเป็น hash จริง
// การเช็คแค่ sheet.getLastRow() < 2 จะพลาดไม่สร้าง hash ที่ถูกต้องให้ ทำให้ล็อกอิน admin/admin ไม่ได้เลย)
function ensureValidAdminSeed(sheet) {
  var values = sheet.getDataRange().getValues();
  var usernameRow = -1, hashRow = -1, usernameVal = '', hashVal = '';
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === 'admin_username') { usernameRow = i + 1; usernameVal = values[i][1]; }
    if (values[i][0] === 'admin_password_hash') { hashRow = i + 1; hashVal = values[i][1]; }
  }
  var needsUsername = !usernameVal;
  var needsHash = !hashVal || !ADMIN_HASH_PATTERN.test(String(hashVal));
  if (usernameRow === -1 || needsUsername) {
    if (usernameRow === -1) { sheet.appendRow(['admin_username', 'admin']); }
    else { sheet.getRange(usernameRow, 2).setValue('admin'); }
  }
  if (hashRow === -1 || needsHash) {
    var defaultHash = hashPassword('admin');
    if (hashRow === -1) { sheet.appendRow(['admin_password_hash', defaultHash]); }
    else { sheet.getRange(hashRow, 2).setValue(defaultHash); }
  }
}

function getSetting(key) {
  var sheet = getSettingsSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return null;
}

function setSetting(key, value) {
  var sheet = getSettingsSheet();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getHeaderMap(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  header.forEach(function (h, idx) { if (h) map[h] = idx; });
  return { header: header, map: map };
}

function rowToRecord(row, headerMap) {
  var rec = {};
  FIELD_KEYS.forEach(function (key) {
    var idx = headerMap.map[key];
    rec[key] = idx === undefined ? '' : row[idx];
  });
  return rec;
}

function stripPII(rec) {
  var clone = {};
  for (var k in rec) {
    if (PII_FIELDS.indexOf(k) === -1) clone[k] = rec[k];
  }
  return clone;
}

// ============================================================================
// Validation & utilities
// ============================================================================
function sanitizeString(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .trim();
}

function generateRecordId() {
  var now = new Date();
  var stamp = Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT+7', 'yyMMddHHmmss');
  var rand = Math.floor(1000 + Math.random() * 9000);
  return 'REC-' + stamp + '-' + rand;
}

function nowIso() {
  return new Date().toISOString();
}

function hashPassword(pw) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw), Utilities.Charset.UTF_8);
  return digest.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function generateToken() {
  return Utilities.getUuid() + '-' + new Date().getTime();
}

function logError(err) {
  try {
    var sheet = getAuditSheet();
    sheet.appendRow([nowIso(), 'system', 'error', '', String(err && err.stack || err)]);
  } catch (e) { /* ignore logging failure */ }
}

function logAudit(actor, action, recordId, detail) {
  try {
    var sheet = getAuditSheet();
    sheet.appendRow([nowIso(), actor || 'unknown', action, recordId || '', detail || '']);
  } catch (e) { /* ignore */ }
}

function validatePayload(payload, isUpdate) {
  var errors = [];
  FIELD_SCHEMA.forEach(function (field) {
    if (field.system) return;
    var val = payload[field.key];
    if (field.required && !isUpdate && (val === undefined || val === null || val === '')) {
      errors.push(field.label + ' (' + field.key + ') จำเป็นต้องกรอก');
    }
    if (field.type === 'number' && val !== undefined && val !== null && val !== '' && isNaN(Number(val))) {
      errors.push(field.label + ' (' + field.key + ') ต้องเป็นตัวเลข');
    }
  });
  if (payload.gender && ['ชาย', 'หญิง'].indexOf(payload.gender) === -1) {
    errors.push('เพศ ต้องเป็น "ชาย" หรือ "หญิง"');
  }
  return errors;
}

// ============================================================================
// Auth (token ผู้ดูแล เก็บใน CacheService)
// ============================================================================
function requireAdmin(auth) {
  var token = auth && auth.token;
  if (!token) return { ok: false, message: 'ไม่ได้เข้าสู่ระบบผู้ดูแล (missing token)' };
  var cache = CacheService.getScriptCache();
  var username = cache.get('admin_token_' + token);
  if (!username) return { ok: false, message: 'Token หมดอายุหรือไม่ถูกต้อง กรุณาเข้าสู่ระบบใหม่' };
  return { ok: true, username: username };
}

function handleAdminLogin(username, password) {
  if (!username || !password) {
    return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน', data: null, error: 'MISSING_CREDENTIAL' };
  }
  var storedUser = getSetting('admin_username');
  var storedHash = getSetting('admin_password_hash');
  if (username !== storedUser || hashPassword(password) !== storedHash) {
    logAudit(username, 'login_failed', '', '');
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', data: null, error: 'INVALID_CREDENTIAL' };
  }
  var token = generateToken();
  CacheService.getScriptCache().put('admin_token_' + token, username, TOKEN_TTL_SECONDS);
  logAudit(username, 'login_success', '', '');
  return { success: true, message: 'เข้าสู่ระบบสำเร็จ', data: { token: token, username: username, expiresInSeconds: TOKEN_TTL_SECONDS }, error: null };
}

function handleChangeAdminCredential(auth, newUsername, newPassword) {
  var check = requireAdmin(auth);
  if (!check.ok) return { success: false, message: check.message, data: null, error: 'UNAUTHORIZED' };
  if (!newUsername || !newPassword || String(newPassword).length < 4) {
    return { success: false, message: 'ชื่อผู้ใช้ห้ามว่าง และรหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร', data: null, error: 'INVALID_INPUT' };
  }
  setSetting('admin_username', sanitizeString(newUsername));
  setSetting('admin_password_hash', hashPassword(newPassword));
  logAudit(check.username, 'change_credential', '', 'username -> ' + newUsername);
  return { success: true, message: 'เปลี่ยนบัญชีผู้ดูแลสำเร็จ กรุณาเข้าสู่ระบบใหม่', data: null, error: null };
}

// ============================================================================
// CRUD handlers
// ============================================================================
function handleCreate(payload) {
  var errors = validatePayload(payload, false);
  if (errors.length) {
    return { success: false, message: 'ข้อมูลไม่ถูกต้อง', data: { errors: errors }, error: 'VALIDATION_ERROR' };
  }

  var sheet = getRecordsSheet();
  var headerInfo = getHeaderMap(sheet);

  // ป้องกันข้อมูลซ้ำด้วย record_id ที่ส่งมา (idempotent create) ถ้ามีอยู่แล้วให้ปฏิเสธ
  if (payload.record_id) {
    var existing = findRowByRecordId(sheet, headerInfo, payload.record_id);
    if (existing) {
      return { success: false, message: 'record_id นี้มีอยู่แล้วในระบบ', data: null, error: 'DUPLICATE_RECORD_ID' };
    }
  }

  var recordId = payload.record_id || generateRecordId();
  var now = nowIso();
  var record = { record_id: recordId, created_at: now, updated_at: now, record_status: 'active', created_by: sanitizeString(payload.created_by) || 'public_form', updated_by: '' };

  FIELD_SCHEMA.forEach(function (field) {
    if (field.system) return;
    var v = payload[field.key];
    if (field.type === 'text') v = sanitizeString(v);
    if (field.type === 'boolean') v = (v === true || v === 'true' || v === 1 || v === '1');
    if (field.type === 'number') v = (v === '' || v === undefined || v === null) ? '' : Number(v);
    record[field.key] = v === undefined ? '' : v;
  });

  // คำนวณ BMI อัตโนมัติถ้าไม่ได้ส่งมาหรือส่งมาว่าง
  if (!record.bmi && record.height_cm && record.weight_kg) {
    var hM = Number(record.height_cm) / 100;
    record.bmi = Math.round((Number(record.weight_kg) / (hM * hM)) * 100) / 100;
  }

  var row = FIELD_KEYS.map(function (k) { return record[k] === undefined ? '' : record[k]; });
  sheet.appendRow(row);
  logAudit(record.created_by, 'create', recordId, '');

  return { success: true, message: 'บันทึกข้อมูลสำเร็จ', data: record, error: null };
}

function findRowByRecordId(sheet, headerInfo, recordId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var idCol = headerInfo.map['record_id'];
  if (idCol === undefined) return null;
  var ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === recordId) return i + 2; // แถวจริงใน sheet (1-indexed, +header)
  }
  return null;
}

function handleUpdate(payload, auth) {
  var check = requireAdmin(auth);
  if (!check.ok) return { success: false, message: check.message, data: null, error: 'UNAUTHORIZED' };
  if (!payload.record_id) return { success: false, message: 'ต้องระบุ record_id', data: null, error: 'MISSING_ID' };

  var errors = validatePayload(payload, true);
  if (errors.length) {
    return { success: false, message: 'ข้อมูลไม่ถูกต้อง', data: { errors: errors }, error: 'VALIDATION_ERROR' };
  }

  var sheet = getRecordsSheet();
  var headerInfo = getHeaderMap(sheet);
  var rowNum = findRowByRecordId(sheet, headerInfo, payload.record_id);
  if (!rowNum) return { success: false, message: 'ไม่พบรายการที่ต้องการแก้ไข', data: null, error: 'NOT_FOUND' };

  var currentRow = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  var current = rowToRecord(currentRow, headerInfo);

  FIELD_SCHEMA.forEach(function (field) {
    if (field.system) return;
    if (payload[field.key] === undefined) return;
    var v = payload[field.key];
    if (field.type === 'text') v = sanitizeString(v);
    if (field.type === 'boolean') v = (v === true || v === 'true' || v === 1 || v === '1');
    if (field.type === 'number') v = (v === '' ? '' : Number(v));
    current[field.key] = v;
  });
  current.updated_at = nowIso();
  current.updated_by = check.username;

  var row = FIELD_KEYS.map(function (k) { return current[k] === undefined ? '' : current[k]; });
  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  logAudit(check.username, 'update', payload.record_id, '');

  return { success: true, message: 'แก้ไขข้อมูลสำเร็จ', data: current, error: null };
}

function handleDelete(recordId, auth) {
  var check = requireAdmin(auth);
  if (!check.ok) return { success: false, message: check.message, data: null, error: 'UNAUTHORIZED' };
  if (!recordId) return { success: false, message: 'ต้องระบุ record_id', data: null, error: 'MISSING_ID' };

  var sheet = getRecordsSheet();
  var headerInfo = getHeaderMap(sheet);
  var rowNum = findRowByRecordId(sheet, headerInfo, recordId);
  if (!rowNum) return { success: false, message: 'ไม่พบรายการที่ต้องการลบ', data: null, error: 'NOT_FOUND' };

  // Soft delete: เปลี่ยนสถานะแทนการลบแถวจริง เพื่อรักษาความสมบูรณ์ของข้อมูลย้อนหลัง
  var statusCol = headerInfo.map['record_status'];
  var updatedAtCol = headerInfo.map['updated_at'];
  var updatedByCol = headerInfo.map['updated_by'];
  if (statusCol !== undefined) sheet.getRange(rowNum, statusCol + 1).setValue('deleted');
  if (updatedAtCol !== undefined) sheet.getRange(rowNum, updatedAtCol + 1).setValue(nowIso());
  if (updatedByCol !== undefined) sheet.getRange(rowNum, updatedByCol + 1).setValue(check.username);

  logAudit(check.username, 'delete', recordId, '');
  return { success: true, message: 'ลบข้อมูลสำเร็จ', data: { record_id: recordId }, error: null };
}

function handleGet(params) {
  var sheet = getRecordsSheet();
  var headerInfo = getHeaderMap(sheet);
  var rowNum = findRowByRecordId(sheet, headerInfo, params.record_id);
  if (!rowNum) return { success: false, message: 'ไม่พบรายการ', data: null, error: 'NOT_FOUND' };
  var row = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  var record = rowToRecord(row, headerInfo);
  var isAdmin = params.token && requireAdmin({ token: params.token }).ok;
  return { success: true, message: 'ok', data: isAdmin ? record : stripPII(record), error: null };
}

function handleList(params) {
  var sheet = getRecordsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'ok', data: [], error: null };
  var headerInfo = getHeaderMap(sheet);
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var records = rows.map(function (r) { return rowToRecord(r, headerInfo); });

  records = applyFilters(records, params);

  var isAdmin = params.token && requireAdmin({ token: params.token }).ok;
  if (!isAdmin) {
    records = records.filter(function (r) { return r.record_status !== 'deleted'; });
    records = records.map(stripPII);
  }

  // เรียงล่าสุดก่อน
  records.sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });

  return { success: true, message: 'ok', data: records, error: null };
}

function applyFilters(records, params) {
  var out = records;
  if (params.gender) out = out.filter(function (r) { return r.gender === params.gender; });
  if (params.date_from) out = out.filter(function (r) { return String(r.created_at) >= String(params.date_from); });
  if (params.date_to) out = out.filter(function (r) { return String(r.created_at) <= String(params.date_to) + 'T23:59:59'; });
  if (params.status) out = out.filter(function (r) { return r.record_status === params.status; });
  if (params.q) {
    var q = String(params.q).toLowerCase();
    out = out.filter(function (r) {
      return (String(r.first_name) + String(r.last_name) + String(r.record_id) + String(r.phone)).toLowerCase().indexOf(q) !== -1;
    });
  }
  return out;
}

// ============================================================================
// Summary handler (สำหรับ Dashboard สาธารณะ — ไม่มีข้อมูลส่วนบุคคล)
// ============================================================================
function handleSummary(params) {
  var sheet = getRecordsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, message: 'ok', data: emptySummary(), error: null };
  var headerInfo = getHeaderMap(sheet);
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var records = rows.map(function (r) { return rowToRecord(r, headerInfo); })
    .filter(function (r) { return r.record_status !== 'deleted'; });

  records = applyFilters(records, params);

  return { success: true, message: 'ok', data: buildSummary(records), error: null };
}

function emptySummary() {
  return { total: 0, genderCount: {}, bmiCategory: {}, waistRisk: {}, avg: {}, vascularTypeDominant: {}, stressBuckets: {}, lastUpdated: nowIso() };
}

function average(arr) {
  var nums = arr.map(Number).filter(function (n) { return !isNaN(n) && n !== 0 || n === 0; }).filter(function (n) { return arr.length && !isNaN(n); });
  var valid = arr.map(Number).filter(function (n) { return !isNaN(n); });
  if (!valid.length) return null;
  return Math.round((valid.reduce(function (a, b) { return a + b; }, 0) / valid.length) * 100) / 100;
}

function buildSummary(records) {
  var total = records.length;
  var genderCount = { 'ชาย': 0, 'หญิง': 0 };
  var bmiCategory = { 'น้ำหนักต่ำกว่าเกณฑ์': 0, 'ปกติ': 0, 'น้ำหนักเกิน': 0, 'โรคอ้วน': 0, 'โรคอ้วนระดับรุนแรง': 0 };
  var waistRisk = { 'ปกติ': 0, 'เกินเกณฑ์มาตรฐานเอเชีย': 0 };
  var visceralRisk = { 'ปกติ': 0, 'เริ่มมีความเสี่ยง': 0, 'มีความเสี่ยงสูง': 0 };
  var vascularTypeDominant = { 'Type1': 0, 'Type2': 0, 'Type3': 0, 'Type4': 0, 'Type5': 0, 'Type6': 0, 'Type7': 0 };
  var stressBuckets = { 'ไม่ดี': 0, 'ปกติ': 0, 'ดี': 0 };
  var diseaseCount = { diabetes: 0, hypertension: 0, lipid: 0, other: 0 };

  records.forEach(function (r) {
    if (genderCount[r.gender] !== undefined) genderCount[r.gender]++;
    var bmiCat = classifyBMI(Number(r.bmi));
    if (bmiCat && bmiCategory[bmiCat] !== undefined) bmiCategory[bmiCat]++;
    var wRisk = classifyWaist(Number(r.waist_cm), r.gender);
    if (wRisk && waistRisk[wRisk] !== undefined) waistRisk[wRisk]++;
    var vRisk = classifyVisceralFat(Number(r.visceral_fat_level));
    if (vRisk && visceralRisk[vRisk] !== undefined) visceralRisk[vRisk]++;
    var dom = dominantVascularType(r);
    if (dom && vascularTypeDominant[dom] !== undefined) vascularTypeDominant[dom]++;
    var sBucket = classifyStressGauge(Number(r.stress_level));
    if (sBucket && stressBuckets[sBucket] !== undefined) stressBuckets[sBucket]++;
    if (r.disease_diabetes === true || r.disease_diabetes === 'TRUE') diseaseCount.diabetes++;
    if (r.disease_hypertension === true || r.disease_hypertension === 'TRUE') diseaseCount.hypertension++;
    if (r.disease_lipid === true || r.disease_lipid === 'TRUE') diseaseCount.lipid++;
    if (r.disease_other === true || r.disease_other === 'TRUE') diseaseCount.other++;
  });

  var avg = {
    age: average(records.map(function (r) { return r.age; })),
    bmi: average(records.map(function (r) { return r.bmi; })),
    weight_kg: average(records.map(function (r) { return r.weight_kg; })),
    body_fat_percent: average(records.map(function (r) { return r.body_fat_percent; })),
    muscle_mass_kg: average(records.map(function (r) { return r.muscle_mass_kg; })),
    body_water_percent: average(records.map(function (r) { return r.body_water_percent; })),
    visceral_fat_level: average(records.map(function (r) { return r.visceral_fat_level; })),
    bmr: average(records.map(function (r) { return r.bmr; })),
    waist_cm: average(records.map(function (r) { return r.waist_cm; })),
    physical_stress: average(records.map(function (r) { return r.physical_stress; })),
    mental_stress: average(records.map(function (r) { return r.mental_stress; })),
    stress_level: average(records.map(function (r) { return r.stress_level; }))
  };

  return {
    total: total,
    genderCount: genderCount,
    bmiCategory: bmiCategory,
    waistRisk: waistRisk,
    visceralRisk: visceralRisk,
    vascularTypeDominant: vascularTypeDominant,
    stressBuckets: stressBuckets,
    diseaseCount: diseaseCount,
    avg: avg,
    lastUpdated: nowIso()
  };
}

// ============================================================================
// เกณฑ์การแปลผล (ต้องตรงกับที่ระบุในแบบฟอร์มต้นฉบับ และตรงกับ index.html)
// ============================================================================
function classifyBMI(bmi) {
  if (!bmi || isNaN(bmi)) return null;
  if (bmi < 18.5) return 'น้ำหนักต่ำกว่าเกณฑ์';
  if (bmi < 23) return 'ปกติ';
  if (bmi < 25) return 'น้ำหนักเกิน';
  if (bmi < 30) return 'โรคอ้วน';
  return 'โรคอ้วนระดับรุนแรง';
}

function classifyWaist(waist, gender) {
  if (!waist || isNaN(waist) || !gender) return null;
  var limit = gender === 'ชาย' ? 90 : 80;
  return waist <= limit ? 'ปกติ' : 'เกินเกณฑ์มาตรฐานเอเชีย';
}

function classifyVisceralFat(level) {
  if (!level || isNaN(level)) return null;
  if (level <= 9) return 'ปกติ';
  if (level <= 14) return 'เริ่มมีความเสี่ยง';
  return 'มีความเสี่ยงสูง';
}

function classifyStressGauge(value) {
  if (value === '' || value === null || value === undefined || isNaN(value)) return null;
  if (value < 34) return 'ไม่ดี';
  if (value < 67) return 'ปกติ';
  return 'ดี';
}

function dominantVascularType(r) {
  var types = ['vascular_type1_percent', 'vascular_type2_percent', 'vascular_type3_percent', 'vascular_type4_percent', 'vascular_type5_percent', 'vascular_type6_percent', 'vascular_type7_percent'];
  var max = -1, idx = -1;
  types.forEach(function (t, i) {
    var v = Number(r[t]);
    if (!isNaN(v) && v > max) { max = v; idx = i; }
  });
  if (idx === -1 || max <= 0) return null;
  return 'Type' + (idx + 1);
}
