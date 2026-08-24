/**
 * ============================================================
 *   MIX FIT JO — نظام إدارة النادي الرياضي (Backend) — v3
 *   Google Apps Script + Google Sheets (Single-Tenant DB)
 *   v3: simple login (no 2FA), batched writes, no flush(),
 *       sessions cached + returned in bootstrap for instant,
 *       fully client-side dashboard/calendar/ledger reactivity.
 * ============================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SESSION_TIMEOUT_MIN = 120;
const TABLE_CACHE_TTL_SEC = 45;
const APP_NAME = 'MIX FIT JO';

const SHEETS = {
  CLIENTS: 'Clients', BOOKINGS: 'Bookings', SESSIONS: 'Sessions',
  SERVICES: 'Services', CLIENT_TYPES: 'ClientTypes', USERS: 'Users',
  AUDIT: 'AuditLog', SETTINGS: 'Settings'
};

// New columns are always appended at the END (never inserted in the middle) — this keeps a
// brand-new sheet and a migrated existing sheet identical, and lets migrateSchema_() safely
// add missing columns without touching existing ones.
const HEADERS = {
  Clients: ['ID','Name','DOB','Gender','Phone','CountryCode','EmergencyPhone','Occupation',
            'MaritalStatus','LeadSource','HealthAnswersJSON','HealthFilesJSON','ClientType',
            'TotalSpent','VisitsCount','CreatedAt','Age'],
  Bookings: ['ID','ClientID','ServiceIDs','SessionsCount','ScheduleJSON','StartDate','EndDate',
             'TotalValue','Paid','Remaining','Active','Status','Notes','CreatedAt'],
  Sessions: ['ID','BookingID','ClientID','Date','Time','Status','CreatedAt'],
  Services: ['ID','Name','Cost','DurationMinutes','CreatedAt'],
  ClientTypes: ['ID','Name','MinSpend','Color','CreatedAt'],
  Users: ['ID','Username','PasswordHash','Email','Role','FullName','CreatedAt'],
  AuditLog: ['ID','Timestamp','User','Action','Entity','EntityID','Details'],
  Settings: ['Key','Value']
};

const JSON_FIELDS = {
  Clients: ['HealthAnswersJSON', 'HealthFilesJSON'],
  Bookings: ['ServiceIDs', 'ScheduleJSON']
};

// Plain "yyyy-MM-dd" strings written to their own cell (not inside a JSON blob) are at risk of
// Google Sheets silently auto-converting them into a real Date-typed cell, which then comes back
// from getValues() as a JS Date object — and when that crosses the google.script.run bridge to the
// browser it serializes as a full ISO datetime ("...T00:00:00.000Z"), which an <input type="date">
// silently rejects and renders EMPTY. This is the root cause behind subscription dates (and the
// financial fields that visually sit right next to them) appearing to "empty out" when reopening an
// existing booking. rowsToObjects_ normalizes any such Date object back to a plain date string on
// every read, and migrateSchema_() repairs already-affected existing cells once.
const DATE_FIELDS = {
  Clients: ['DOB'],
  Bookings: ['StartDate', 'EndDate'],
  Sessions: ['Date']
};

const DEFAULT_SERVICES = ['تدريب شخصي','تدريب عائلات','تدريب مجموعات','جلسات علاجية','زيارة / استشارة','تجربة','حصة واحدة','تغذية','قياسات'];
const DEFAULT_CLIENT_TYPES = [
  { Name: 'Basic', MinSpend: 0, Color: '#8a8f98' },
  { Name: 'Standard', MinSpend: 150, Color: '#4f7dfb' },
  { Name: 'Premium', MinSpend: 400, Color: '#d5212c' },
  { Name: 'VIP', MinSpend: 800, Color: '#fbbf24' },
  { Name: 'Company', MinSpend: 1500, Color: '#34d399' }
];
const HEALTH_QUESTIONS = [
  'هل كان لديك أي مشاكل قلبية وطلب منك الطبيب عدم ممارسة النشاط البدني؟',
  'هل تشعر بألم في الصدر عند ممارسة النشاط البدني؟',
  'هل شعرت بأي ألم في الصدر خلال الشهر الماضي بالرغم من عدم ممارستك للنشاط البدني؟',
  'هل سبق وأن شعرت بدوار أو إغماء أو فقدان للاتزان؟',
  'هل تعاني من أية مشاكل أو إصابات في العظام أو المفاصل يمكن أن تزداد سوءاً مع النشاط البدني؟',
  'هل تعاني حالياً من ارتفاع ضغط الدم أو مشكلة قلبية تتطلب أدوية موصوفة؟',
  'هل تعاني من السكري من النوع الأول الذي يتطلب علاج بالأنسولين؟',
  'هل تعلم عن وجود أية أسباب أخرى تمنعك من ممارسة التمرين أو زيادة نشاطك البدني؟'
];

/* ============================================================
 *  BOOTSTRAP / SETUP  — safe to run any number of times
 * ============================================================ */

function setupDatabase() {
  migrateSchema_();

  Object.keys(HEADERS).forEach(name => {
    let sheet = SS.getSheetByName(name);
    if (!sheet) sheet = SS.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      const headerRow = sheet.getRange(1, 1, 1, HEADERS[name].length);
      headerRow.setValues([HEADERS[name]]);
      headerRow.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  const defaultSheet = SS.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0) SS.deleteSheet(defaultSheet);

  const usersSheet = sheet_(SHEETS.USERS);
  if (usersSheet.getLastRow() < 2) {
    appendRowFromHeaders_(usersSheet, HEADERS.Users, {
      ID: Utilities.getUuid(), Username: 'admin', PasswordHash: hashPassword_('Admin@123'),
      Email: '', Role: 'Owner', FullName: 'Administrator', CreatedAt: new Date()
    });
  }

  const servicesSheet = sheet_(SHEETS.SERVICES);
  if (servicesSheet.getLastRow() < 2) {
    DEFAULT_SERVICES.forEach(name => {
      appendRowFromHeaders_(servicesSheet, HEADERS.Services, {
        ID: Utilities.getUuid(), Name: name, Cost: 0, DurationMinutes: 60, CreatedAt: new Date()
      });
    });
  }

  const typesSheet = sheet_(SHEETS.CLIENT_TYPES);
  if (typesSheet.getLastRow() < 2) {
    DEFAULT_CLIENT_TYPES.forEach(ct => {
      appendRowFromHeaders_(typesSheet, HEADERS.ClientTypes, {
        ID: Utilities.getUuid(), Name: ct.Name, MinSpend: ct.MinSpend, Color: ct.Color, CreatedAt: new Date()
      });
    });
  }

  repairDateColumns_();
  clearAllTableCache_();
  Logger.log('MIX FIT JO database ready. Default login (first run only): admin / Admin@123 — change it from inside the app after logging in.');
}

/** Adds any headers missing from an already-existing sheet, at the end, without touching existing columns/rows. */
/** Forces plain-text format on date-string columns (prevents Sheets from silently auto-converting
 *  them into Date-typed cells going forward) and rewrites any cell that has ALREADY been
 *  auto-converted back into a plain "yyyy-MM-dd" string. Safe to run repeatedly. */
function repairDateColumns_() {
  const tz = Session.getScriptTimeZone();
  Object.keys(DATE_FIELDS).forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol === 0) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    DATE_FIELDS[sheetName].forEach(fieldName => {
      const colIndex = headers.indexOf(fieldName) + 1;
      if (colIndex === 0) return;
      const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
      range.setNumberFormat('@'); // plain text from now on — prevents future auto-conversion
      const values = range.getValues();
      let changed = false;
      const fixed = values.map(r => {
        if (r[0] instanceof Date) { changed = true; return [Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')]; }
        return r;
      });
      if (changed) { range.setValues(fixed); Logger.log('Repaired auto-converted date cells in ' + sheetName + '.' + fieldName); }
    });
  });
}

function migrateSchema_() {
  Object.keys(HEADERS).forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) return;
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;
    const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const missing = HEADERS[name].filter(h => currentHeaders.indexOf(h) === -1);
    if (!missing.length) return;

    const startCol = lastCol + 1;
    const headerRange = sheet.getRange(1, startCol, 1, missing.length);
    headerRange.setValues([missing]);
    headerRange.setFontWeight('bold');
    Logger.log('Migrated sheet "' + name + '": added column(s) ' + missing.join(', '));

    if (name === 'Clients' && missing.indexOf('Age') !== -1) backfillClientAges_(sheet);
  });
}

function backfillClientAges_(sheet) {
  const headers = getHeaders_(sheet);
  const dobCol = headers.indexOf('DOB') + 1;
  const ageCol = headers.indexOf('Age') + 1;
  const lastRow = sheet.getLastRow();
  if (dobCol === 0 || ageCol === 0 || lastRow < 2) return;
  const dobValues = sheet.getRange(2, dobCol, lastRow - 1, 1).getValues();
  const ages = dobValues.map(r => {
    if (!r[0]) return [''];
    const dobStr = r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[0]);
    return [calcAge_(dobStr)];
  });
  sheet.getRange(2, ageCol, ages.length, 1).setValues(ages);
}

/** Manual diagnostic only — run from the Apps Script editor. Reports missing sheets/columns without changing anything. */
function verifyDatabaseIntegrity() {
  const report = [];
  Object.keys(HEADERS).forEach(name => {
    const sheet = SS.getSheetByName(name);
    if (!sheet) { report.push('❌ Missing sheet: ' + name); return; }
    const lastCol = sheet.getLastColumn();
    const currentHeaders = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const missing = HEADERS[name].filter(h => currentHeaders.indexOf(h) === -1);
    if (missing.length) report.push('⚠ ' + name + ' — missing column(s): ' + missing.join(', '));
    else report.push('✅ ' + name + ' OK (' + Math.max(sheet.getLastRow() - 1, 0) + ' data row(s))');
  });
  Logger.log(report.join('\n'));
  return report;
}

function hashPassword_(pwd) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function calcAge_(dobStr) {
  if (!dobStr) return '';
  const dob = new Date(dobStr);
  if (isNaN(dob.getTime())) return '';
  const today = new Date();
  let age = today.getFullYear() - dob.getFullYear();
  const m = today.getMonth() - dob.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
  return age;
}

/* ============================================================
 *  WEB APP ENTRY POINT
 * ============================================================ */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
 *  GENERIC SHEET HELPERS
 * ============================================================ */

function sheet_(name) {
  const s = SS.getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function appendRowFromHeaders_(sheet, headers, dataObj) {
  const row = headers.map(h => {
    let val = dataObj[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  });
  sheet.appendRow(row);
}

/** Writes many rows in ONE API call instead of looping appendRow — critical for booking
 *  packages with many sessions (12 sessions = 1 call instead of 12). */
function appendRowsBatch_(sheet, headers, dataObjArray) {
  if (!dataObjArray.length) return;
  const rows = dataObjArray.map(dataObj => headers.map(h => {
    let val = dataObj[h];
    if (val === undefined || val === null) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return val;
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

/** Replaces ALL data rows of a sheet in a single pass (clear + one setValues) — used for bulk
 *  delete/filter operations so removing N rows costs O(1) API calls instead of O(N) deleteRow calls. */
function replaceAllRows_(sheet, headers, objectsArray) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (objectsArray.length) {
    const rows = objectsArray.map(obj => headers.map(h => {
      let val = obj[h];
      if (val === undefined || val === null) return '';
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    }));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function setCellsFromObject_(sheet, row, dataObj) {
  const headers = getHeaders_(sheet);
  Object.keys(dataObj).forEach(key => {
    const col = headers.indexOf(key);
    if (col === -1) return;
    let val = dataObj[key];
    if (typeof val === 'object' && val !== null) val = JSON.stringify(val);
    sheet.getRange(row, col + 1).setValue(val);
  });
}

function rowsToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const jsonFields = JSON_FIELDS[sheet.getName()] || [];
  const dateFields = DATE_FIELDS[sheet.getName()] || [];
  const tz = Session.getScriptTimeZone();
  return values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = r[i];
        if (jsonFields.indexOf(h) !== -1) {
          try { val = val ? JSON.parse(val) : []; } catch (e) { val = []; }
        } else if (dateFields.indexOf(h) !== -1 && val instanceof Date) {
          val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        }
        obj[h] = val;
      });
      return obj;
    });
}

function findRowIndexById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function addDaysStr_(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function rangeBoundsStr_(range) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  let start, end;
  if (range === 'today') {
    start = new Date(now); end = new Date(now);
  } else if (range === 'week') {
    const dow = now.getDay();
    start = new Date(now); start.setDate(now.getDate() - dow);
    end = new Date(start); end.setDate(start.getDate() + 6);
  } else if (range === 'year') {
    start = new Date(now.getFullYear(), 0, 1);
    end = new Date(now.getFullYear(), 11, 31);
  } else {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
    end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  }
  return { start: Utilities.formatDate(start, tz, 'yyyy-MM-dd'), end: Utilities.formatDate(end, tz, 'yyyy-MM-dd') };
}

/* ============================================================
 *  READ CACHE  — cuts repeat full-sheet scans to ~45s TTL,
 *  invalidated immediately on any write to that table.
 * ============================================================ */

function getCachedRows_(sheetName) {
  const cache = CacheService.getScriptCache();
  const key = 'tbl_' + sheetName;
  try {
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { /* fall through to a fresh read */ }

  const rows = rowsToObjects_(sheet_(sheetName));
  try { cache.put(key, JSON.stringify(rows), TABLE_CACHE_TTL_SEC); } catch (e) { /* value too large — skip caching, still return fresh data */ }
  return rows;
}
function invalidateTableCache_(sheetName) { CacheService.getScriptCache().remove('tbl_' + sheetName); }
function clearAllTableCache_() { Object.keys(SHEETS).forEach(k => invalidateTableCache_(SHEETS[k])); }

/* ============================================================
 *  AUDIT LOG
 * ============================================================ */

function logAction_(user, action, entity, entityId, details) {
  appendRowFromHeaders_(sheet_(SHEETS.AUDIT), HEADERS.AuditLog, {
    ID: Utilities.getUuid(), Timestamp: new Date(), User: user, Action: action,
    Entity: entity, EntityID: entityId, Details: details || ''
  });
}

function getAuditLog(token) {
  requireSession_(token);
  return rowsToObjects_(sheet_(SHEETS.AUDIT)).reverse().slice(0, 300);
}

/* ============================================================
 *  AUTH — simple username/password (no 2FA — removed per request)
 * ============================================================ */

function findUser_(username) {
  return rowsToObjects_(sheet_(SHEETS.USERS)).find(u => u.Username === username);
}

function login(username, password) {
  const user = findUser_(username);
  if (!user || user.PasswordHash !== hashPassword_(password)) {
    return { success: false, message: 'بيانات الدخول غير صحيحة' };
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, JSON.stringify({
    userId: user.ID, username: user.Username, role: user.Role, fullName: user.FullName
  }), SESSION_TIMEOUT_MIN * 60);
  logAction_(user.Username, 'LOGIN', 'User', user.ID, 'تسجيل دخول ناجح');
  return { success: true, token: token, user: { username: user.Username, role: user.Role, fullName: user.FullName } };
}

function checkSession(token) {
  if (!token) return { valid: false };
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session_' + token);
  if (!raw) return { valid: false };
  cache.put('session_' + token, raw, SESSION_TIMEOUT_MIN * 60);
  return { valid: true, user: JSON.parse(raw) };
}

function logout(token) {
  CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function requireSession_(token) {
  const check = checkSession(token);
  if (!check.valid) throw new Error('SESSION_EXPIRED');
  return check.user;
}

function changePassword(token, oldPassword, newPassword) {
  const sessionUser = requireSession_(token);
  const sheet = sheet_(SHEETS.USERS);
  const row = findRowIndexById_(sheet, sessionUser.userId);
  if (row === -1) throw new Error('User not found');
  const headers = getHeaders_(sheet);
  const currentHash = sheet.getRange(row, headers.indexOf('PasswordHash') + 1).getValue();
  if (currentHash !== hashPassword_(oldPassword)) {
    return { success: false, message: 'كلمة المرور الحالية غير صحيحة' };
  }
  sheet.getRange(row, headers.indexOf('PasswordHash') + 1).setValue(hashPassword_(newPassword));
  logAction_(sessionUser.username, 'UPDATE', 'User', sessionUser.userId, 'تغيير كلمة المرور');
  return { success: true };
}

/* ============================================================
 *  BOOTSTRAP — ONE round trip for everything the app needs on load.
 *  Includes sessions now too, so Dashboard/Calendar/Ledger can compute
 *  everything reactively on the client with ZERO further network calls.
 * ============================================================ */

function getBootstrapData(token) {
  requireSession_(token);
  return {
    clients: getClientsLight_(),
    services: getCachedRows_(SHEETS.SERVICES),
    clientTypes: getCachedRows_(SHEETS.CLIENT_TYPES),
    bookings: getCachedRows_(SHEETS.BOOKINGS),
    sessions: getCachedRows_(SHEETS.SESSIONS)
  };
}

/* ============================================================
 *  CLIENTS
 *
 *  The health questionnaire (8 Arabic questions + notes) and the file
 *  list are the heaviest fields on a client row, repeated for every
 *  client in every list/bootstrap payload — with enough clients this
 *  both risks exceeding the ~100KB CacheService value limit (silently
 *  disabling caching) and slows every single load. getClientsLight_()
 *  strips them down to two small counts for list views; the full
 *  detail is fetched on-demand for exactly one client at a time via
 *  getClientDetail(), only when that client's modal is actually opened.
 * ============================================================ */

function getClientsLight_() {
  return getCachedRows_(SHEETS.CLIENTS).map(c => {
    const light = Object.assign({}, c);
    light.HealthFilesCount = (c.HealthFilesJSON || []).length;
    light.HealthYesCount = (c.HealthAnswersJSON || []).filter(a => a.answer === 'نعم').length;
    delete light.HealthAnswersJSON;
    delete light.HealthFilesJSON;
    return light;
  });
}

function getClients(token) { requireSession_(token); return getClientsLight_(); }

/** Full detail for exactly one client (including health answers + files) — a fast targeted
 *  single-row read, not a full-sheet scan. */
function getClientDetail(token, clientId) {
  requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, clientId);
  if (row === -1) throw new Error('Client not found');
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const jsonFields = JSON_FIELDS.Clients;
  const dateFields = DATE_FIELDS.Clients;
  const tz = Session.getScriptTimeZone();
  const obj = {};
  headers.forEach((h, i) => {
    let val = values[i];
    if (jsonFields.indexOf(h) !== -1) {
      try { val = val ? JSON.parse(val) : []; } catch (e) { val = []; }
    } else if (dateFields.indexOf(h) !== -1 && val instanceof Date) {
      val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
    }
    obj[h] = val;
  });
  return obj;
}

function addClient(token, data) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENTS);
  const id = Utilities.getUuid();
  const age = calcAge_(data.DOB);
  const record = {
    ID: id, Name: data.Name, DOB: data.DOB || '', Gender: data.Gender || '',
    Phone: data.Phone, CountryCode: data.CountryCode || '+962', EmergencyPhone: data.EmergencyPhone || '',
    Occupation: data.Occupation || '', MaritalStatus: data.MaritalStatus || '', LeadSource: data.LeadSource || '',
    HealthAnswersJSON: data.HealthAnswers || [], HealthFilesJSON: [], ClientType: 'Basic',
    TotalSpent: 0, VisitsCount: 0, CreatedAt: new Date(), Age: age
  };
  appendRowFromHeaders_(sheet, HEADERS.Clients, record);
  invalidateTableCache_(SHEETS.CLIENTS);
  logAction_(user.username, 'CREATE', 'Client', id, 'إضافة عميل: ' + data.Name);
  return { success: true, id: id, client: record };
}

function updateClient(token, id, data) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Client not found');
  const patch = {};
  ['Name','DOB','Gender','Phone','CountryCode','EmergencyPhone','Occupation','MaritalStatus','LeadSource'].forEach(f => {
    if (data[f] !== undefined) patch[f] = data[f];
  });
  if (data.HealthAnswers !== undefined) patch.HealthAnswersJSON = data.HealthAnswers;
  if (data.DOB !== undefined) patch.Age = calcAge_(data.DOB);
  setCellsFromObject_(sheet, row, patch);
  invalidateTableCache_(SHEETS.CLIENTS);
  logAction_(user.username, 'UPDATE', 'Client', id, 'تعديل بيانات عميل');
  return { success: true, patch: patch };
}

function deleteClient(token, id) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Client not found');

  // Bulk-remove this client's bookings + their sessions in O(1) passes (not a per-row loop).
  const bookingsSheet = sheet_(SHEETS.BOOKINGS);
  const allBookings = rowsToObjects_(bookingsSheet);
  const removedBookingIds = allBookings.filter(b => String(b.ClientID) === String(id)).map(b => b.ID);
  const remainingBookings = allBookings.filter(b => String(b.ClientID) !== String(id));
  replaceAllRows_(bookingsSheet, HEADERS.Bookings, remainingBookings);

  if (removedBookingIds.length) {
    const sessionsSheet = sheet_(SHEETS.SESSIONS);
    const remainingSessions = rowsToObjects_(sessionsSheet).filter(s => removedBookingIds.indexOf(s.BookingID) === -1);
    replaceAllRows_(sessionsSheet, HEADERS.Sessions, remainingSessions);
    invalidateTableCache_(SHEETS.SESSIONS);
  }

  sheet.deleteRow(row);
  invalidateTableCache_(SHEETS.CLIENTS);
  invalidateTableCache_(SHEETS.BOOKINGS);
  logAction_(user.username, 'DELETE', 'Client', id, 'حذف عميل و ' + removedBookingIds.length + ' حجز مرتبط');
  return { success: true, removedBookingIds: removedBookingIds };
}

/** DriveApp.getFoldersByName() searches the whole Drive and is genuinely slow — caching the
 *  resolved folder ID makes every upload after the first one for the same client near-instant. */
function getOrCreateFolder_(name, parent) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'folder_' + (parent ? parent.getId() : 'root') + '_' + name;
  const cachedId = cache.get(cacheKey);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch (e) { /* folder was removed — fall through to search */ }
  }
  const root = parent || DriveApp.getRootFolder();
  const it = root.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : root.createFolder(name);
  try { cache.put(cacheKey, folder.getId(), 21600); } catch (e) { /* ignore */ }
  return folder;
}

function uploadHealthFile(token, clientId, base64Data, mimeType, fileName) {
  const user = requireSession_(token);
  const rootFolder = getOrCreateFolder_('MixFitJo_HealthFiles');
  const clientFolder = getOrCreateFolder_(String(clientId), rootFolder);

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = clientFolder.createFile(blob);
  // Some Workspace domains restrict external sharing and this call can throw — never let a sharing
  // policy failure abort the whole upload; the file is still created and linked either way.
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* org policy may block this — link still works for signed-in staff */ }

  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, clientId);
  if (row === -1) throw new Error('Client not found');
  const headers = getHeaders_(sheet);
  const filesCol = headers.indexOf('HealthFilesJSON') + 1;
  let files = [];
  try { files = JSON.parse(sheet.getRange(row, filesCol).getValue() || '[]'); } catch (e) { files = []; }
  files.push({ id: file.getId(), name: fileName, url: file.getUrl() });
  sheet.getRange(row, filesCol).setValue(JSON.stringify(files));
  invalidateTableCache_(SHEETS.CLIENTS);

  logAction_(user.username, 'UPDATE', 'Client', clientId, 'رفع ملف صحي: ' + fileName);
  return { success: true, files: files };
}

function deleteHealthFile(token, clientId, fileId) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, clientId);
  if (row === -1) throw new Error('Client not found');
  const headers = getHeaders_(sheet);
  const filesCol = headers.indexOf('HealthFilesJSON') + 1;
  let files = [];
  try { files = JSON.parse(sheet.getRange(row, filesCol).getValue() || '[]'); } catch (e) { files = []; }
  files = files.filter(f => f.id !== fileId);
  sheet.getRange(row, filesCol).setValue(JSON.stringify(files));
  invalidateTableCache_(SHEETS.CLIENTS);
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* ignore */ }
  logAction_(user.username, 'UPDATE', 'Client', clientId, 'حذف ملف صحي');
  return { success: true, files: files };
}

/** Recomputes a client's cumulative spend, visit count, and auto-assigned tier/badge.
 *  Accepts pre-loaded bookings/sessions to avoid redundant reads when called in a loop. */
function recalcClientTotals_(clientId, preloadedBookings, preloadedSessions) {
  const bookings = (preloadedBookings || getCachedRows_(SHEETS.BOOKINGS)).filter(b => String(b.ClientID) === String(clientId));
  const totalSpent = bookings.reduce((sum, b) => sum + (Number(b.Paid) || 0), 0);
  const sessions = (preloadedSessions || getCachedRows_(SHEETS.SESSIONS)).filter(s => String(s.ClientID) === String(clientId));
  const visits = sessions.filter(s => s.Status === 'حضر').length;

  const types = getCachedRows_(SHEETS.CLIENT_TYPES).slice().sort((a, b) => Number(b.MinSpend) - Number(a.MinSpend));
  let assignedType = types.length ? types[types.length - 1].Name : 'Basic';
  for (const ty of types) { if (totalSpent >= Number(ty.MinSpend)) { assignedType = ty.Name; break; } }

  const sheet = sheet_(SHEETS.CLIENTS);
  const row = findRowIndexById_(sheet, clientId);
  if (row !== -1) {
    setCellsFromObject_(sheet, row, { TotalSpent: totalSpent, VisitsCount: visits, ClientType: assignedType });
  }
  invalidateTableCache_(SHEETS.CLIENTS);
  return { TotalSpent: totalSpent, VisitsCount: visits, ClientType: assignedType };
}

/* ============================================================
 *  SERVICES  (Settings tab 2)
 * ============================================================ */

function getServices(token) { requireSession_(token); return getCachedRows_(SHEETS.SERVICES); }

function addService(token, data) {
  const user = requireSession_(token);
  const id = Utilities.getUuid();
  const record = { ID: id, Name: data.Name, Cost: data.Cost || 0, DurationMinutes: data.DurationMinutes || 60, CreatedAt: new Date() };
  appendRowFromHeaders_(sheet_(SHEETS.SERVICES), HEADERS.Services, record);
  invalidateTableCache_(SHEETS.SERVICES);
  logAction_(user.username, 'CREATE', 'Service', id, 'إضافة خدمة: ' + data.Name);
  return { success: true, id: id, service: record };
}

function updateService(token, id, data) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.SERVICES);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Service not found');
  setCellsFromObject_(sheet, row, data);
  invalidateTableCache_(SHEETS.SERVICES);
  logAction_(user.username, 'UPDATE', 'Service', id, 'تعديل خدمة');
  return { success: true, patch: data };
}

function deleteService(token, id) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.SERVICES);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Service not found');
  sheet.deleteRow(row);
  invalidateTableCache_(SHEETS.SERVICES);
  logAction_(user.username, 'DELETE', 'Service', id, 'حذف خدمة');
  return { success: true };
}

/* ============================================================
 *  CLIENT TYPES  (Settings tab 1)
 * ============================================================ */

function getClientTypes(token) { requireSession_(token); return getCachedRows_(SHEETS.CLIENT_TYPES); }

function addClientType(token, data) {
  const user = requireSession_(token);
  const id = Utilities.getUuid();
  const record = { ID: id, Name: data.Name, MinSpend: data.MinSpend || 0, Color: data.Color || '#4f7dfb', CreatedAt: new Date() };
  appendRowFromHeaders_(sheet_(SHEETS.CLIENT_TYPES), HEADERS.ClientTypes, record);
  invalidateTableCache_(SHEETS.CLIENT_TYPES);
  logAction_(user.username, 'CREATE', 'ClientType', id, 'إضافة تصنيف: ' + data.Name);
  return { success: true, id: id, type: record };
}

function updateClientType(token, id, data) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENT_TYPES);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('ClientType not found');
  setCellsFromObject_(sheet, row, data);
  invalidateTableCache_(SHEETS.CLIENT_TYPES);
  logAction_(user.username, 'UPDATE', 'ClientType', id, 'تعديل تصنيف');
  return { success: true, patch: data };
}

function deleteClientType(token, id) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.CLIENT_TYPES);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('ClientType not found');
  sheet.deleteRow(row);
  invalidateTableCache_(SHEETS.CLIENT_TYPES);
  logAction_(user.username, 'DELETE', 'ClientType', id, 'حذف تصنيف');
  return { success: true };
}

/* ============================================================
 *  BOOKINGS  (subscriptions/packages)
 * ============================================================ */

function getBookings(token) { requireSession_(token); return getCachedRows_(SHEETS.BOOKINGS); }

function computeEndDate_(scheduleJSON) {
  if (!scheduleJSON || !scheduleJSON.length) return '';
  return scheduleJSON.map(s => s.date).sort().pop();
}

/** Builds the session record objects for a schedule WITHOUT writing them yet (lets callers batch). */
function buildSessionRecords_(bookingId, clientId, scheduleJSON) {
  return (scheduleJSON || []).map(s => ({
    ID: Utilities.getUuid(), BookingID: bookingId, ClientID: clientId,
    Date: s.date, Time: s.time || '', Status: 'مجدولة', CreatedAt: new Date()
  }));
}

function addBooking(token, data) {
  const user = requireSession_(token);
  const id = Utilities.getUuid();
  const endDate = computeEndDate_(data.ScheduleJSON) || data.StartDate;
  const paid = Number(data.Paid) || 0;
  const total = Number(data.TotalValue) || 0;

  const record = {
    ID: id, ClientID: data.ClientID, ServiceIDs: data.ServiceIDs || [],
    SessionsCount: data.SessionsCount || (data.ScheduleJSON || []).length,
    ScheduleJSON: data.ScheduleJSON || [], StartDate: data.StartDate, EndDate: endDate,
    TotalValue: total, Paid: paid, Remaining: total - paid,
    Active: data.Active !== false, Status: 'نشط', Notes: data.Notes || '', CreatedAt: new Date()
  };
  appendRowFromHeaders_(sheet_(SHEETS.BOOKINGS), HEADERS.Bookings, record);

  const sessionRecords = buildSessionRecords_(id, data.ClientID, data.ScheduleJSON);
  appendRowsBatch_(sheet_(SHEETS.SESSIONS), HEADERS.Sessions, sessionRecords);

  invalidateTableCache_(SHEETS.BOOKINGS);
  invalidateTableCache_(SHEETS.SESSIONS);
  const totals = recalcClientTotals_(data.ClientID);
  logAction_(user.username, 'CREATE', 'Booking', id, 'إضافة حجز/اشتراك جديد');
  return { success: true, id: id, booking: record, sessions: sessionRecords, clientTotals: totals };
}

function updateBooking(token, id, data) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.BOOKINGS);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Booking not found');

  const patch = {};
  ['ClientID','ServiceIDs','StartDate','Notes','Status'].forEach(f => { if (data[f] !== undefined) patch[f] = data[f]; });

  let newSessionRecords = null;
  if (data.ScheduleJSON !== undefined) {
    patch.ScheduleJSON = data.ScheduleJSON;
    patch.SessionsCount = data.ScheduleJSON.length;
    patch.EndDate = computeEndDate_(data.ScheduleJSON);
    const clientId = data.ClientID || sheet.getRange(row, getHeaders_(sheet).indexOf('ClientID') + 1).getValue();

    // Replace this booking's sessions in ONE read + ONE write of the Sessions sheet (not a
    // delete-loop followed by an append-loop).
    const sessionsSheet = sheet_(SHEETS.SESSIONS);
    const otherSessions = rowsToObjects_(sessionsSheet).filter(s => String(s.BookingID) !== String(id));
    newSessionRecords = buildSessionRecords_(id, clientId, data.ScheduleJSON);
    replaceAllRows_(sessionsSheet, HEADERS.Sessions, otherSessions.concat(newSessionRecords));
    invalidateTableCache_(SHEETS.SESSIONS);
  }
  if (data.TotalValue !== undefined || data.Paid !== undefined) {
    const headers = getHeaders_(sheet);
    const currentTotal = data.TotalValue !== undefined ? Number(data.TotalValue) : Number(sheet.getRange(row, headers.indexOf('TotalValue') + 1).getValue());
    const currentPaid = data.Paid !== undefined ? Number(data.Paid) : Number(sheet.getRange(row, headers.indexOf('Paid') + 1).getValue());
    patch.TotalValue = currentTotal; patch.Paid = currentPaid; patch.Remaining = currentTotal - currentPaid;
  }
  if (data.Active !== undefined) patch.Active = data.Active;

  setCellsFromObject_(sheet, row, patch);
  invalidateTableCache_(SHEETS.BOOKINGS);
  const clientId = data.ClientID || sheet.getRange(row, getHeaders_(sheet).indexOf('ClientID') + 1).getValue();
  const totals = recalcClientTotals_(clientId);
  logAction_(user.username, 'UPDATE', 'Booking', id, 'تعديل حجز/اشتراك');
  return { success: true, patch: patch, sessions: newSessionRecords, clientTotals: totals, clientId: clientId };
}

function toggleBookingActive(token, id, active) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.BOOKINGS);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Booking not found');
  setCellsFromObject_(sheet, row, { Active: active });
  invalidateTableCache_(SHEETS.BOOKINGS);
  logAction_(user.username, 'UPDATE', 'Booking', id, active ? 'تفعيل الاشتراك' : 'إيقاف تفعيل الاشتراك');
  return { success: true, active: active };
}

function deleteBooking(token, id) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.BOOKINGS);
  const row = findRowIndexById_(sheet, id);
  if (row === -1) throw new Error('Booking not found');
  const clientId = sheet.getRange(row, getHeaders_(sheet).indexOf('ClientID') + 1).getValue();

  const sessionsSheet = sheet_(SHEETS.SESSIONS);
  const remainingSessions = rowsToObjects_(sessionsSheet).filter(s => String(s.BookingID) !== String(id));
  replaceAllRows_(sessionsSheet, HEADERS.Sessions, remainingSessions);
  invalidateTableCache_(SHEETS.SESSIONS);

  sheet.deleteRow(row);
  invalidateTableCache_(SHEETS.BOOKINGS);
  const totals = recalcClientTotals_(clientId);
  logAction_(user.username, 'DELETE', 'Booking', id, 'حذف حجز/اشتراك');
  return { success: true, clientId: clientId, clientTotals: totals };
}

/* ============================================================
 *  SESSIONS  (individual visits — attendance)
 * ============================================================ */

function getSessions(token) { requireSession_(token); return getCachedRows_(SHEETS.SESSIONS); }

function updateSessionStatus(token, sessionId, status) {
  const user = requireSession_(token);
  const sheet = sheet_(SHEETS.SESSIONS);
  const row = findRowIndexById_(sheet, sessionId);
  if (row === -1) throw new Error('Session not found');
  const headers = getHeaders_(sheet);
  const clientId = sheet.getRange(row, headers.indexOf('ClientID') + 1).getValue();
  setCellsFromObject_(sheet, row, { Status: status });
  invalidateTableCache_(SHEETS.SESSIONS);
  const totals = recalcClientTotals_(clientId);
  logAction_(user.username, 'UPDATE', 'Session', sessionId, 'تحديث حالة الحصة: ' + status);
  return { success: true, clientId: clientId, clientTotals: totals };
}

/* ============================================================
 *  CLIENT LEDGER  (cumulative spend / badges) — kept as a backend
 *  endpoint for reference/print use; the app itself now computes
 *  this client-side from the bootstrap payload for instant updates.
 * ============================================================ */

function getClientLedger(token) {
  requireSession_(token);
  const clients = getCachedRows_(SHEETS.CLIENTS);
  const types = getCachedRows_(SHEETS.CLIENT_TYPES);
  const bookings = getCachedRows_(SHEETS.BOOKINGS);
  return clients.map(c => {
    const type = types.find(t => t.Name === c.ClientType) || {};
    const clientBookings = bookings.filter(b => String(b.ClientID) === String(c.ID));
    const activeBooking = clientBookings.filter(b => b.Active).sort((a, b) => b.EndDate.localeCompare(a.EndDate))[0];
    return {
      ID: c.ID, Name: c.Name, Phone: c.Phone, CountryCode: c.CountryCode,
      ClientType: c.ClientType, TypeColor: type.Color || '#8a8f98',
      TotalSpent: c.TotalSpent, VisitsCount: c.VisitsCount,
      LastEndDate: activeBooking ? activeBooking.EndDate : ''
    };
  }).sort((a, b) => Number(b.TotalSpent) - Number(a.TotalSpent));
}

/* ============================================================
 *  WHATSAPP INTEGRATION  (wa.me — no paid API)
 * ============================================================ */

function buildWhatsAppLink(token, clientId, templateType, bookingId) {
  requireSession_(token);
  const client = getCachedRows_(SHEETS.CLIENTS).find(c => String(c.ID) === String(clientId));
  if (!client) throw new Error('Client not found');

  const bookings = getCachedRows_(SHEETS.BOOKINGS).filter(b => String(b.ClientID) === String(clientId));
  const booking = bookingId ? bookings.find(b => b.ID === bookingId) : bookings.sort((a, b) => b.CreatedAt < a.CreatedAt ? -1 : 1)[0];
  const services = getCachedRows_(SHEETS.SERVICES);
  const serviceNames = booking ? (booking.ServiceIDs || []).map(sid => (services.find(s => s.ID === sid) || {}).Name).filter(Boolean).join('، ') : '';

  const nextSession = getCachedRows_(SHEETS.SESSIONS)
    .filter(s => String(s.ClientID) === String(clientId) && s.Status === 'مجدولة' && s.Date >= todayStr_())
    .sort((a, b) => (a.Date + a.Time).localeCompare(b.Date + b.Time))[0];

  let message = '';
  switch (templateType) {
    case 'newClient':
      message = `مرحباً ${client.Name}، أهلاً بك في عائلة MIX FIT JO! 💪\nتم تسجيل اشتراكك في: ${serviceNames}\nعدد الحصص: ${booking ? booking.SessionsCount : ''}\nمن ${booking ? booking.StartDate : ''} إلى ${booking ? booking.EndDate : ''}\nنتمنى لك رحلة رياضية ملهمة!`;
      break;
    case 'reminder':
      message = `تذكير: لديك حصة ${nextSession ? 'بتاريخ ' + nextSession.Date + ' الساعة ' + nextSession.Time : 'قريباً'}. بانتظارك في MIX FIT JO 🔥`;
      break;
    case 'financial':
      message = `مرحباً ${client.Name}، بيانات اشتراكك المالية:\nالإجمالي: ${booking ? booking.TotalValue : 0} د.أ\nالمدفوع: ${booking ? booking.Paid : 0} د.أ\nالمتبقي: ${booking ? booking.Remaining : 0} د.أ`;
      break;
    case 'postpone':
      message = `مرحباً ${client.Name}، نأسف للإزعاج، نظراً لظرف طارئ يرجى العلم أن حصتك ${nextSession ? 'بتاريخ ' + nextSession.Date : ''} سيتم تأجيلها. سنتواصل معك لتحديد موعد بديل.`;
      break;
    case 'thanks':
      message = `شكراً لانضمامك لعائلة MIX FIT JO يا ${client.Name}! 🏆 نفتخر بك كعضو ${client.ClientType}. نتطلع لرؤيتك تحقق أهدافك معنا!`;
      break;
    case 'expiry':
      message = `مرحباً ${client.Name}، نود تذكيرك بأن اشتراكك سينتهي بتاريخ ${booking ? booking.EndDate : ''}. نسعد بتجديد اشتراكك لمواصلة رحلتك الرياضية معنا 💪`;
      break;
    default:
      message = `مرحباً ${client.Name}،`;
  }

  const phone = String(client.CountryCode || '').replace('+', '') + String(client.Phone).replace(/^0/, '');
  return `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
}
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}