/**
 * Centered in Christ staff/admin backend for Google Apps Script.
 *
 * Deploy this Apps Script project as a Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * SHEET_ID can be saved as a Script Property named SHEET_ID.
 * If it is not set, the fallback sheet ID below is used.
 */
const FALLBACK_SHEET_ID = '1PdYM_zYNnWEz5A36v31erUBvRlPLwUK4dljEc5GWOIo';
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || FALLBACK_SHEET_ID;
const PASSWORD_SHEET_NAME = 'Password';
const PASSWORD_RANGE = 'A1:B2';
const PASSWORD_UPDATE_RANGE = 'A2:B2';
const PRAYER_REQUESTS_SHEET_NAME = 'Form Responses';
const PRAYER_REQUEST_TEXT_HEADER = 'Please provide as much detailed information as possible';
const PRAYER_REQUEST_CONFIDENTIAL_HEADER = 'Is this confidential?';
const PRAYER_REQUEST_FIRST_NAME_HEADER = 'First Name';
const PRAYER_REQUEST_SHARE_VALUE = 'No, please share with as many people as possible';
const PRAYER_REQUEST_SHARE_PREFIX = 'no';
const PRAYER_MODERATION_SHEET_NAME = 'PrayerModeration';
const PRAYER_MODERATION_HEADERS = ['requestId', 'status', 'requestText', 'createdAt', 'updatedAt', 'firstName'];
const PRAYER_STATUS_APPROVED = 'approved';
const PRAYER_STATUS_DENIED = 'denied';
const SESSION_SHEET_NAME = 'StaffSessions';
const SESSION_TTL_MS = 12 * 60 * 60 * 1000;
const DEFAULT_YOUTUBE_CHANNEL_ID = 'UCubyomJfZo_C11A5fxRWnGw';

function doPost(e) {
  return handleRequest_(e);
}

function doGet(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  try {
    const request = parseRequest_(e);
    const action = String(request.action || '');

    if (action === 'ping' || (!action && !hasPostBody_(e))) {
      return json_({
        ok: true,
        service: 'centeredinchrist-staff-auth',
        action: action || 'status',
        timestamp: new Date().toISOString()
      });
    }
    if (action === 'login') {
      return json_(login_(request));
    }
    if (action === 'session') {
      return json_(validateSession_(request.token));
    }
    if (action === 'logout') {
      return json_(logout_(request.token));
    }
    if (action === 'updatePasswords') {
      return json_(updatePasswords_(request));
    }
    if (action === 'prayerRequests') {
      return json_(prayerRequests_(request));
    }
    if (action === 'moderatePrayerRequest') {
      return json_(moderatePrayerRequest_(request));
    }
    if (action === 'publicPrayerRequests') {
      return json_(publicPrayerRequests_());
    }
    if (action === 'removePublicPrayerRequest') {
      return json_(removePublicPrayerRequest_(request));
    }
    if (action === 'liveStatus') {
      return json_(liveStatus_());
    }

    log_('unknown action', { action: action, requestKeys: Object.keys(request) });
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (error) {
    log_('request failed', { error: error.message });
    return json_({ ok: false, error: 'Authentication service is unavailable.' });
  }
}

function parseRequest_(e) {
  if (!e) {
    return {};
  }

  const parameters = e.parameter || {};
  const content = e.postData && e.postData.contents ? String(e.postData.contents) : '';
  const type = e.postData && e.postData.type ? String(e.postData.type) : '';

  if (!content) {
    return parameters;
  }

  if (type.indexOf('application/json') !== -1 || content.trim().charAt(0) === '{') {
    return Object.assign({}, parameters, JSON.parse(content));
  }

  if (type.indexOf('application/x-www-form-urlencoded') !== -1 || content.indexOf('=') !== -1) {
    return Object.assign({}, parameters, parseFormBody_(content));
  }

  return parameters;
}

function hasPostBody_(e) {
  return Boolean(e && e.postData && e.postData.contents);
}

function parseFormBody_(content) {
  const parsed = {};
  content.split('&').forEach(function(part) {
    if (!part) {
      return;
    }

    const separator = part.indexOf('=');
    const rawKey = separator === -1 ? part : part.slice(0, separator);
    const rawValue = separator === -1 ? '' : part.slice(separator + 1);
    const key = decodeFormComponent_(rawKey);
    if (key) {
      parsed[key] = decodeFormComponent_(rawValue);
    }
  });
  return parsed;
}

function decodeFormComponent_(value) {
  return decodeURIComponent(String(value || '').replace(/\+/g, ' '));
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function spreadsheet_() {
  const sheetId = String(SHEET_ID || '').trim();
  if (!sheetId) {
    throw new Error('SHEET_ID is not configured.');
  }
  return SpreadsheetApp.openById(sheetId);
}

function passwordSheet_() {
  const sheet = spreadsheet_().getSheetByName(PASSWORD_SHEET_NAME);
  if (!sheet) {
    throw new Error('Password sheet was not found.');
  }
  return sheet;
}

function prayerRequestsSheet_() {
  const sheet = spreadsheet_().getSheetByName(PRAYER_REQUESTS_SHEET_NAME);
  if (!sheet) {
    throw new Error('Prayer requests sheet was not found.');
  }
  return sheet;
}

function prayerModerationSheet_() {
  const spreadsheet = spreadsheet_();
  let sheet = spreadsheet.getSheetByName(PRAYER_MODERATION_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PRAYER_MODERATION_SHEET_NAME);
    sheet.appendRow(PRAYER_MODERATION_HEADERS);
    sheet.hideSheet();
  } else {
    const headerRange = sheet.getRange(1, 1, 1, PRAYER_MODERATION_HEADERS.length);
    const headers = headerRange.getValues()[0].map(function(header) {
      return String(header || '');
    });
    const headersNeedUpdate = PRAYER_MODERATION_HEADERS.some(function(header, index) {
      return headers[index] !== header;
    });
    if (headersNeedUpdate) {
      headerRange.setValues([PRAYER_MODERATION_HEADERS]);
    }
  }
  return sheet;
}

function sessionSheet_() {
  const spreadsheet = spreadsheet_();
  let sheet = spreadsheet.getSheetByName(SESSION_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SESSION_SHEET_NAME);
    sheet.appendRow(['tokenHash', 'role', 'passwordHash', 'createdAt', 'lastSeenAt', 'expiresAt', 'revokedAt']);
    sheet.hideSheet();
  }
  return sheet;
}

function readPasswords_() {
  const values = passwordSheet_().getRange(PASSWORD_RANGE).getValues();
  const headers = values[0];
  const passwords = values[1];
  const adminIndex = headers.indexOf('ADMIN');
  const staffIndex = headers.indexOf('STAFF');

  if (adminIndex === -1 || staffIndex === -1) {
    throw new Error('Password sheet must have ADMIN and STAFF in row 1.');
  }

  return {
    admin: String(passwords[adminIndex] || ''),
    staff: String(passwords[staffIndex] || '')
  };
}

function writePasswords_(passwords) {
  passwordSheet_().getRange(PASSWORD_UPDATE_RANGE).setValues([[passwords.admin, passwords.staff]]);
}

function login_(request) {
  const password = String(request.password || '');
  if (!password) {
    return { ok: false, error: 'Please enter a password.' };
  }

  const passwords = readPasswords_();
  let role = '';
  if (passwords.admin && password === passwords.admin) {
    role = 'admin';
  } else if (passwords.staff && password === passwords.staff) {
    role = 'staff';
  }

  if (!role) {
    log_('failed login attempt', {});
    return { ok: false, error: 'That password was not recognized.' };
  }

  const token = Utilities.getUuid() + Utilities.getUuid() + String(Date.now());
  const now = Date.now();
  const expiresAt = now + SESSION_TTL_MS;
  const rolePassword = role === 'admin' ? passwords.admin : passwords.staff;

  sessionSheet_().appendRow([
    sha256_(token),
    role,
    sha256_(rolePassword),
    new Date(now).toISOString(),
    new Date(now).toISOString(),
    new Date(expiresAt).toISOString(),
    ''
  ]);

  log_('login succeeded', { role: role });
  return { ok: true, authenticated: true, role: role, token: token };
}

function validateSession_(token) {
  const session = findSession_(token);
  if (!session.valid) {
    return { ok: true, authenticated: false };
  }

  const passwords = readPasswords_();
  const currentPassword = session.role === 'admin' ? passwords.admin : passwords.staff;
  if (!currentPassword || sha256_(currentPassword) !== session.passwordHash) {
    revokeSessionRow_(session.rowNumber);
    log_('session invalidated by password change', { role: session.role });
    return { ok: true, authenticated: false };
  }

  const now = Date.now();
  if (Date.parse(session.expiresAt) <= now) {
    revokeSessionRow_(session.rowNumber);
    return { ok: true, authenticated: false };
  }

  sessionSheet_().getRange(session.rowNumber, 5, 1, 2).setValues([[
    new Date(now).toISOString(),
    new Date(now + SESSION_TTL_MS).toISOString()
  ]]);

  return { ok: true, authenticated: true, role: session.role };
}

function logout_(token) {
  const session = findSession_(token);
  if (session.valid) {
    revokeSessionRow_(session.rowNumber);
    log_('logout', { role: session.role });
  }
  return { ok: true, authenticated: false };
}

function updatePasswords_(request) {
  const sessionResult = validateSession_(request.token);
  if (!sessionResult.authenticated) {
    return { ok: false, error: 'Please log in again.' };
  }
  if (sessionResult.role !== 'admin') {
    log_('non-admin password update rejected', { role: sessionResult.role });
    return { ok: false, error: 'Admin access is required.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const current = readPasswords_();
    const next = {
      admin: String(request.adminPassword || '').trim() ? String(request.adminPassword) : current.admin,
      staff: String(request.staffPassword || '').trim() ? String(request.staffPassword) : current.staff
    };
    const adminChanged = next.admin !== current.admin;
    const staffChanged = next.staff !== current.staff;

    if (!adminChanged && !staffChanged) {
      return { ok: false, error: 'Enter a new admin or staff password.' };
    }

    writePasswords_(next);
    if (staffChanged) {
      revokeRoleSessions_('staff');
    }
    if (adminChanged) {
      revokeRoleSessions_('admin');
    }

    log_('passwords updated', { adminChanged: adminChanged, staffChanged: staffChanged });
    return {
      ok: true,
      success: true,
      loggedOut: adminChanged,
      message: adminChanged
        ? 'Password updated. Please log in again with the new admin password.'
        : 'Password updated.'
    };
  } finally {
    lock.releaseLock();
  }
}

function staffSession_(request, logMessage) {
  const sessionResult = validateSession_(request.token);
  if (!sessionResult.authenticated) {
    return { ok: false, error: 'Please log in again.' };
  }
  return { ok: true, role: sessionResult.role };
}

function adminSession_(request, logMessage) {
  const sessionResult = staffSession_(request, logMessage);
  if (!sessionResult.ok) {
    return sessionResult;
  }
  if (sessionResult.role !== 'admin') {
    log_(logMessage, { role: sessionResult.role });
    return { ok: false, error: 'Admin access is required.' };
  }
  return { ok: true };
}

function sourcePrayerRequests_() {
  const sheet = prayerRequestsSheet_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(function(header) {
    return normalizeHeader_(header);
  });
  const requestIndex = headers.indexOf(normalizeHeader_(PRAYER_REQUEST_TEXT_HEADER));
  const confidentialIndex = headers.indexOf(normalizeHeader_(PRAYER_REQUEST_CONFIDENTIAL_HEADER));
  const firstNameIndex = headers.indexOf(normalizeHeader_(PRAYER_REQUEST_FIRST_NAME_HEADER));

  if (requestIndex === -1 || confidentialIndex === -1) {
    throw new Error('Prayer request columns were not found.');
  }

  const prayerRequests = [];
  values.slice(1).forEach(function(row, index) {
    const rowNumber = index + 2;
    const confidentialAnswer = String(row[confidentialIndex] || '').trim();
    const requestText = String(row[requestIndex] || '').trim();
    if (isPrayerShareable_(confidentialAnswer) && requestText) {
      prayerRequests.push({
        id: prayerRequestId_(rowNumber, requestText),
        text: requestText,
        firstName: firstNameIndex === -1 ? '' : String(row[firstNameIndex] || '').trim(),
        rowNumber: rowNumber
      });
    }
  });

  return prayerRequests.reverse();
}

function isPrayerShareable_(confidentialAnswer) {
  const normalizedAnswer = normalizeHeader_(String(confidentialAnswer || ''));
  const normalizedShareValue = normalizeHeader_(PRAYER_REQUEST_SHARE_VALUE);

  if (!normalizedAnswer) {
    return false;
  }

  return normalizedAnswer === normalizedShareValue || normalizedAnswer.indexOf(PRAYER_REQUEST_SHARE_PREFIX) === 0;
}

function prayerModerationMap_() {
  const sheet = prayerModerationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, PRAYER_MODERATION_HEADERS.length).getValues();
  const moderation = {};
  rows.forEach(function(row, index) {
    const requestId = String(row[0] || '');
    if (requestId) {
      moderation[requestId] = {
        rowNumber: index + 2,
        status: String(row[1] || ''),
        text: String(row[2] || ''),
        createdAt: row[3],
        updatedAt: row[4],
        firstName: String(row[5] || '').trim()
      };
    }
  });
  return moderation;
}

function prayerRequests_(request) {
  const session = adminSession_(request, 'non-admin prayer request access rejected');
  if (!session.ok) {
    return session;
  }

  const moderation = prayerModerationMap_();
  const prayerRequests = sourcePrayerRequests_().filter(function(prayerRequest) {
    return !moderation[prayerRequest.id];
  });

  return { ok: true, prayerRequests: prayerRequests };
}

function moderatePrayerRequest_(request) {
  const session = adminSession_(request, 'non-admin prayer moderation rejected');
  if (!session.ok) {
    return session;
  }

  const requestId = String(request.requestId || '').trim();
  const status = String(request.status || '').trim().toLowerCase();
  if (!requestId) {
    return { ok: false, error: 'Prayer request ID is required.' };
  }
  if (status !== PRAYER_STATUS_APPROVED && status !== PRAYER_STATUS_DENIED) {
    return { ok: false, error: 'Choose approve or deny.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const prayerRequest = sourcePrayerRequests_().filter(function(item) {
      return item.id === requestId;
    })[0];
    if (!prayerRequest) {
      return { ok: false, error: 'Prayer request was not found.' };
    }

    const sheet = prayerModerationSheet_();
    const moderation = prayerModerationMap_();
    const now = new Date().toISOString();
    if (moderation[requestId]) {
      sheet.getRange(moderation[requestId].rowNumber, 2, 1, 5).setValues([[
        status,
        prayerRequest.text,
        moderation[requestId].createdAt || now,
        now,
        prayerRequest.firstName
      ]]);
    } else {
      sheet.appendRow([requestId, status, prayerRequest.text, now, now, prayerRequest.firstName]);
    }

    if (status === PRAYER_STATUS_APPROVED) {
      deleteSourcePrayerRequest_(prayerRequest);
    }
  } finally {
    lock.releaseLock();
  }

  log_('prayer request moderated', { requestId: requestId, status: status });
  return { ok: true, success: true, status: status };
}

function publicPrayerRequests_() {
  const prayerRequests = approvedPrayerRequests_();

  return { ok: true, prayerRequests: prayerRequests };
}

function removePublicPrayerRequest_(request) {
  const session = staffSession_(request, 'non-staff public prayer removal rejected');
  if (!session.ok) {
    return session;
  }

  const requestId = String(request.requestId || '').trim();
  if (!requestId) {
    return { ok: false, error: 'Prayer request ID is required.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = prayerModerationSheet_();
    const moderation = prayerModerationMap_();
    const prayerRequest = moderation[requestId];
    if (!prayerRequest || prayerRequest.status !== PRAYER_STATUS_APPROVED) {
      return { ok: false, error: 'Approved prayer request was not found.' };
    }

    sheet.getRange(prayerRequest.rowNumber, 2, 1, 4).setValues([[
      PRAYER_STATUS_DENIED,
      prayerRequest.text,
      prayerRequest.createdAt || new Date().toISOString(),
      new Date().toISOString()
    ]]);
  } finally {
    lock.releaseLock();
  }

  log_('public prayer request removed', { requestId: requestId, role: session.role });
  return { ok: true, success: true };
}

function approvedPrayerRequests_() {
  const sheet = prayerModerationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, PRAYER_MODERATION_HEADERS.length).getValues();
  const prayerRequests = [];
  rows.forEach(function(row) {
    const requestId = String(row[0] || '');
    const status = String(row[1] || '');
    const requestText = String(row[2] || '').trim();
    if (requestId && status === PRAYER_STATUS_APPROVED && requestText) {
      prayerRequests.push({
        id: requestId,
        text: requestText,
        firstName: String(row[5] || '').trim(),
        approvedAt: row[4] || row[3] || ''
      });
    }
  });

  prayerRequests.sort(function(left, right) {
    return Date.parse(right.approvedAt) - Date.parse(left.approvedAt);
  });
  return prayerRequests;
}

function prayerRequestId_(rowNumber, requestText) {
  return sha256_(String(rowNumber) + '\n' + String(requestText || ''));
}

function deleteSourcePrayerRequest_(prayerRequest) {
  const rowNumber = Number(prayerRequest && prayerRequest.rowNumber);
  if (rowNumber >= 2) {
    prayerRequestsSheet_().deleteRow(rowNumber);
  }
}

function normalizeHeader_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function liveStatus_() {
  const properties = PropertiesService.getScriptProperties();
  const apiKey = properties.getProperty('YOUTUBE_API_KEY');
  const channelId = properties.getProperty('YOUTUBE_CHANNEL_ID') || DEFAULT_YOUTUBE_CHANNEL_ID;

  if (!apiKey) {
    return { ok: true, live: false };
  }

  const url = 'https://www.googleapis.com/youtube/v3/search?'
    + 'part=snippet'
    + '&channelId=' + encodeURIComponent(channelId)
    + '&eventType=live'
    + '&type=video'
    + '&key=' + encodeURIComponent(apiKey);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    log_('youtube live status failed', { status: response.getResponseCode() });
    return { ok: true, live: false };
  }
  const data = JSON.parse(response.getContentText());
  const item = data.items && data.items[0];
  return {
    ok: true,
    live: Boolean(item),
    videoId: item && item.id && item.id.videoId ? item.id.videoId : null
  };
}

function findSession_(token) {
  if (!token) {
    return { valid: false };
  }

  const tokenHash = sha256_(String(token));
  const sheet = sessionSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { valid: false };
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (let index = rows.length - 1; index >= 0; index--) {
    const row = rows[index];
    if (row[0] === tokenHash && !row[6]) {
      return {
        valid: true,
        rowNumber: index + 2,
        role: row[1],
        passwordHash: row[2],
        createdAt: row[3],
        lastSeenAt: row[4],
        expiresAt: row[5]
      };
    }
  }

  return { valid: false };
}

function revokeSessionRow_(rowNumber) {
  sessionSheet_().getRange(rowNumber, 7).setValue(new Date().toISOString());
}

function revokeRoleSessions_(role) {
  const sheet = sessionSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const revokedAt = new Date().toISOString();
  rows.forEach(function(row, index) {
    if (row[1] === role && !row[6]) {
      sheet.getRange(index + 2, 7).setValue(revokedAt);
    }
  });
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(function(byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function log_(message, details) {
  console.log(JSON.stringify({
    feature: 'staff-auth',
    message: message,
    details: details || {},
    at: new Date().toISOString()
  }));
}
