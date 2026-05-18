/**
 * Centered in Christ staff/admin backend for Google Apps Script.
 *
 * Deploy this Apps Script project as a Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * SHEET_ID should already be saved as a Script Property named SHEET_ID.
 * If you used a const in your test code instead, replace the fallback string below.
 */
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || 'PASTE_YOUR_EXISTING_SHEET_ID_HERE';
const PASSWORD_SHEET_NAME = 'Password';
const PASSWORD_RANGE = 'A1:B2';
const PASSWORD_UPDATE_RANGE = 'A2:B2';
const SESSION_SHEET_NAME = 'StaffSessions';
const SESSION_TTL_MS = 12 * 60 * 60 * 1000;
const DEFAULT_YOUTUBE_CHANNEL_ID = 'UCubyomJfZo_C11A5fxRWnGw';

function doPost(e) {
  try {
    const request = parseRequest_(e);
    const action = String(request.action || '');

    if (action === 'ping') {
      return json_(ping_());
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
    if (action === 'liveStatus') {
      return json_(liveStatus_());
    }

    return json_({ ok: false, error: 'Unknown action.' });
  } catch (error) {
    log_('request failed', { error: error.message });
    return json_({ ok: false, error: 'Authentication service is unavailable.' });
  }
}

function doGet(e) {
  const request = e && e.parameter ? e.parameter : {};
  if (request.action === 'ping') {
    return json_(ping_());
  }
  return json_({ ok: true, service: 'centeredinchrist-staff-auth' });
}

function parseRequest_(e) {
  if (!e) {
    return {};
  }

  const content = e.postData && e.postData.contents ? e.postData.contents : '';
  const type = e.postData && e.postData.type ? e.postData.type : '';

  if (content) {
    if (type.indexOf('application/json') !== -1 || type.indexOf('text/plain') !== -1) {
      return JSON.parse(content);
    }
    if (type.indexOf('application/x-www-form-urlencoded') !== -1) {
      return parseFormBody_(content);
    }
  }

  return e.parameter || {};
}

function parseFormBody_(content) {
  return content.split('&').reduce(function(result, pair) {
    if (!pair) {
      return result;
    }
    const parts = pair.split('=');
    const key = decodeFormValue_(parts.shift() || '');
    const value = decodeFormValue_(parts.join('='));
    result[key] = value;
    return result;
  }, {});
}

function decodeFormValue_(value) {
  return decodeURIComponent(String(value || '').replace(/\+/g, ' '));
}

function ping_() {
  return {
    ok: true,
    service: 'centeredinchrist-staff-auth',
    action: 'ping',
    timestamp: new Date().toISOString()
  };
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function spreadsheet_() {
  if (!SHEET_ID || SHEET_ID === 'PASTE_YOUR_EXISTING_SHEET_ID_HERE') {
    throw new Error('SHEET_ID is not configured.');
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

function passwordSheet_() {
  const sheet = spreadsheet_().getSheetByName(PASSWORD_SHEET_NAME);
  if (!sheet) {
    throw new Error('Password sheet was not found.');
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
