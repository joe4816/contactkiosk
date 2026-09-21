/**
 * ContactKiosk Apps Script backend
 * GitHub front end: https://joe4816.github.io/contactkiosk/
 *
 * Deploy as a Web App:
 *   Execute as: Me (deploying user)
 *   Who has access: users in the CCSD domain
 *
 * The GitHub front end uses signed session tokens + JSONP so Safari can
 * communicate cross-origin without exposing staff PINs in the public repo.
 */

const CONTACT_KIOSK_ID = '1hJ2pG7FZ2tn2vKyk3bOZm_iFlgxOGprIC9Roxyp4_-c';
const FRONTEND_URL = 'https://joe4816.github.io/contactkiosk/';
const MASTER_BULK_PIN = '0374';
const SESSION_HOURS = 12;
const BULK_HOURS = 12;

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').trim();

  try {
    if (action === 'auth') return handleAuth_();
    if (action === 'config') return jsonp_(e, getConfigResponse_(e));
    if (action === 'verifyBulk') return jsonp_(e, verifyBulkResponse_(e));
    if (action === 'record') return jsonp_(e, recordResponse_(e));
    if (action === 'health') return jsonp_(e, { ok: true, service: 'ContactKiosk', version: 1 });

    return jsonp_(e, {
      ok: false,
      error: 'Unknown action.'
    });
  } catch (err) {
    return jsonp_(e, {
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

function handleAuth_() {
  const email = getCurrentStaffEmail_();

  if (!email) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font-family:Arial;padding:32px">Unable to identify your CCSD Google account.</div>'
    );
  }

  const staff = getStaffRecord_(email);
  if (!staff.authorized) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<div style="font-family:Arial;padding:32px">This account is not authorized for ContactKiosk.<br><br>' +
      escapeHtml_(email) + '</div>'
    );
  }

  const token = makeToken_({
    sub: email,
    scope: 'staff',
    exp: Date.now() + SESSION_HOURS * 60 * 60 * 1000
  });

  const redirect = FRONTEND_URL + '?session=' + encodeURIComponent(token);

  return HtmlService.createHtmlOutput(
    '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>ContactKiosk</title>' +
    '<body style="font-family:Arial;text-align:center;padding:40px">Connecting ContactKiosk…' +
    '<script>location.replace(' + JSON.stringify(redirect) + ');<\/script></body>'
  );
}

function getConfigResponse_(e) {
  const session = requireStaffSession_(e);
  const email = session.sub;
  const staff = getStaffRecord_(email);

  if (!staff.authorized) throw new Error('Staff access is no longer active.');

  const config = getConfigForStaff_(email);

  return {
    ok: true,
    staff: {
      email: email,
      displayName: staff.displayName || email
    },
    buttons: config.buttons,
    hasBulkMode: !!config.lockedReason,
    bulkReason: config.lockedReason || ''
  };
}

function verifyBulkResponse_(e) {
  const session = requireStaffSession_(e);
  const attempt = String(e.parameter.pin || '').trim();
  const config = getConfigForStaff_(session.sub);

  const valid =
    attempt === MASTER_BULK_PIN ||
    (config.personalPin && attempt === config.personalPin);

  if (!valid) {
    return { ok: false, error: 'Incorrect PIN.' };
  }

  if (!config.lockedReason) {
    return { ok: false, error: 'No bulk reason is configured for this staff member.' };
  }

  const bulkToken = makeToken_({
    sub: session.sub,
    scope: 'bulk',
    exp: Date.now() + BULK_HOURS * 60 * 60 * 1000
  });

  return {
    ok: true,
    bulkToken: bulkToken,
    bulkReason: config.lockedReason
  };
}

function recordResponse_(e) {
  const session = requireStaffSession_(e);
  const email = session.sub;

  const qrValue = String(e.parameter.qr || '').trim();
  const requestedService = String(e.parameter.service || '').trim();
  const isBulk = String(e.parameter.bulk || '') === '1';

  if (!qrValue) throw new Error('Student QR is required.');

  const config = getConfigForStaff_(email);
  let service = requestedService;
  let source = 'Kiosk';

  if (isBulk) {
    const bulkToken = String(e.parameter.bulkToken || '').trim();
    const bulk = verifyToken_(bulkToken);

    if (!bulk || bulk.scope !== 'bulk' || bulk.sub !== email) {
      throw new Error('Bulk authorization expired. Re-enter the bulk PIN.');
    }

    if (!config.lockedReason) throw new Error('No bulk reason is configured.');
    service = config.lockedReason;
    source = 'Bulk';
  } else {
    const allowed = config.buttons.filter(Boolean);
    if (!allowed.includes(service)) {
      throw new Error('That visit reason is not currently configured.');
    }
  }

  const ss = SpreadsheetApp.openById(CONTACT_KIOSK_ID);
  const log = ss.getSheetByName('Log');
  if (!log) throw new Error('Log sheet not found.');

  const eventId = Utilities.getUuid();
  const timestamp = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    log.appendRow([
      eventId,
      timestamp,
      email,
      qrValue,
      service,
      source
    ]);
  } finally {
    lock.releaseLock();
  }

  return {
    ok: true,
    eventId: eventId,
    service: service,
    source: source
  };
}

function getCurrentStaffEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function getStaffRecord_(email) {
  const ss = SpreadsheetApp.openById(CONTACT_KIOSK_ID);
  const sheet = ss.getSheetByName('Access');
  if (!sheet) throw new Error('Access sheet not found.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return { authorized: false };

  const headers = values[0].map(String);
  const idx = {};
  headers.forEach((h, i) => idx[String(h).trim()] = i);

  const emailCol = idx['Email'];
  if (emailCol == null) throw new Error('Access sheet is missing Email.');

  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][emailCol] || '').trim().toLowerCase();
    if (rowEmail !== email) continue;

    const shareStatus = idx['Share Status'] == null ? '' : String(values[r][idx['Share Status']] || '').trim();
    const appAccess = idx['App Access'] == null ? '' : String(values[r][idx['App Access']] || '').trim();

    const active =
      shareStatus.toLowerCase() !== 'removed' &&
      appAccess.toLowerCase() !== 'no access';

    return {
      authorized: active,
      displayName: idx['Display Name'] == null ? '' : String(values[r][idx['Display Name']] || '').trim(),
      appAccess: appAccess,
      shareStatus: shareStatus
    };
  }

  return { authorized: false };
}

function getConfigForStaff_(email) {
  const ss = SpreadsheetApp.openById(CONTACT_KIOSK_ID);
  const sheet = ss.getSheetByName('Config');
  if (!sheet) throw new Error('Config sheet not found.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('Config sheet has no configuration rows.');

  const headers = values[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  let defaultRow = null;
  let staffRow = null;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rowEmail = String(row[idx['StaffEmail']] || '').trim().toLowerCase();

    if (rowEmail === 'default') defaultRow = row;
    if (rowEmail === email) staffRow = row;
  }

  const base = defaultRow || [];
  const own = staffRow || [];

  const read = (columnName) => {
    const i = idx[columnName];
    if (i == null) return '';
    const ownValue = String(own[i] || '').trim();
    if (ownValue !== '') return ownValue;
    return String(base[i] || '').trim();
  };

  const buttons = [];
  for (let n = 1; n <= 8; n++) {
    const v = read('Button' + n);
    if (v) buttons.push(v);
  }

  return {
    personalPin: read('Pin'),
    buttons: buttons,
    lockedReason: read('LockedReason')
  };
}

function requireStaffSession_(e) {
  const token = String((e && e.parameter && e.parameter.session) || '').trim();
  const payload = verifyToken_(token);

  if (!payload || payload.scope !== 'staff' || !payload.sub) {
    throw new Error('Staff session is missing or expired.');
  }

  return payload;
}

function makeToken_(payload) {
  const body = base64UrlEncode_(JSON.stringify(payload));
  const signature = sign_(body);
  return body + '.' + signature;
}

function verifyToken_(token) {
  if (!token || token.indexOf('.') < 0) return null;

  const parts = token.split('.');
  if (parts.length !== 2) return null;

  const body = parts[0];
  const suppliedSig = parts[1];
  const expectedSig = sign_(body);

  if (!constantTimeEqual_(suppliedSig, expectedSig)) return null;

  let payload;
  try {
    payload = JSON.parse(base64UrlDecode_(body));
  } catch (err) {
    return null;
  }

  if (!payload.exp || Number(payload.exp) < Date.now()) return null;
  return payload;
}

function sign_(value) {
  const secret = getSigningSecret_();
  const bytes = Utilities.computeHmacSha256Signature(
    value,
    secret,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function getSigningSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('CONTACT_KIOSK_SIGNING_SECRET');

  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('CONTACT_KIOSK_SIGNING_SECRET', secret);
  }

  return secret;
}

function base64UrlEncode_(text) {
  return Utilities.base64EncodeWebSafe(
    Utilities.newBlob(text).getBytes()
  ).replace(/=+$/g, '');
}

function base64UrlDecode_(encoded) {
  const bytes = Utilities.base64DecodeWebSafe(encoded);
  return Utilities.newBlob(bytes).getDataAsString();
}

function constantTimeEqual_(a, b) {
  a = String(a || '');
  b = String(b || '');

  if (a.length !== b.length) return false;

  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function jsonp_(e, data) {
  const callback = String((e && e.parameter && e.parameter.callback) || 'contactKioskCallback');

  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    throw new Error('Invalid callback name.');
  }

  const json = JSON.stringify(data).replace(/</g, '\\u003c');
  return ContentService
    .createTextOutput(callback + '(' + json + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
