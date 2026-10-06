const REQUEST_HEADERS = [
  'Received',
  'Name',
  'Email',
  'Project type',
  'Message',
  'Status',
  'Last reply'
];

function doGet(e) {
  return page_('The portfolio admin API is ready. Open the portfolio and use the footer admin shortcut.');
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};
  switch (params.action) {
    case 'submit':
      return receiveRequest_(params);
    case 'admin-api-login':
    case 'admin-api-list':
    case 'admin-api-reply':
    case 'admin-api-logout':
      return adminApi_(params);
    default:
      return page_('Unable to process this request. Please return to the portfolio and try again.');
  }
}

function receiveRequest_(params) {
  const name = clean_(params.name, 120);
  const email = clean_(params.email, 254).toLowerCase();
  const projectType = clean_(params.projectType, 80);
  const message = clean_(params.message, 5000);
  const allowedTypes = ['Website', 'E-commerce', 'UI / UX design', 'Graphic design', 'Something else'];

  if (params.website) return page_('Your request could not be accepted.');
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !message || !allowedTypes.includes(projectType)) {
    return page_('Please return to the portfolio and enter a valid name, email, project type, and message.');
  }

  const sheet = requestSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let row;
  try {
    row = sheet.getLastRow() + 1;
    sheet.appendRow([
      new Date(),
      sheetText_(name),
      sheetText_(email),
      sheetText_(projectType),
      sheetText_(message),
      'New',
      ''
    ]);
  } finally {
    lock.releaseLock();
  }

  let emailSent = true;
  try {
    GmailApp.sendEmail(
      email,
      'We received your project enquiry',
      'Hi ' + name + ',\n\nThank you for reaching out to Mujtaba Abdul Mannan. Your project enquiry has been received. I will review the details and get back to you soon.\n\nProject type: ' + projectType + '\n\nBest,\nMujtaba Abdul Mannan',
      { name: 'Mujtaba Abdul Mannan' }
    );
  } catch (error) {
    emailSent = false;
    console.error('The request was saved, but its confirmation email could not be sent: ' + error);
    sheet.getRange(row, 6).setValue('New — confirmation email failed');
  }

  return receipt_(emailSent);
}

function adminApi_(params) {
  const requestId = String(params.requestId || '');
  const nonce = String(params.nonce || '');
  let result;
  switch (params.action) {
    case 'admin-api-login':
      result = adminLogin_(params);
      break;
    case 'admin-api-list':
      result = adminList_(params);
      break;
    case 'admin-api-reply':
      result = adminReply_(params);
      break;
    case 'admin-api-logout':
      result = adminLogout_(params);
      break;
    default:
      result = { ok: false, error: 'Unsupported admin request.' };
  }
  return bridgeResponse_(requestId, nonce, result);
}

function adminLogin_(params) {
  const properties = PropertiesService.getScriptProperties();
  const expectedEmail = (properties.getProperty('ADMIN_EMAIL') || '').trim().toLowerCase();
  const expectedPassword = properties.getProperty('ADMIN_PASSWORD') || '';
  const parentOrigin = properties.getProperty('PORTFOLIO_ORIGIN') || '';
  if (!expectedEmail || !expectedPassword) return { ok: false, error: 'Admin credentials are not configured in Apps Script properties.' };
  if (!parentOrigin || !/^https?:\/\/[^/]+$/.test(parentOrigin)) return { ok: false, error: 'Set PORTFOLIO_ORIGIN in Apps Script properties to your published portfolio origin, for example https://your-domain.com.' };

  const email = clean_(params.username, 254).toLowerCase();
  const password = String(params.password || '');
  const cache = CacheService.getScriptCache();
  const attemptKey = 'login-attempts:' + expectedEmail;
  const attempts = Number(cache.get(attemptKey) || 0);
  if (attempts >= 5) return { ok: false, error: 'Too many attempts. Wait 10 minutes before trying again.' };

  if (email !== expectedEmail || password !== expectedPassword) {
    cache.put(attemptKey, String(attempts + 1), 600);
    return { ok: false, error: 'Those credentials were not accepted.' };
  }

  cache.remove(attemptKey);
  const token = Utilities.getUuid();
  cache.put('admin-session:' + token, expectedEmail, 3600);
  return { ok: true, token: token };
}

function adminList_(params) {
  const token = String(params.token || '');
  if (!isAdminSession_(token)) return { ok: false, expired: true, error: 'Your admin session expired. Please sign in again.' };
  try {
    return { ok: true, requests: getRequests_() };
  } catch (error) {
    console.error('Unable to load portfolio contact requests: ' + error);
    return { ok: false, error: 'Could not load requests from Google Sheets. Check the Apps Script Sheets permissions.' };
  }
}

function adminReply_(params) {
  const token = String(params.token || '');
  if (!isAdminSession_(token)) return { ok: false, expired: true, error: 'Your admin session expired. Please sign in again.' };

  const row = Number(params.row);
  const subject = clean_(params.subject, 200);
  const message = clean_(params.reply, 5000);
  let sheet;
  try {
    sheet = requestSheet_();
  } catch (error) {
    console.error('Unable to access the contact request sheet: ' + error);
    return { ok: false, error: 'Could not access Google Sheets. Check the Apps Script Sheets permissions.' };
  }
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow() || !subject || !message) {
    return { ok: false, error: 'Enter a valid request, subject, and message.' };
  }

  const recipient = String(sheet.getRange(row, 3).getValue()).trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
    return { ok: false, error: 'This request has no valid client email address.' };
  }

  try {
    GmailApp.sendEmail(recipient, subject, message, { name: 'Mujtaba Abdul Mannan' });
  } catch (error) {
    console.error('Admin reply email failed for sheet row ' + row + ': ' + error);
    return { ok: false, error: 'The email could not be sent. Check Gmail authorization and quota, then try again.' };
  }

  try {
    sheet.getRange(row, 6).setValue('Replied');
    sheet.getRange(row, 7).setValue(new Date());
  } catch (error) {
    console.error('Reply sent, but the sheet status could not be updated for row ' + row + ': ' + error);
    return { ok: true, warning: 'Email sent to ' + recipient + ', but its sheet status could not be updated.' };
  }
  return { ok: true, message: 'Reply sent to ' + recipient + '.' };
}

function adminLogout_(params) {
  const token = String(params.token || '');
  if (token) CacheService.getScriptCache().remove('admin-session:' + token);
  return { ok: true };
}

function isAdminSession_(token) {
  return Boolean(token && CacheService.getScriptCache().get('admin-session:' + token));
}

function requestSheet_() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty('SHEET_ID');
  let spreadsheet;
  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create('Mujtaba Portfolio Contact Requests');
    properties.setProperty('SHEET_ID', spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName('Requests');
  if (!sheet) sheet = spreadsheet.insertSheet('Requests');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REQUEST_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function bridgeResponse_(requestId, nonce, result) {
  if (!/^[a-zA-Z0-9-]{16,100}$/.test(requestId) || !/^[a-zA-Z0-9-]{16,100}$/.test(nonce)) {
    return HtmlService.createHtmlOutput('Invalid secure request.').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }
  const parentOrigin = PropertiesService.getScriptProperties().getProperty('PORTFOLIO_ORIGIN') || '';
  const payload = safeJson_({
    channel: 'portfolio-admin',
    requestId: requestId,
    nonce: nonce,
    result: result
  });
  const target = safeJson_(parentOrigin);
  const html = '<!doctype html><meta charset="utf-8"><script>(function(){' +
    'var payload=' + payload + ';var target=' + target + ';' +
    'if(target&&parent!==self&&document.referrer){' +
    'try{if(new URL(document.referrer).origin===target)parent.postMessage(payload,target)}catch(e){}' +
    '}})();</script>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Secure portfolio connection')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getRequests_() {
  const sheet = requestSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, REQUEST_HEADERS.length).getValues();
  return rows.map(function(row, index) {
    return {
      row: index + 2,
      received: row[0] instanceof Date
        ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm')
        : String(row[0] || ''),
      name: String(row[1] || ''),
      email: String(row[2] || ''),
      projectType: String(row[3] || ''),
      message: String(row[4] || ''),
      status: String(row[5] || 'New'),
      lastReply: row[6] instanceof Date
        ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm')
        : String(row[6] || '')
    };
  }).reverse();
}

function safeJson_(value) {
  return JSON.stringify(value).replace(/[<>&\u2028\u2029]/g, function(character) {
    return '\\u' + character.charCodeAt(0).toString(16).padStart(4, '0');
  });
}

function clean_(value, maxLength) {
  return String(value || '').trim().slice(0, maxLength);
}

function sheetText_(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function receipt_(emailSent) {
  const message = emailSent
    ? 'Your enquiry is safely with us. A confirmation has been sent to your email, and Mujtaba will be in touch soon.'
    : 'Your enquiry has been saved, but the confirmation email could not be sent. Mujtaba can still review your request.';
  return page_(message);
}

function page_(message) {
  const safeMessage = String(message).replace(/[&<>"]/g, function(character) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[character];
  });
  return HtmlService.createHtmlOutput(
    '<!doctype html><html lang="en"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Request received</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#080807;color:#f5f3ee;font:16px/1.6 Arial,sans-serif}.card{max-width:560px;margin:24px;padding:clamp(28px,7vw,64px);border:1px solid #34302c;background:#11100f}p{font-size:1.2rem}a{color:#ff6845}</style>' +
    '<main class="card"><p>' + safeMessage + '</p><p>You can close this tab and return to the portfolio.</p></main></html>'
  ).setTitle('Request received — Mujtaba Abdul Mannan');
}
