/**
 * Sistah-Friends Google Sheets form receiver
 *
 * 1) Create a Google Sheet.
 * 2) Extensions > Apps Script.
 * 3) Paste this file in.
 * 4) Run setupSistahFriendsDatabase() once.
 * 5) Deploy as Web App: Execute as Me; access Anyone.
 * 6) Paste the /exec URL into SISTAH_FRIENDS_GOOGLE_SCRIPT_URL in sistahfriends.html.
 */
const NOTIFICATION_EMAIL = 'designerdestinationsandevents@gmail.com';
const DATABASE_PROPERTY = 'SISTAH_FRIENDS_SPREADSHEET_ID';

function setupSistahFriendsDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the Sistah-Friends Google Sheet, then run setup again.');

  PropertiesService.getScriptProperties().setProperty(DATABASE_PROPERTY, ss.getId());

  ['Members', 'Referrals', 'RSVPs', 'Suggestions', 'Other Submissions'].forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.setFrozenRows(1);
  });

  SpreadsheetApp.getUi().alert('Sistah-Friends database setup is complete. Next, deploy this Apps Script as a Web App.');
}

function doGet(e) {
  const view = e && e.parameter ? String(e.parameter.view || '') : '';

  if (view === 'community') {
    return renderCommunitySpotlight_();
  }

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><body style="font-family:Arial,sans-serif;padding:28px">' +
      '<h2>Sistah-Friends form service is active.</h2>' +
      '<p>This endpoint receives website form submissions.</p>' +
      '</body></html>'
    )
    .setTitle('Sistah-Friends');
}

function doPost(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const ssId = PropertiesService.getScriptProperties().getProperty(DATABASE_PROPERTY);
    if (!ssId) throw new Error('Database is not configured. Run setupSistahFriendsDatabase() first.');

    const ss = SpreadsheetApp.openById(ssId);
    const subject = params._subject || 'Sistah-Friends Website Submission';
    const tabName = chooseTab_(subject);
    const sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);

    const data = cleanData_(params);
    data['Submitted At'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    data['Source'] = 'Sistah-Friends Website';

    appendFlexibleRow_(sheet, data);
    sendNotification_(subject, data);

    return redirectPage_(sanitizeRedirect_(params._next));
  } catch (err) {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: 'Sistah-Friends form error',
      htmlBody: '<p>A website submission could not be saved.</p><p><strong>Error:</strong> ' +
        escapeHtml_(String(err)) + '</p>'
    });
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><body style="font-family:Arial;padding:30px">' +
      '<h2>We received an error while saving your submission.</h2>' +
      '<p>Please email ' + NOTIFICATION_EMAIL + ' so we can assist.</p></body></html>'
    );
  }
}

function chooseTab_(subject) {
  const s = String(subject).toLowerCase();
  if (s.includes('member profile')) return 'Members';
  if (s.includes('referral')) return 'Referrals';
  if (s.includes('rsvp')) return 'RSVPs';
  if (s.includes('suggestion')) return 'Suggestions';
  return 'Other Submissions';
}

function cleanData_(params) {
  const out = {};
  Object.keys(params).forEach(key => {
    if (key.startsWith('_')) return;
    out[key] = params[key];
  });
  return out;
}

function appendFlexibleRow_(sheet, data) {
  const lastCol = Math.max(sheet.getLastColumn(), 0);
  let headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String) : [];
  const newHeaders = Object.keys(data).filter(key => !headers.includes(key));

  if (newHeaders.length) {
    headers = headers.concat(newHeaders);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#44006f')
      .setFontColor('#ffffff');
  }

  sheet.appendRow(headers.map(h => data[h] !== undefined ? data[h] : ''));
  sheet.autoResizeColumns(1, headers.length);
}

function sendNotification_(subject, data) {
  const rows = Object.keys(data).map(key =>
    '<tr><td style="padding:6px 10px;border:1px solid #ddd;font-weight:bold">' +
    escapeHtml_(key) + '</td><td style="padding:6px 10px;border:1px solid #ddd">' +
    escapeHtml_(String(data[key] || '')) + '</td></tr>'
  ).join('');

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    htmlBody:
      '<div style="font-family:Arial,sans-serif">' +
      '<h2 style="color:#44006f">New Sistah-Friends Website Submission</h2>' +
      '<table style="border-collapse:collapse">' + rows + '</table>' +
      '<p style="color:#666">This submission was also saved to the Sistah-Friends Google Sheet.</p>' +
      '</div>'
  });
}

function sanitizeRedirect_(url) {
  const fallback = 'https://www.optimaxproai.com/thank-you.html';
  if (!url) return fallback;
  const value = String(url);
  if (/^https:\/\/(www\.)?optimaxproai\.com\//i.test(value)) return value;
  return fallback;
}

function redirectPage_(url) {
  const safe = escapeHtml_(url);
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8">' +
    '<meta http-equiv="refresh" content="0;url=' + safe + '">' +
    '</head><body><p>Thank you. Redirecting…</p>' +
    '<script>window.top.location.href=' + JSON.stringify(url) + ';</script>' +
    '</body></html>'
  );
}


function renderCommunitySpotlight_() {
  const ssId = PropertiesService.getScriptProperties().getProperty(DATABASE_PROPERTY);
  if (!ssId) throw new Error('Database is not configured. Run setupSistahFriendsDatabase() first.');

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('Members');

  if (!sheet || sheet.getLastRow() < 2) {
    return HtmlService.createHtmlOutput(communityHtml_([]))
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();

  const idx = {};
  headers.forEach((h, i) => idx[String(h).trim()] = i);

  const rows = values.map(row => ({
    firstName: getCell_(row, idx, 'First Name'),
    lastName: getCell_(row, idx, 'Last Name'),
    spotlightName: getCell_(row, idx, 'Spotlight Name'),
    spotlightUrl: getCell_(row, idx, 'Spotlight URL'),
    permission: getCell_(row, idx, 'Show in Sistah-Friends Spotlight')
  }))
  .filter(item =>
    item.spotlightName &&
    item.spotlightUrl &&
    String(item.permission).toLowerCase() === 'yes'
  );

  // If an existing member updates her record later, the newest entry appears first.
  rows.reverse();

  return HtmlService
    .createHtmlOutput(communityHtml_(rows))
    .setTitle('Sistah-Friends Spotlight')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCell_(row, idx, header) {
  return idx[header] !== undefined ? String(row[idx[header]] || '').trim() : '';
}

function safePublicUrl_(value) {
  const url = String(value || '').trim();
  if (!/^https?:\/\//i.test(url)) return '';
  return url;
}

function communityHtml_(items) {
  const safeItems = items
    .map(item => {
      const url = safePublicUrl_(item.spotlightUrl);
      if (!url) return '';

      const displayName = escapeHtml_(item.spotlightName);
      const memberName = escapeHtml_(
        [item.firstName, item.lastName].filter(Boolean).join(' ')
      );

      return '<a class="spotlight-row" href="' + escapeHtml_(url) + '" target="_blank" rel="noopener">' +
        '<span class="spotlight-name">' + displayName + '</span>' +
        (memberName ? '<span class="member-name">' + memberName + '</span>' : '') +
        '<span class="visit">Visit ↗</span>' +
      '</a>';
    })
    .filter(Boolean)
    .join('');

  const body = safeItems || (
    '<div class="empty">' +
      '<strong>Our Sistah-Friends Spotlight is growing.</strong>' +
      '<span>Sistahs can add a business, LinkedIn profile, app, project or other link when they complete their member profile.</span>' +
    '</div>'
  );

  return '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
      '*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;color:#28162d;background:#fff}' +
      '.list{padding:4px 18px}' +
      '.spotlight-row{display:grid;grid-template-columns:minmax(0,1.7fr) minmax(0,1fr) auto;gap:14px;align-items:center;padding:14px 4px;border-bottom:1px solid #eadfe9;text-decoration:none;color:inherit}' +
      '.spotlight-row:last-child{border-bottom:0}' +
      '.spotlight-row:hover .spotlight-name{color:#bc0d96}' +
      '.spotlight-name{font-weight:800;color:#520074}' +
      '.member-name{font-size:13px;color:#6e6072}' +
      '.visit{font-size:12px;font-weight:800;color:#e31379;white-space:nowrap}' +
      '.empty{padding:34px 12px;text-align:center;color:#6e6072;display:grid;gap:7px}' +
      '.empty strong{color:#520074}' +
      '@media(max-width:600px){.spotlight-row{grid-template-columns:1fr auto}.member-name{grid-column:1}.visit{grid-column:2;grid-row:1 / span 2}}' +
    '</style></head><body><div class="list">' + body + '</div></body></html>';
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
