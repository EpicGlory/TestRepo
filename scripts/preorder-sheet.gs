// Cosmic Pig BBQ — All-forms Google Apps Script
// Handles: pre-orders, email signups, and contact messages
// Each writes to its own sheet tab and emails a notification.
//
// Deploy as a Web App:
//   Extensions > Apps Script > Deploy > New deployment
//   Type: Web app | Execute as: Me | Who has access: Anyone
// Copy the Web App URL into js/main.js (FORMS_SCRIPT_URL constant).

var NOTIFY_EMAIL = 'ryan@cosmicpigbbq.com';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var type = data.form_type || 'preorder';

    if (type === 'signup') {
      handleSignup(data);
    } else if (type === 'contact') {
      handleContact(data);
    } else {
      handlePreorder(data);
    }

    return ContentService
      .createTextOutput('OK')
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService
      .createTextOutput('Error: ' + err.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ---- helpers ----

function getOrCreateSheet(name, headers) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function ts() {
  return Utilities.formatDate(
    new Date(), 'America/Denver', 'MM/dd/yyyy HH:mm:ss'
  );
}

// ---- handlers ----

function handlePreorder(data) {
  var sheet = getOrCreateSheet('Pre-Orders', [
    'Timestamp', 'First Name', 'Last Name', 'Email', 'Phone',
    'City', 'State', 'Delivery Preference', 'Rubs', 'Quantity', 'Source', 'Notes'
  ]);
  sheet.appendRow([
    ts(),
    data.first_name          || '',
    data.last_name           || '',
    data.email               || '',
    data.phone               || '',
    data.city                || '',
    data.state               || '',
    data.delivery_preference || '',
    data.rubs                || '',
    data.quantity            || '',
    data.source              || '',
    data.notes               || ''
  ]);
  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to:      NOTIFY_EMAIL,
      subject: 'New Pre-Order: ' + (data.first_name || '') + ' ' + (data.last_name || ''),
      body:
        'Name:     ' + (data.first_name || '') + ' ' + (data.last_name || '') + '\n' +
        'Email:    ' + (data.email || '') + '\n' +
        'Phone:    ' + (data.phone || 'not provided') + '\n' +
        'City:     ' + (data.city || 'not provided') + '\n' +
        'State:    ' + (data.state || 'not provided') + '\n' +
        'Rubs:     ' + (data.rubs || 'none selected') + '\n' +
        'Quantity: ' + (data.quantity || 'not specified') + '\n' +
        'Delivery: ' + (data.delivery_preference || 'not specified') + '\n' +
        'Source:   ' + (data.source || '') + '\n' +
        'Notes:    ' + (data.notes || '')
    });
  }
}

function handleSignup(data) {
  var sheet = getOrCreateSheet('Email Signups', [
    'Timestamp', 'First Name', 'Email', 'Phone'
  ]);
  sheet.appendRow([
    ts(),
    data.first_name || '',
    data.email      || '',
    data.phone      || ''
  ]);
  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to:      NOTIFY_EMAIL,
      subject: 'New Email Signup: ' + (data.first_name || '') + ' <' + (data.email || '') + '>',
      body:
        'Name:  ' + (data.first_name || '') + '\n' +
        'Email: ' + (data.email || '') + '\n' +
        'Phone: ' + (data.phone || 'not provided')
    });
  }
}

function handleContact(data) {
  var sheet = getOrCreateSheet('Contact Messages', [
    'Timestamp', 'First Name', 'Last Name', 'Email', 'Subject', 'Message'
  ]);
  sheet.appendRow([
    ts(),
    data.first_name || '',
    data.last_name  || '',
    data.email      || '',
    data.subject    || '',
    data.message    || ''
  ]);
  if (NOTIFY_EMAIL) {
    MailApp.sendEmail({
      to:      NOTIFY_EMAIL,
      replyTo: data.email || '',
      subject: 'Contact [' + (data.subject || 'general') + ']: ' +
               (data.first_name || '') + ' ' + (data.last_name || ''),
      body:
        'From:    ' + (data.first_name || '') + ' ' + (data.last_name || '') +
                  ' <' + (data.email || '') + '>\n' +
        'Subject: ' + (data.subject || '') + '\n\n' +
        (data.message || '')
    });
  }
}
