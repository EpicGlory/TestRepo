// Cosmic Pig BBQ — Pre-Order Google Apps Script
// Deploy as a Web App:
//   Extensions > Apps Script > Deploy > New deployment
//   Type: Web app | Execute as: Me | Who has access: Anyone
// Then copy the Web App URL into js/main.js (PREORDER_SCRIPT_URL constant).

var NOTIFY_EMAIL = 'ryan@cosmicpigbbq.com';
var SHEET_NAME   = 'Pre-Orders';

function doPost(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp', 'First Name', 'Last Name', 'Email', 'Phone',
        'City/State', 'Delivery Preference', 'Rubs', 'Quantity', 'Source', 'Notes'
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    }

    var data = JSON.parse(e.postData.contents);
    var ts   = Utilities.formatDate(
      new Date(), 'America/Denver', 'MM/dd/yyyy HH:mm:ss'
    );

    sheet.appendRow([
      ts,
      data.first_name        || '',
      data.last_name         || '',
      data.email             || '',
      data.phone             || '',
      data.city_state        || '',
      data.delivery_preference || '',
      data.rubs              || '',
      data.quantity          || '',
      data.source            || '',
      data.notes             || ''
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to:      NOTIFY_EMAIL,
        subject: 'New Pre-Order: ' + (data.first_name || '') + ' ' + (data.last_name || ''),
        body:
          'Name:     ' + (data.first_name || '') + ' ' + (data.last_name || '') + '\n' +
          'Email:    ' + (data.email || '') + '\n' +
          'Phone:    ' + (data.phone || 'not provided') + '\n' +
          'Location: ' + (data.city_state || 'not provided') + '\n' +
          'Rubs:     ' + (data.rubs || 'none selected') + '\n' +
          'Quantity: ' + (data.quantity || 'not specified') + '\n' +
          'Delivery: ' + (data.delivery_preference || 'not specified') + '\n' +
          'Source:   ' + (data.source || '') + '\n' +
          'Notes:    ' + (data.notes || '') + '\n'
      });
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
