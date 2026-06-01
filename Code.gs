// Apps Script: Run Club, Soho House Tokyo x Adherence Studio RSVP capture
// Bound to a Google Sheet. Receives a POST from the landing page form
// and appends one row per request to the interest list.
//
// SETUP
// 1. Create a Google Sheet (this becomes your interest list).
// 2. Extensions > Apps Script. Delete the starter code, paste this in.
// 3. Deploy > New deployment > type "Web app".
//      - Execute as: Me
//      - Who has access: Anyone
//    Authorize when prompted, then copy the Web app URL (ends in /exec).
// 4. Paste that URL into index.html as the ENDPOINT value.

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Social', 'Soho House', 'Event', 'User Agent']);
      sheet.getRange('A1:G1').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.social || '',
      data.member || '',
      data.event || 'Run Club · Soho House Tokyo x Adherence Studio (June 7)',
      data.userAgent || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Open the Web App URL in a browser to confirm it's live.
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Run Club RSVP endpoint live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
