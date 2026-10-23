// Google Apps Script — paste into Extensions > Apps Script in your Google Sheet

const SHEET_NAME = 'Sheet1';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // Check if email already submitted
    const emails = sheet.getRange('B2:B' + sheet.getLastRow()).getValues().flat();
    const emailLower = (data.email || '').toLowerCase();
    if (emails.some(e => e.toLowerCase() === emailLower)) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'already_played' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Append new score
    sheet.appendRow([
      data.name,
      emailLower,
      data.score,
      new Date().toISOString()
    ]);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange('A2:C' + lastRow).getValues();
  const leaderboard = data
    .map(row => ({ name: row[0], email: row[1], score: row[2] }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);

  // Strip emails from response for privacy
  const safe = leaderboard.map(e => ({ name: e.name, score: e.score }));

  return ContentService.createTextOutput(JSON.stringify(safe))
    .setMimeType(ContentService.MimeType.JSON);
}
