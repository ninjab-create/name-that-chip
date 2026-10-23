// Google Apps Script — paste into Extensions > Apps Script in your Google Sheet
// Returns HTML with postMessage so hidden iframes carry Google auth cookies.
// Supports GET-based submit (?submitName=&submitEmail=&submitScore=) for iframe compat.

const SHEET_NAME = 'Sheet1';

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var params = e ? e.parameter : {};

  // ---- SUBMIT via GET (iframe-friendly) ----
  if (params.submitName && params.submitEmail && params.submitScore) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var emailLower = params.submitEmail.toLowerCase();
      if (lastRow >= 2) {
        var emails = sheet.getRange('B2:B' + lastRow).getValues().flat();
        if (emails.some(function(em) { return String(em).toLowerCase() === emailLower; })) {
          return buildHtmlResponse({ type: 'submitScore', error: 'already_played' });
        }
      }
      sheet.appendRow([
        params.submitName,
        emailLower,
        Number(params.submitScore),
        new Date().toISOString()
      ]);
      return buildHtmlResponse({ type: 'submitScore', success: true });
    } finally {
      lock.releaseLock();
    }
  }

  // ---- CHECK EMAIL ----
  if (params.checkEmail) {
    var emailLower2 = params.checkEmail.toLowerCase();
    if (lastRow >= 2) {
      var emails2 = sheet.getRange('B2:B' + lastRow).getValues().flat();
      var found = emails2.some(function(em) { return String(em).toLowerCase() === emailLower2; });
      if (found) {
        var rows = sheet.getRange('A2:C' + lastRow).getValues();
        var row = rows.find(function(r) { return String(r[1]).toLowerCase() === emailLower2; });
        return buildHtmlResponse({ type: 'checkEmail', already_played: true, score: row ? row[2] : 0 });
      }
    }
    return buildHtmlResponse({ type: 'checkEmail', already_played: false });
  }

  // ---- LEADERBOARD ----
  if (lastRow < 2) {
    return buildHtmlResponse({ type: 'leaderboard', data: [] });
  }

  var data = sheet.getRange('A2:C' + lastRow).getValues();
  var leaderboard = data
    .map(function(row) { return { name: row[0], score: row[2] }; })
    .sort(function(a, b) { return b.score - a.score; })
    .slice(0, 20);

  return buildHtmlResponse({ type: 'leaderboard', data: leaderboard });
}

// Keep doPost for backwards compat (direct API calls)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var emails = sheet.getRange('B2:B' + lastRow).getValues().flat();
      var emailLower = (data.email || '').toLowerCase();
      if (emails.some(function(em) { return String(em).toLowerCase() === emailLower; })) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'already_played' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    sheet.appendRow([data.name, (data.email || '').toLowerCase(), data.score, new Date().toISOString()]);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function buildHtmlResponse(payload) {
  var json = JSON.stringify(payload);
  var html = '<!DOCTYPE html><html><body><script>'
    + 'window.onload=function(){'
    + 'window.parent.postMessage(' + json + ',"*");'
    + '};'
    + '<\/script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
