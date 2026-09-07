const SPREADSHEET_ID = 'YOUR SPREADSHEET ID HERE';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[1]; // Tab 2

    const raw = (e.postData && e.postData.contents) ? e.postData.contents : '';
    const text = unescapeIfJsonString(raw);
    const payload = parseKeyValueLines(text);

    writeToColumns(sheet, payload);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Decodes a JSON-stringified body back into real newlines, if that's what arrived
function unescapeIfJsonString(text) {
  try {
    const parsed = JSON.parse(text);
    if (typeof parsed === 'string') {
      return parsed;
    }
  } catch (err) {
    // not a JSON string literal - use as-is
  }
  return text;
}

// Splits "NRIC:4564F\nName:Tester" into { "NRIC": "4564F", "Name": "Tester" }
function parseKeyValueLines(text) {
  const result = {};
  text.split('\n').forEach(function (line) {
    line = line.trim();
    if (!line) return;
    const idx = line.indexOf(':');
    if (idx === -1) return;
    const key = line.substring(0, idx).trim();
    const value = line.substring(idx + 1).trim();
    if (key) result[key] = value;
  });
  return result;
}

// Strips spaces/punctuation and lowercases so field names match headers
// regardless of formatting differences
function normalize(str) {
  return str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
}

function writeToColumns(sheet, payload) {
  // Read headers from row 1 only, so stray data further down can't affect the width
  const headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];

  const columnMap = {};
  headers.forEach(function (header, i) {
    if (header) columnMap[normalize(header)] = i + 1;
  });

  const newRow = sheet.getLastRow() + 1;

  // Timestamp always in column A, with explicit date + time format
  sheet.getRange(newRow, 1)
    .setValue(new Date())
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');

  Object.keys(payload).forEach(function (key) {
    const col = columnMap[normalize(key)];
    if (col) {
      sheet.getRange(newRow, col).setValue(payload[key]);
    }
  });
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'Webhook is live. Use POST to submit data.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
