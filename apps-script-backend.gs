// Paste this whole file into Extensions > Apps Script in your Google Sheet,
// replacing any starter code. Then deploy as a Web App (see setup notes).

var SHEET_NAME = 'Registrations';
var PROP_KEY = 'CATEGORIES';
var DEFAULT_CATEGORIES = ['U400', 'U600', 'U1000', 'U1500', 'Open'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Rating', 'Category', 'USCF ID']);
  }
  return sheet;
}

function getCategories_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed;
    } catch (e) {}
  }
  return DEFAULT_CATEGORIES;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = values.slice(1);
  var entries = rows
    .filter(function (r) { return r[1]; })
    .map(function (r) {
      return {
        timestamp: r[0],
        name: r[1],
        rating: r[2],
        category: r[3],
        uscfid: r[4] || null
      };
    });
  return jsonResponse_({ categories: getCategories_(), entries: entries });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'Bad request' });
  }

  var action = body.action;

  if (action === 'register') {
    if (!body.name || body.rating === undefined || !body.category) {
      return jsonResponse_({ ok: false, error: 'Missing fields' });
    }
    var sheet = getSheet_();
    sheet.appendRow([
      new Date().toISOString(),
      body.name,
      body.rating,
      body.category,
      body.uscfid || ''
    ]);
    return jsonResponse_({ ok: true });
  }

  if (action === 'saveCategories') {
    if (!Array.isArray(body.categories) || body.categories.length === 0) {
      return jsonResponse_({ ok: false, error: 'No categories' });
    }
    PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(body.categories));
    return jsonResponse_({ ok: true });
  }

  if (action === 'clear') {
    var sheetToClear = getSheet_();
    var lastRow = sheetToClear.getLastRow();
    if (lastRow > 1) {
      sheetToClear.deleteRows(2, lastRow - 1);
    }
    return jsonResponse_({ ok: true });
  }

  return jsonResponse_({ ok: false, error: 'Unknown action' });
}
