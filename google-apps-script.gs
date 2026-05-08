/**
 * @OnlyCurrentDoc
 */
var SHEET_NAME = 'ありがとう日記ログ';
var SPREADSHEET_ID = '';
var SCRIPT_VERSION = '2026-05-08-sync-load-compat';

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    var result = appendPayload_(payload);
    return createResponse_(result);
  } catch (err) {
    return createResponse_({
      ok: false,
      error: getErrorMessage_(err),
      version: SCRIPT_VERSION
    });
  }
}

function doGet(e) {
  var callback = e && e.parameter ? e.parameter.callback : '';

  try {
    if (e && e.parameter && e.parameter.action === 'load') {
      return createResponse_(loadDiary_(e.parameter.writer || ''), callback);
    }

    if (e && e.parameter && e.parameter.payload) {
      return createResponse_(appendPayload_(parsePayload_(e)), callback);
    }

    return createResponse_({
      ok: true,
      message: '接続OK: ありがとう日記の受信用Webアプリです。version=' + SCRIPT_VERSION,
      version: SCRIPT_VERSION
    }, callback);
  } catch (err) {
    return createResponse_({
      ok: false,
      error: getErrorMessage_(err),
      version: SCRIPT_VERSION
    }, callback);
  }
}

function appendPayload_(payload) {
  var sheet = getLogSheet_();
  var entries = normalizeEntries_(payload.entries);

  if (!payload.date || entries.length === 0) {
    throw new Error('日付またはありがとう内容がありません。');
  }

  sheet.appendRow([
    new Date(),
    payload.savedAt || '',
    payload.date || '',
    payload.dateLabel || '',
    payload.writer || '',
    entries.length,
    entries.length >= 3 ? '達成' : '未達成',
    entries[0] || '',
    entries[1] || '',
    entries[2] || '',
    entries.slice(3).join('\n'),
    entries.join('\n'),
    payload.source || '',
    payload.userAgent || ''
  ]);

  return {
    ok: true,
    row: sheet.getLastRow(),
    sheetName: sheet.getName(),
    count: entries.length,
    version: SCRIPT_VERSION
  };
}

function loadDiary_(writer) {
  var sheet = getLogSheet_();
  var values = sheet.getDataRange().getValues();
  var days = {};
  var targetWriter = String(writer || '').trim();

  for (var index = 1; index < values.length; index += 1) {
    var row = values[index];
    var rowWriter = String(row[4] || '').trim();
    var date = formatDateKey_(row[2]);
    var entries = normalizeEntries_([
      row[7],
      row[8],
      row[9]
    ].concat(String(row[10] || '').split('\n')));

    if (!date || entries.length === 0) {
      continue;
    }

    if (targetWriter && rowWriter !== targetWriter) {
      continue;
    }

    days[date] = entries;
  }

  return {
    ok: true,
    days: days,
    dayCount: Object.keys(days).length,
    sheetName: sheet.getName(),
    writer: targetWriter,
    version: SCRIPT_VERSION
  };
}

function getLogSheet_() {
  var spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('スプレッドシートからApps Scriptを開くか、SPREADSHEET_IDを設定してください。');
  }

  var sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '受信日時',
      '保存日時',
      '日付',
      '日付表示',
      'お名前',
      '数',
      '達成',
      'ありがとう1',
      'ありがとう2',
      'ありがとう3',
      '追加のありがとう',
      '全文',
      '送信元',
      'ブラウザ'
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return parseJson_(e.parameter.payload);
  }

  if (e && e.postData && e.postData.contents) {
    return parseJson_(e.postData.contents);
  }

  return {};
}

function parseJson_(text) {
  try {
    return JSON.parse(text);
  } catch (err) {
    return {};
  }
}

function normalizeEntries_(entries) {
  if (!Array.isArray(entries)) {
    return [];
  }

  var cleanEntries = [];
  for (var index = 0; index < entries.length; index += 1) {
    var entry = String(entries[index] || '').trim();
    if (entry) {
      cleanEntries.push(entry);
    }
  }
  return cleanEntries;
}

function formatDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function createResponse_(payload, callback) {
  var body = JSON.stringify(payload);

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(err) {
  if (err && err.message) {
    return err.message;
  }
  return String(err);
}
