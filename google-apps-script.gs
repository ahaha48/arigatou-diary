/**
 * @OnlyCurrentDoc
 */
const SHEET_NAME = 'ありがとう日記ログ';
const SPREADSHEET_ID = '';
const SCRIPT_VERSION = '2026-05-08-sync-load';

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    return createResponse_(appendPayload_(payload));
  } catch (error) {
    return createResponse_({ ok: false, error: error.message });
  }
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;

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
      version: SCRIPT_VERSION,
    }, callback);
  } catch (error) {
    return createResponse_({ ok: false, error: error.message }, callback);
  }
}

function appendPayload_(payload) {
  const sheet = getLogSheet_();
  const entries = Array.isArray(payload.entries) ? payload.entries : [];

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
    payload.userAgent || '',
  ]);

  return {
    ok: true,
    row: sheet.getLastRow(),
    sheetName: sheet.getName(),
    count: entries.length,
    version: SCRIPT_VERSION,
  };
}

function loadDiary_(writer) {
  const sheet = getLogSheet_();
  const values = sheet.getDataRange().getValues();
  const days = {};
  const targetWriter = String(writer || '').trim();

  for (let index = 1; index < values.length; index += 1) {
    const row = values[index];
    const rowWriter = String(row[4] || '').trim();
    const date = formatDateKey_(row[2]);
    const entries = [
      row[7],
      row[8],
      row[9],
      ...String(row[10] || '').split('\n'),
    ].map((entry) => String(entry || '').trim()).filter(Boolean);

    if (!date || entries.length === 0) continue;
    if (targetWriter && rowWriter !== targetWriter) continue;

    days[date] = entries;
  }

  return {
    ok: true,
    days,
    dayCount: Object.keys(days).length,
    sheetName: sheet.getName(),
    writer: targetWriter,
    version: SCRIPT_VERSION,
  };
}

function formatDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function getLogSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('スプレッドシートからApps Scriptを開くか、SPREADSHEET_IDを設定してください。');
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

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
      'ブラウザ',
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    try {
      return JSON.parse(e.parameter.payload);
    } catch (error) {
      return {};
    }
  }

  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    return {};
  }
}

function createResponse_(payload, callback) {
  const body = JSON.stringify(payload);

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}
