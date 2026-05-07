/**
 * @OnlyCurrentDoc
 */
const SHEET_NAME = 'ありがとう日記ログ';
const SPREADSHEET_ID = '';

function doPost(e) {
  const sheet = getLogSheet_();
  const payload = parsePayload_(e);
  const entries = Array.isArray(payload.entries) ? payload.entries : [];

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

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput('ありがとう日記の受信用Webアプリです。')
    .setMimeType(ContentService.MimeType.TEXT);
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
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    return {};
  }
}
