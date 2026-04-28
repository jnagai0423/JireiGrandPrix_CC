/**
 * Gmail -> Google Sheets import for a mailing list.
 *
 * 使い方:
 * 1. このスクリプトを出力先スプレッドシートに紐付ける
 * 2. CONFIG の SUBJECT_KEYWORDS / SUBJECT_PREFIXES を必要に応じて変更
 * 3. 初回だけ importCgMailToSheet() を手動実行して権限承認
 * 4. createHourlyTrigger() または createDailyTrigger() を手動実行
 */
const CONFIG = {
  // スプレッドシートに紐付くコンテナバインドGASなら空でOK。
  // スタンドアロンGASで使う場合は、対象スプレッドシートIDを入れてください。
  SPREADSHEET_ID: '',

  MAILING_LIST_ADDRESS: 'cg@cloudcircus.co.jp',
  MASTER_SHEET_NAME: 'メール一覧',
  PROCESSED_SHEET_NAME: '_processed_message_ids',

  // 件名フィルター: どちらかを満たすメールだけ取り込みます。
  SUBJECT_KEYWORDS: [
    // '資料請求',
    // '問い合わせ',
  ],
  SUBJECT_PREFIXES: [
    '【お問い合わせ】',
  ],

  // Gmail検索対象期間。定期実行なら重複防止があるため広めでも問題ありません。
  SEARCH_NEWER_THAN: '30d',
  MAX_THREADS_PER_RUN: 100,

  // 本文が長すぎる場合の上限。不要なら大きめの値にしてください。
  BODY_MAX_LENGTH: 50000,
};

const OUTPUT_HEADERS = ['受信日時', '送信者名', 'メールアドレス', '件名', '本文'];

/**
 * メイン処理。条件に一致する未処理メールを「メール一覧」に追記し、
 * その後、送信元メールアドレスのドメインごとに別シートへ振り分けます。
 */
function importCgMailToSheet() {
  const ss = getTargetSpreadsheet();
  const masterSheet = ensureMasterSheet(ss);
  const processedSheet = ensureProcessedSheet(ss);
  const processedIds = getProcessedMessageIds(processedSheet);

  const query = buildGmailQuery();
  const threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS_PER_RUN);
  const rows = [];
  const newlyProcessedIds = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const messageId = message.getId();
      if (processedIds.has(messageId)) return;
      if (!isMessageForMailingList(message)) return;

      const subject = message.getSubject() || '';
      if (!matchesSubjectCondition(subject)) return;

      const from = parseFrom(message.getFrom());
      rows.push([
        message.getDate(),
        from.name,
        from.email,
        subject,
        normalizeBody(message.getPlainBody()),
      ]);
      newlyProcessedIds.push([messageId, new Date()]);
    });
  });

  if (rows.length > 0) {
    const startRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(startRow, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);
    processedSheet
      .getRange(processedSheet.getLastRow() + 1, 1, newlyProcessedIds.length, 2)
      .setValues(newlyProcessedIds);
  }

  splitMasterSheetBySenderDomain(ss, masterSheet);
  Logger.log(`取り込み完了: ${rows.length} 件 / Gmail検索: ${threads.length} スレッド`);
}

/**
 * マスターシートの全データを読み、送信者メールアドレスのドメインごとに
 * apple.com / toyota.co.jp のようなシートへ再出力します。
 */
function splitMasterSheetBySenderDomain(ss, masterSheet) {
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  const values = masterSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
  const rowsByDomain = {};

  values.forEach(row => {
    const email = String(row[2] || '').trim();
    const domain = extractDomain(email);
    if (!domain) return;
    if (!rowsByDomain[domain]) rowsByDomain[domain] = [];
    rowsByDomain[domain].push(row);
  });

  Object.keys(rowsByDomain).sort().forEach(domain => {
    const sheetName = buildSafeSheetName(domain);
    const sheet = ensureSheetWithHeaders(ss, sheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    sheet
      .getRange(2, 1, rowsByDomain[domain].length, OUTPUT_HEADERS.length)
      .setValues(rowsByDomain[domain]);
    formatOutputSheet(sheet);
  });
}

/** 毎時実行トリガーを作成します。初回セットアップ時に手動実行してください。 */
function createHourlyTrigger() {
  deleteTriggersForFunction_('importCgMailToSheet');
  ScriptApp.newTrigger('importCgMailToSheet').timeBased().everyHours(1).create();
}

/** 毎日実行トリガーを作成します。初回セットアップ時に手動実行してください。 */
function createDailyTrigger() {
  deleteTriggersForFunction_('importCgMailToSheet');
  ScriptApp.newTrigger('importCgMailToSheet').timeBased().everyDays(1).atHour(9).create();
}

/** 対象関数の既存トリガーを削除します。 */
function deleteImportTriggers() {
  deleteTriggersForFunction_('importCgMailToSheet');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gmail取り込み')
    .addItem('今すぐ取り込み', 'importCgMailToSheet')
    .addSeparator()
    .addItem('毎時トリガー作成', 'createHourlyTrigger')
    .addItem('毎日トリガー作成', 'createDailyTrigger')
    .addItem('取り込みトリガー削除', 'deleteImportTriggers')
    .addToUi();
}

function getTargetSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function buildGmailQuery() {
  return [
    'in:inbox',
    `to:${CONFIG.MAILING_LIST_ADDRESS}`,
    `newer_than:${CONFIG.SEARCH_NEWER_THAN}`,
  ].join(' ');
}

function isMessageForMailingList(message) {
  const target = CONFIG.MAILING_LIST_ADDRESS.toLowerCase();
  const fields = [
    message.getTo(),
    message.getCc(),
    message.getBcc(),
  ].join(' ').toLowerCase();
  return fields.indexOf(target) !== -1;
}

function matchesSubjectCondition(subject) {
  const normalizedSubject = String(subject || '');
  const keywords = CONFIG.SUBJECT_KEYWORDS.filter(Boolean);
  const prefixes = CONFIG.SUBJECT_PREFIXES.filter(Boolean);

  if (keywords.length === 0 && prefixes.length === 0) return true;

  const hasKeyword = keywords.some(keyword => normalizedSubject.indexOf(keyword) !== -1);
  const hasPrefix = prefixes.some(prefix => normalizedSubject.indexOf(prefix) === 0);
  return hasKeyword || hasPrefix;
}

function parseFrom(fromText) {
  const text = String(fromText || '').trim();
  const match = text.match(/^(.*)<([^<>]+)>$/);
  if (!match) {
    return {
      name: '',
      email: text.replace(/^mailto:/i, '').trim(),
    };
  }

  return {
    name: match[1].replace(/^"|"$/g, '').trim(),
    email: match[2].replace(/^mailto:/i, '').trim(),
  };
}

function extractDomain(email) {
  const match = String(email || '').toLowerCase().match(/@([^@\s>]+)$/);
  return match ? match[1] : '';
}

function normalizeBody(body) {
  const text = String(body || '').replace(/\r\n/g, '\n').trim();
  if (text.length <= CONFIG.BODY_MAX_LENGTH) return text;
  return text.slice(0, CONFIG.BODY_MAX_LENGTH) + '\n...本文が長いため省略';
}

function ensureMasterSheet(ss) {
  const sheet = ensureSheetWithHeaders(ss, CONFIG.MASTER_SHEET_NAME);
  formatOutputSheet(sheet);
  return sheet;
}

function ensureSheetWithHeaders(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const headerValues = sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).getValues()[0];
  const hasHeaders = OUTPUT_HEADERS.every((header, index) => headerValues[index] === header);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
  }
  return sheet;
}

function ensureProcessedSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME) || ss.insertSheet(CONFIG.PROCESSED_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, 2).getValues()[0];
  if (headers[0] !== 'message_id' || headers[1] !== 'processed_at') {
    sheet.getRange(1, 1, 1, 2).setValues([['message_id', 'processed_at']]);
  }
  sheet.hideSheet();
  return sheet;
}

function getProcessedMessageIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return new Set(ids.filter(Boolean).map(String));
}

function buildSafeSheetName(domain) {
  return String(domain || 'unknown')
    .replace(/[\[\]\*\?\/\\:]/g, '_')
    .slice(0, 100);
}

function formatOutputSheet(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);
  sheet.setColumnWidth(5, 500);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setWrap(true);
  }
}

function deleteTriggersForFunction_(functionName) {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === functionName)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}
