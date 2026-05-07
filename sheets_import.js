/**
 * メイン処理。条件に一致する未処理メールを「メール一覧」に追記し、
 * その後、送信元メールアドレスのドメインごとに別シートへ振り分けます。
 *
 * 処理の要点:
 * - Gmail検索 -> 条件判定 -> マスター追記 -> 処理済みID記録 -> ドメイン分割
 * - 処理済みIDを保存することで、定期実行時の重複取り込みを防ぎます。
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
      // すでに取り込み済みのメッセージはスキップ。
      if (processedIds.has(messageId)) return;
      // To/Cc/Bcc のどこにも対象MLアドレスがなければ対象外。
      if (!isMessageForMailingList(message)) return;

      const subject = message.getSubject() || '';
      // 件名条件に一致しない場合は対象外。
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

  // 1件以上あるときだけシート書き込みを行う。
  if (rows.length > 0) {
    const startRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(startRow, 1, rows.length, OUTPUT_HEADERS.length).setValues(rows);
    processedSheet
      .getRange(processedSheet.getLastRow() + 1, 1, newlyProcessedIds.length, 2)
      .setValues(newlyProcessedIds);
  }

  // マスター追記後にドメイン別シートを再構築。
  splitMasterSheetBySenderDomain(ss, masterSheet);
  Logger.log(`取り込み完了: ${rows.length} 件 / Gmail検索: ${threads.length} スレッド`);
}

/** コンテナバインド/スタンドアロン両対応で対象スプレッドシートを返す。 */
function getTargetSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Gmail検索クエリを組み立てる。 */
function buildGmailQuery() {
  return [
    'in:inbox',
    `to:${CONFIG.MAILING_LIST_ADDRESS}`,
    `newer_than:${CONFIG.SEARCH_NEWER_THAN}`,
  ].join(' ');
}

/** メールヘッダ(To/Cc/Bcc)に対象MLアドレスが含まれるかを判定する。 */
function isMessageForMailingList(message) {
  const target = CONFIG.MAILING_LIST_ADDRESS.toLowerCase();
  const fields = [
    message.getTo(),
    message.getCc(),
    message.getBcc(),
  ].join(' ').toLowerCase();
  return fields.indexOf(target) !== -1;
}

/**
 * 件名条件判定。
 * - KEYWORDS: 部分一致
 * - PREFIXES: 前方一致
 * - 両方空なら全件許可
 */
function matchesSubjectCondition(subject) {
  const normalizedSubject = String(subject || '');
  const keywords = CONFIG.SUBJECT_KEYWORDS.filter(Boolean);
  const prefixes = CONFIG.SUBJECT_PREFIXES.filter(Boolean);

  if (keywords.length === 0 && prefixes.length === 0) return true;

  const hasKeyword = keywords.some(keyword => normalizedSubject.indexOf(keyword) !== -1);
  const hasPrefix = prefixes.some(prefix => normalizedSubject.indexOf(prefix) === 0);
  return hasKeyword || hasPrefix;
}

/** "表示名 <mail@example.com>" 形式を name/email に分解する。 */
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

/** 本文を整形し、長すぎる場合は設定値で打ち切る。 */
function normalizeBody(body) {
  const text = String(body || '').replace(/\r\n/g, '\n').trim();
  if (text.length <= CONFIG.BODY_MAX_LENGTH) return text;
  return text.slice(0, CONFIG.BODY_MAX_LENGTH) + '\n...本文が長いため省略';
}

/** 処理済み管理シートを保証し、必要ならヘッダーを初期化する。 */
function ensureProcessedSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET_NAME) || ss.insertSheet(CONFIG.PROCESSED_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, 2).getValues()[0];
  if (headers[0] !== 'message_id' || headers[1] !== 'processed_at') {
    sheet.getRange(1, 1, 1, 2).setValues([['message_id', 'processed_at']]);
  }
  sheet.hideSheet();
  return sheet;
}

/** 処理済み message_id を Set で返し、高速に重複判定できるようにする。 */
function getProcessedMessageIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return new Set(ids.filter(Boolean).map(String));
}
