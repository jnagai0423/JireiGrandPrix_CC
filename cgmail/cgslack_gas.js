/** ML・共有アドレス（To または Cc にのみ含まれるメールを出力） */
const FILTER_EMAIL = 'cg@cloudcircus.co.jp';
const SHEET_NAME = 'cg一覧';
const SPREADSHEET_ID = '1hi6IzHBo4tnRjfRFimJlfBldN9IxTWZDkSE96GlMnTM';

function fetchFilteredEmails() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['日時', '送信者', '件名', '本文（先頭200文字）', 'メールID']);
  }

  // 既存のメールIDを取得（重複防止）
  const existingIds = sheet.getRange(2, 5, Math.max(sheet.getLastRow() - 1, 1), 1)
    .getValues().flat().filter(String);

  const query = `(to:${FILTER_EMAIL} OR cc:${FILTER_EMAIL}) in:inbox`;
  const threads = GmailApp.search(query, 0, 50); // 最大50スレッド

  const rows = [];
  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const id = message.getId();
      if (existingIds.includes(id)) continue; // 重複スキップ
      if (!isAddressInToOrCc_(message, FILTER_EMAIL)) continue;

      rows.push([
        message.getDate(),
        message.getFrom(),
        message.getSubject(),
        message.getPlainBody().slice(0, 200),
        id,
      ]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  Logger.log(`${rows.length} 件追加しました`);
}

/**
 * To / Cc のヘッダ文字列に、指定アドレスが含まれるか（スレッド内の他メッセージを除外）
 */
function isAddressInToOrCc_(message, email) {
  const esc = email.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp('\\b' + esc + '\\b', 'i');
  const to = message.getTo() || '';
  const cc = message.getCc() || '';
  return re.test(to) || re.test(cc);
}
