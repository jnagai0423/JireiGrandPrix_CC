/**
 * delivery_writing@cloudcircus.co.jp 宛（メーリングリスト）で
 * 受信トレイにあるメールをスプレッドシートに出力するスクリプト
 *
 * 【このリポジトリと Google の「接続」について】
 * Cursor から script.google.com へ直接プッシュすることはできません。
 * コードは Apps Script エディタにコピーするか、clasp（Google の CLI）で push してください。
 *
 * 【バインド（コンテナ）にしたい場合】
 * スプレッドシートを開き「拡張機能 > Apps スクリプト」で開いたプロジェクトにこのコードを置くと、
 * そのブックに紐づいた（コンテナバインド）スクリプトになります。
 * メニュー「メール出力」は、バインド済みのときだけスプレッドシートを開いたタイミングで表示されます。
 * （script.google.com だけのスタンドアロンプロジェクトでは onOpen はシートを開いても動きません。）
 *
 * 想定スプレッドシート:
 * https://docs.google.com/spreadsheets/d/1Ahx2-vEILypuBQjaDi2kSOkIYLUE5XLXPhtplECEca8/edit
 *
 * 既存の Apps Script プロジェクト（スタンドアロン）例:
 * https://script.google.com/home/projects/1YuzZuN2CDbcrlMTWktFw7ochUHBgC2YB0CpsTagRS5z67FFUYEkO3Euo/edit
 * → この場合は下の SPREADSHEET_ID を必ず設定し、実行はエディタの「実行」かトリガーで行ってください。
 *
 * 検索条件を変えたい場合は GMAIL_QUERY を編集してください。
 */

/**
 * 出力先スプレッドシート ID（URL の /d/ と /edit のあいだ）。
 * コンテナバインドで同じブックだけ使うなら null にして getActiveSpreadsheet() に任せてもよい。
 */
const SPREADSHEET_ID = '1Ahx2-vEILypuBQjaDi2kSOkIYLUE5XLXPhtplECEca8';

/** Gmail 検索クエリ（必要に応じて調整） */
const GMAIL_QUERY =
  'to:delivery_writing@cloudcircus.co.jp in:inbox';

/** 1 回の search で取得するスレッド数（最大 500） */
const PAGE_SIZE = 500;

/** シート名 */
const SHEET_NAME = 'delivery_writing';

/** ヘッダー行 */
const HEADERS = [
  '受信日時',
  '送信者',
  '宛先',
  '件名',
  'スニペット',
  'スレッド ID',
  'メッセージ ID',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('メール出力')
    .addItem('受信トレイをシートに出力', 'exportInboxToSheet')
    .addToUi();
}

/**
 * 検索に一致したメールをシートに書き込む（既存の同名シートは上書き）
 */
function exportInboxToSheet() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートを開けません。SPREADSHEET_ID を設定するか、スプレッドシートにバインドしてください。');
  }

  const rows = collectRows_(GMAIL_QUERY);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length + 1, HEADERS.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, HEADERS.length);

  SpreadsheetApp.getUi().alert(
    '完了',
    '件数: ' + rows.length + ' 件を「' + SHEET_NAME + '」に出力しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * @param {string} query
 * @return {Array<Array<string|Date>>}
 */
function collectRows_(query) {
  const rows = [];
  let start = 0;

  while (true) {
    const threads = GmailApp.search(query, start, PAGE_SIZE);
    if (!threads || threads.length === 0) {
      break;
    }

    for (let t = 0; t < threads.length; t++) {
      const thread = threads[t];
      const messages = thread.getMessages();
      for (let m = 0; m < messages.length; m++) {
        const msg = messages[m];
        // 受信トレイにないメッセージは除外（アーカイブ済み等）
        if (!msg.isInInbox()) {
          continue;
        }
        rows.push([
          msg.getDate(),
          msg.getFrom(),
          msg.getTo(),
          msg.getSubject(),
          msg.getPlainBody().slice(0, 500).replace(/\s+/g, ' ').trim(),
          thread.getId(),
          msg.getId(),
        ]);
      }
    }

    if (threads.length < PAGE_SIZE) {
      break;
    }
    start += PAGE_SIZE;
  }

  // 新しい順
  rows.sort(function (a, b) {
    return b[0] - a[0];
  });

  return rows;
}
