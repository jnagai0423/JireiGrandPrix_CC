/**
 * Gmail 取り込み処理の設定定義。
 *
 * 使い方:
 * 1. このスクリプトを出力先スプレッドシートに紐付ける
 * 2. CONFIG の SUBJECT_KEYWORDS / SUBJECT_PREFIXES を必要に応じて変更
 * 3. 初回だけ importCgMailToSheet() を手動実行して権限承認
 *
 * 補足:
 * - このファイルは「値の定義のみ」を担います。
 * - 実処理は sheets_import.js / sheets_split.js / sheets_ui_triggers.js にあります。
 * - トリガーは Apps Script 画面から手動で設定する運用を想定しています。
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

// すべての出力シートで共通利用する列定義。
const OUTPUT_HEADERS = ['受信日時', '送信者名', 'メールアドレス', '件名', '本文'];
