/** フォーム回答シートを取得（先頭シートが CONFIG などでもずれないようにする） */
function getFormResponseSheet(ss) {
  const explicit = String(CONFIG.FORM_RESPONSE_SHEET_NAME || '').trim();
  if (explicit) {
    const named = ss.getSheetByName(explicit);
    if (named) return named;
    Logger.log('CONFIG.FORM_RESPONSE_SHEET_NAME が見つかりません: ' + explicit);
  }
  const formSheets = ss.getSheets().filter(s => /^フォームの回答/.test(s.getName()));
  if (formSheets.length) return formSheets[0];
  return ss.getSheets()[0];
}

// フォーム送信トリガー用（「フォーム送信時」に設定）
function onFormSubmit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFormResponseSheet(ss);
  const lastRow = sheet.getLastRow();
  const outCol = ensureOutputColumns(sheet);

  // ステータス更新（処理中）
  sheet.getRange(lastRow, outCol.statusCol).setValue('処理中...');

  try {
    const data = getRowDataByHeader(sheet, lastRow);

    if (!data.clientName) throw new Error('顧客名が空です');

    // Gemini API でAIコメント生成
    const aiComment = generateAIComment(data);

    // スライド生成
    const slideUrl = createSlide(data, aiComment);

    // スプレッドシートに結果書き込み
    sheet.getRange(lastRow, outCol.slideUrlCol).setValue(slideUrl);
    sheet.getRange(lastRow, outCol.statusCol).setValue('完了');

    // メール通知
    sendNotification(data.clientName, slideUrl, aiComment);

    Logger.log(`[完了] ${data.clientName} → ${slideUrl}`);

  } catch (err) {
    sheet.getRange(lastRow, outCol.statusCol).setValue('エラー: ' + err.message);
    Logger.log('[エラー] ' + err.toString());
    GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '[エラー] 事例スライド生成失敗', err.toString());
  }
}

// 完成通知メール
function sendNotification(clientName, url, aiComment) {
  if (!CONFIG.NOTIFY_EMAIL || CONFIG.NOTIFY_EMAIL === 'your-email@example.com') return;

  const subject = `【完成】${clientName} の事例スライドが生成されました`;
  const body = `
${clientName} の事例スライドが自動生成されました。

▼ スライドを開く
${url}

── AIコメント ──
${aiComment}

このメールは自動送信です。
  `.trim();

  GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, body);
}

// スプレッドシートを開いた時のカスタムメニュー
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('事例GP')
    .addItem('① コピー元スライドを新規作成（手元に無いときのみ）', 'createTemplate')
    .addSeparator()
    .addItem('② 最新行でスライドを手動生成', 'runManually')
    .addToUi();
}

// UIアラートを安全に表示（表示不可環境ではログへ）
function safeAlert(title, message) {
  if (!CONFIG.ENABLE_UI_ALERT) {
    Logger.log(`safeAlert: 無効化中\n${title}\n${message}`);
    return;
  }

  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('safeAlert: UI表示をスキップしました: ' + e.message);
    Logger.log(`${title}\n${message}`);
  }
}

// 手動実行（最新行データ）
function runManually() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFormResponseSheet(ss);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('データがありません。フォームから回答を入力してください。');
    return;
  }

  const data = getRowDataByHeader(sheet, lastRow);
  const clientName = data.clientName || '（未入力）';

  const result = ui.alert(
    '手動生成',
    `最新行（行${lastRow}）のデータでスライドを生成します。\n\n顧客名: ${clientName}\n\n実行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    onFormSubmit();
    const outCol = ensureOutputColumns(sheet);
    const url = sheet.getRange(lastRow, outCol.slideUrlCol).getValue();
    safeAlert('完了', `スライドが生成されました。\n\n${url}`);
  }
}
