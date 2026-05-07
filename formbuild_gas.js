/**
 * Sol事例紹介ページ 入力フォーム v3（1ページ版）
 */

const FORM_TITLE_V3   = 'Sol事例紹介ページ 入力フォーム';
const SHEET_RESPONSES_V3 = '回答データ';
const SHEET_LOG_V3       = '実行ログ';

const STAFF_LIST_V3 = [
  '上田いばら','對馬綾香','奥田快','田口徹','高橋秀吾',
  '井上翔太','山口岳人','藤井絵梨奈','宮本菜央','圡井富夏城','山中柊弥','その他',
];

const FIELD_MAP_V3 = {
  'タイムスタンプ':                     'timestamp',
  'メールアドレス':                     'email',
  '担当者名':                           'staff_name',
  '顧客企業名':                         'company_name',
  '業種':                               'industry',
  '従業員数':                           'employees',
  '本社所在地':                         'location',
  '顧客企業のWebサイトURL':             'website_url',
  '導入したCloudCIRCUS製品':            'cc_product',
  '導入効果を一言で（キャッチコピー）': 'catchcopy',
  '課題①':                             'challenge_1',
  '課題②（任意）':                     'challenge_2',
  '解決策':                             'solution',
  '導入効果':                           'result_summary',
  '顧客の声（引用コメント）':           'customer_voice',
  'KGI・最大の成果':                   'kgi_text',
  'KPI①の項目名と数値':                'kpi1',
  'KPI②の項目名と数値（任意）':        'kpi2',
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = createForm_();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  normalizeSheet_(ss);
  setupLogSheet_(ss);
  const url = form.getPublishedUrl();
  writeLog_(ss, 'SETUP', 'セットアップ完了', url);
  SpreadsheetApp.getUi().alert('✅ セットアップ完了！\n\n回答用URL:\n' + url);
}

function createForm_() {
  const form = FormApp.create(FORM_TITLE_V3);
  form.setTitle(FORM_TITLE_V3);
  form.setDescription(
    '事例紹介スライドの自動生成に使用します。\n' +
    '★付きは必須、それ以外は任意（空欄OK）です。\n' +
    '入力内容はCloudCIRCUS社内でのみ使用します。'
  );
  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage('ご入力ありがとうございました！数営業日以内にスライドをお送りします。');

  form.addListItem()
    .setTitle('担当者名')
    .setRequired(true)
    .setChoiceValues(STAFF_LIST_V3);

  form.addTextItem()
    .setTitle('顧客企業名')
    .setRequired(true)
    .setHelpText('例：株式会社テックソリューションズ　※「様」不要');

  form.addListItem()
    .setTitle('業種')
    .setRequired(true)
    .setChoiceValues(['情報通信業','製造業','建設業','小売業','医療・福祉','不動産業','金融・保険業','教育・学習支援業','サービス業','その他']);

  form.addTextItem()
    .setTitle('従業員数')
    .setRequired(true)
    .setHelpText('例：250名');

  form.addTextItem()
    .setTitle('本社所在地')
    .setRequired(true)
    .setHelpText('例：東京都千代田区');

  form.addTextItem()
    .setTitle('顧客企業のWebサイトURL')
    .setRequired(true)
    .setHelpText('例：https://example.co.jp');

  form.addCheckboxItem()
    .setTitle('導入したCloudCIRCUS製品')
    .setRequired(true)
    .setChoiceValues(['BowNow','AIチャットボット','BlueMonkey','Actibook','Sitekit','その他']);

  form.addTextItem()
    .setTitle('導入効果を一言で（キャッチコピー）')
    .setRequired(true)
    .setHelpText('例：問い合わせ数が3ヶ月でフォーム比3倍に');

  form.addParagraphTextItem()
    .setTitle('課題①')
    .setRequired(true)
    .setHelpText('導入前に抱えていた最大の課題を1〜2文で。');

  form.addParagraphTextItem()
    .setTitle('課題②（任意）')
    .setRequired(false)
    .setHelpText('2つ目の課題があれば。なければ空欄でOK。');

  form.addParagraphTextItem()
    .setTitle('解決策')
    .setRequired(true)
    .setHelpText('CloudCIRCUSの製品・サービスでどう解決したか1〜2文で。');

  form.addParagraphTextItem()
    .setTitle('導入効果')
    .setRequired(true)
    .setHelpText('数値を交えて具体的に。例：導入3ヶ月でチャットボット経由の問い合わせがフォーム比約3倍に');

  form.addParagraphTextItem()
    .setTitle('顧客の声（引用コメント）')
    .setRequired(true)
    .setHelpText('お客様の言葉をそのまま引用してください。');

  form.addTextItem()
    .setTitle('KGI・最大の成果')
    .setRequired(true)
    .setHelpText('例：商談獲得数が目標の130%達成　例：月間CV数が前年比2倍');

  form.addTextItem()
    .setTitle('KPI①の項目名と数値')
    .setRequired(true)
    .setHelpText('「項目名：数値」の形式で。例：新規問い合わせ件数：29件');

  form.addTextItem()
    .setTitle('KPI②の項目名と数値（任意）')
    .setRequired(false)
    .setHelpText('例：ROAS：350%　例：LTV：120,000円');

  return form;
}

function normalizeSheet_(ss) {
  Utilities.sleep(1000);
  const sheets = ss.getSheets();
  let sheet = sheets.find(s => s.getName().startsWith('フォームの回答')) || sheets[sheets.length - 1];
  sheet.setName(SHEET_RESPONSES_V3);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers.map(h => FIELD_MAP_V3[h] || h)]);
  const statusCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, statusCol).setValue('status');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, statusCol)
    .setBackground('#1F497D').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.autoResizeColumns(1, statusCol);
}

function setupLogSheet_(ss) {
  const s = ss.getSheetByName(SHEET_LOG_V3) || ss.insertSheet(SHEET_LOG_V3);
  s.clearContents();
  s.getRange(1,1,1,4).setValues([['日時','イベント','メッセージ','詳細']])
    .setBackground('#1F497D').setFontColor('#FFFFFF').setFontWeight('bold');
  s.setFrozenRows(1);
  [160,100,280,320].forEach((w,i) => s.setColumnWidth(i+1, w));
}

function writeLog_(ss, event, msg, detail) {
  const s = ss.getSheetByName(SHEET_LOG_V3);
  if (s) s.appendRow([new Date().toLocaleString('ja-JP'), event, msg, detail || '']);
}

function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const data = { email: e.response.getRespondentEmail() || '' };
    e.response.getItemResponses().forEach(ir => {
      const key = FIELD_MAP_V3[ir.getItem().getTitle()] || ir.getItem().getTitle();
      data[key] = ir.getResponse() || '';
    });
    writeLog_(ss, 'SUBMIT', '回答受付: ' + (data.company_name || ''), data.email);
    // generatePptx_(data); ← PPT自動生成をここに追加予定
  } catch(err) {
    writeLog_(ss, 'ERROR', err.message, '');
  }
}