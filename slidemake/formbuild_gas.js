/**
 * Sol事例紹介ページ 入力フォーム v4
 * v3をベースに以下を追加：
 * - 顧客の課題・お悩み/伸びしろ（スライド3用・任意）→ 空欄時は「抱えていた課題」から自動流用
 * - CC社からの提案/施策（スライド3用・任意）→ 空欄時は解決策から自動流用
 * - AI活用（スライド5用・任意3問）→ すべて空欄時はスライド5を省略
 */

const FORM_TITLE_V3 = 'Sol事例紹介ページ 入力フォーム';
const SHEET_RESPONSES_V3 = '回答';
const SHEET_LOG_V3 = '実行ログ';
const OUTPUT_FOLDER_ID_V3 = '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e';
const PROP_FORM_ID_V3 = 'FORM_ID_V3';

const STAFF_LIST_V3 = [
  '上田いばら', '對馬綾香', '奥田快', '田口徹', '高橋秀吾',
  '井上翔太', '山口岳人', '藤井絵梨奈', '宮本菜央', '圡井富夏城', '山中柊弥', 'その他'
];

const FIELD_MAP_V3 = {
  'タイムスタンプ': 'timestamp',
  'メールアドレス': 'email',
  'あなたのお名前を入力してください': 'staff_name',
  '顧客企業名を入力してください': 'company_name',
  '顧客企業の業種を選択してください': 'industry',
  '顧客企業の従業員数を入力してください': 'employees',
  '顧客企業の本社所在地を入力してください': 'location',
  '顧客企業のWebサイトURLを入力してください': 'website_url',
  '顧客企業のクラウドサーカス導入製品を選択してください': 'cc_product',
  '成果を「一言」で記載してください': 'catchcopy',
  '抱えていた課題': 'challenge',
  '解決手法': 'solution',
  '得られた成果': 'result_summary',
  '注目ポイント（特にアピールしたいこと）を記載してください': 'highlight_point',
  '顧客にとっての最終目標（KGI）を記載してください': 'kgi_text',
  'KPI①の「項目名」と「数値」を記載してください': 'kpi1',
  'KPI②の「項目名」と「数値」を記載してください（あれば・任意）': 'kpi2',
  '顧客企業が抱えていた課題についてより詳しく記載ください': 'slide3_problem',
  'CC社からの解決手法についてより詳しく記載ください': 'slide3_proposal',
  '解決にあたってどの部分にAIを使用したかを記載ください（あれば・任意）': 'ai_usage',
  'AI活用前（あれば・任意）': 'ai_before',
  'AI活用後（あれば・任意）': 'ai_after'
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('setup: start');
  const form = getOrCreateForm_();
  removeLegacyFormItems_(form);
  Logger.log('setup: form ready');
  // フォーム自体も指定フォルダへ移動（既存フォームでも揃える）
  moveFileToOutputFolder_(form.getId(), 'form');
  moveFileToOutputFolder_(ss.getId(), 'spreadsheet');
  Logger.log('setup: files moved to output folder');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Logger.log('setup: destination set');
  normalizeSheet_(ss);
  Logger.log('setup: sheet normalized');
  setupLogSheet_(ss);
  Logger.log('setup: log sheet ready');
  const url = form.getPublishedUrl();
  writeLog_(ss, 'SETUP', 'セットアップ完了', url);
  Logger.log(url);
  Logger.log('setup: done');
  return url;
}

function getOrCreateForm_() {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty(PROP_FORM_ID_V3);
  if (formId) {
    try {
      const existingForm = FormApp.openById(formId);
      return existingForm;
    } catch (err) {
      Logger.log('stored form not found, recreate: ' + err.message);
    }
  }

  const form = createForm_();
  moveFileToOutputFolder_(form.getId(), 'form');
  props.setProperty(PROP_FORM_ID_V3, form.getId());
  return form;
}

function applyHeaderLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = getOrCreateForm_();
  removeLegacyFormItems_(form);
  // 手動実行時もフォーム配置を揃える
  moveFileToOutputFolder_(form.getId(), 'form');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  normalizeSheet_(ss);
  setupLogSheet_(ss);
  writeLog_(ss, 'MANUAL', 'ヘッダーレイアウト適用', '');
}

/**
 * 現在このGASが紐づけている「正しいフォーム」のID/URLをログに出します。
 * Drive上の不要フォーム削除の判断材料にしてください。
 */
function debugFormInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = getOrCreateForm_();
  const url = form.getPublishedUrl();
  // detailにURL、msgにIDを出す
  writeLog_(ss, 'FORM_INFO', form.getId(), url);
}

function removeLegacyFormItems_(form) {
  const legacyTitles = [
    '抱えていた課題②（あれば・任意）'
  ];
  for (let i = form.getItems().length - 1; i >= 0; i--) {
    const item = form.getItems()[i];
    const title = item.getTitle && item.getTitle();
    if (legacyTitles.includes(title)) {
      form.deleteItem(i);
      Logger.log('legacy item removed: ' + title);
    }
  }
}

/**
 * フォームIDを初期化して新規フォームを作り直します。
 * 「実行は成功するがフォームが見つからない」時の復旧用。
 */
function recreateFormAndSetup() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_FORM_ID_V3);
  return setup();
}

function moveFileToOutputFolder_(fileId, label) {
  try {
    const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID_V3);
    const file = DriveApp.getFileById(fileId);
    file.moveTo(folder);
    Logger.log('move success: ' + label + ' -> ' + folder.getName());
  } catch (err) {
    Logger.log('move failed: ' + label + ' / ' + err.message);
  }
}

function createForm_() {
  const form = FormApp.create(FORM_TITLE_V3);
  form.setTitle(FORM_TITLE_V3);
  form.setDescription(
    '事例紹介スライドの自動生成に使用します。\n' +
    '★付きは必須、それ以外は任意（空欄OK）です。\n' +
    '入力した内容は、CC社内でのみ使用します。'
  );
  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage('ご入力ありがとうございました！数営業日以内にスライドをお送りします。');

  form.addListItem()
    .setTitle('あなたのお名前を入力してください')
    .setRequired(true)
    .setChoiceValues(STAFF_LIST_V3);

  form.addTextItem()
    .setTitle('顧客企業名を入力してください')
    .setRequired(true)
    .setHelpText('例：株式会社ユニリタ（敬称略）');

  form.addListItem()
    .setTitle('顧客企業の業種を選択してください')
    .setRequired(true)
    .setChoiceValues([
      '製造業（繊維/化学/医薬/食品/資源/ゴム/硝子/鉄鋼/金属/機械/電機）',
      '情報通信業（SIer/SaaS開発/インターネット）',
      '金融・保険業（銀行/証券/商品先物/保険/その他金融）',
      'サービス業（専門サービス/広告/人材ビジネス/娯楽/エンタテインメント/その他）',
      '運輸・郵便業（海運・空運・陸運・倉庫）',
      '医療・福祉業',
      '卸売・小売業',
      '不動産業',
      '建設業',
      '宿泊・飲食業',
      '教育・学習支援業',
      '農業・林業・漁業・鉱業',
      '電気・ガス・水道業',
      '公共（地方自治体/公共団体/非営利団体）',
      'その他'
    ]);

  form.addTextItem()
    .setTitle('顧客企業の従業員数を入力してください')
    .setRequired(false)
    .setHelpText('例：200名');

  form.addTextItem()
    .setTitle('顧客企業の本社所在地を入力してください')
    .setRequired(false)
    .setHelpText('例：東京都新宿区');

  form.addTextItem()
    .setTitle('顧客企業のWebサイトURLを入力してください')
    .setRequired(true)
    .setHelpText('例：https://cloudcircus.jp/');

  form.addCheckboxItem()
    .setTitle('顧客企業のクラウドサーカス導入製品を選択してください')
    .setRequired(true)
    .setChoiceValues([
      'BowNow',
      'BlueMonkey',
      'Actibook',
      'Plusdb',
      'IZANAI',
      'FullStar',
      'なし'
    ])
    .showOtherOption(true);

  form.addTextItem()
    .setTitle('成果を「一言」で記載してください')
    .setRequired(true)
    .setHelpText('例：問い合わせ数が3ヶ月で3倍に、etc');

  form.addParagraphTextItem()
    .setTitle('抱えていた課題')
    .setRequired(true)
    .setHelpText('顧客企業が抱えていた課題を「簡潔」に記載ください');

  form.addParagraphTextItem()
    .setTitle('解決手法')
    .setRequired(true)
    .setHelpText('コンサルティング/BPOサービスでどう解決したか「簡潔」に記載ください');

  form.addParagraphTextItem()
    .setTitle('得られた成果')
    .setRequired(true)
    .setHelpText('数値を交えて「簡潔」に記載ください、例：支援3ヶ月で問い合わせが約3倍に、etc');

  form.addParagraphTextItem()
    .setTitle('注目ポイント（特にアピールしたいこと）を記載してください')
    .setRequired(true)
    .setHelpText('今回の事例で特に伝えたいポイントを具体的に記載してください');

  form.addTextItem()
    .setTitle('顧客にとっての最終目標（KGI）を記載してください')
    .setRequired(false)
    .setHelpText('例：商談獲得数が目標の130%達成、月間CV数が前年比2倍に、etc');

  form.addTextItem()
    .setTitle('KPI①の「項目名」と「数値」を記載してください')
    .setRequired(true)
    .setHelpText('「項目名：数値」の形式で、例：新規問い合わせ件数：29件、etc');

  form.addTextItem()
    .setTitle('KPI②の「項目名」と「数値」を記載してください（あれば・任意）')
    .setRequired(false)
    .setHelpText('例：ROAS：350%、例：LTV：120,000円');

  form.addParagraphTextItem()
    .setTitle('顧客企業が抱えていた課題についてより詳しく記載ください')
    .setRequired(false)
    .setHelpText('空欄の場合は課題のテキストを自動で表記します');

  form.addParagraphTextItem()
    .setTitle('CC社からの解決手法についてより詳しく記載ください')
    .setRequired(false)
    .setHelpText('空欄の場合は解決手法のテキストを自動で表記します');

  form.addParagraphTextItem()
    .setTitle('解決にあたってどの部分にAIを使用したかを記載ください（あれば・任意）')
    .setRequired(false)
    .setHelpText('AI活用がある場合のみ記載してください');

  form.addParagraphTextItem()
    .setTitle('AI活用前（あれば・任意）')
    .setRequired(false)
    .setHelpText('AI活用前の状態を記載してください');

  form.addParagraphTextItem()
    .setTitle('AI活用後（あれば・任意）')
    .setRequired(false)
    .setHelpText('AI活用後の変化を記載してください');

  return form;
}

function normalizeSheet_(ss) {
  Logger.log('normalizeSheet_: start');
  const existing = ss.getSheetByName(SHEET_RESPONSES_V3);
  const detected = findResponseSheet_(ss);
  const sheet = (existing && detected.getSheetId() !== existing.getSheetId()) ? existing : detected;

  // フォーム連携シートは削除できないため、削除は行わず「回答」シートを優先して使う
  if (sheet.getName() !== SHEET_RESPONSES_V3) {
    sheet.setName(SHEET_RESPONSES_V3);
  }
  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) {
    Logger.log('normalizeSheet_: no headers found (lastCol<=0). skip header mapping.');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const statusCol = sheet.getLastColumn() + 1;
  if (!headers.includes('ステータス') && !headers.includes('status')) {
    sheet.getRange(1, statusCol).setValue('ステータス');
  }
  sheet.setFrozenRows(0);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground(null)
    .setFontColor(null)
    .setFontWeight('normal');
  Logger.log('normalizeSheet_: done');
}

function findResponseSheet_(ss) {
  const existing = ss.getSheetByName(SHEET_RESPONSES_V3);
  const sheets = ss.getSheets();
  const formNamedSheets = sheets.filter(s =>
    s.getName().startsWith('回答リスト') ||
    s.getName().startsWith('フォームの回答') ||
    s.getName().startsWith('Form Responses') ||
    s.getName().startsWith('Form_Responses')
  );
  const latestFormNamedSheet = formNamedSheets.length > 0 ? formNamedSheets[formNamedSheets.length - 1] : null;
  const looksLikeResponseSheet = s => {
    const lastCol = s.getLastColumn();
    if (lastCol <= 0) return false;
    const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
    return headers.some(h => Object.prototype.hasOwnProperty.call(FIELD_MAP_V3, h));
  };

  if (latestFormNamedSheet && (!existing || latestFormNamedSheet.getSheetId() !== existing.getSheetId())) {
    return latestFormNamedSheet;
  }

  return existing ||
    sheets.find(looksLikeResponseSheet) ||
    sheets[sheets.length - 1];
}

function setupLogSheet_(ss) {
  const s = ss.getSheetByName(SHEET_LOG_V3) || ss.insertSheet(SHEET_LOG_V3);
  s.clearContents();
  s.getRange(1, 1, 1, 4).setValues([['日時', 'イベント', 'メッセージ', '詳細']])
    .setBackground('#1F497D').setFontColor('#FFFFFF').setFontWeight('bold');
  s.setFrozenRows(1);
}

function writeLog_(ss, event, msg, detail) {
  const s = ss.getSheetByName(SHEET_LOG_V3);
  if (s) s.appendRow([new Date().toLocaleString('ja-JP'), event, msg, detail || '']);
}