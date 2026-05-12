/**
 * Sol事例紹介ページ Google Slides自動生成 最終版
 * ・ポップアップなし（Logger.logのみ）
 * ・ヘッダーが日本語・field_id どちらでも動作
 */

const TEMPLATE_SLIDE_ID = '1zlkbDsFCQgXuXvoAIkjkybOozpO8CquaBRpvgagSf2k';
const OUTPUT_FOLDER_ID = '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e';
const SHEET_RESPONSES = '回答';
const SHEET_LOG = '実行ログ';

const FIELD_MAP_SLIDES = {
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
  '顧客企業が抱えていた課題を「簡潔」に記載してください': 'challenge',
  '※課題をどのように解決したか「簡潔」に記載してください': 'solution',
  '特に、得られた成果を「数値」で記載してください': 'result_summary',
  '特にアピールしたいこと（注目ポイント）を記載してください': 'highlight_point',
  '顧客企業が抱えていた課題についてより「詳しく」記載してください': 'slide3_problem',
  'CC社からの解決手法についてより「詳しく」記載してください': 'slide3_proposal',
  'どの部分にAIを使用したかを記載ください（あれば・任意）': 'ai_usage',

  // 旧フォーム文言（互換）
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
  'AI活用後（あれば・任意）': 'ai_after',
};

const FIELD_ALIASES_SLIDES = {
  // 入力データ側の互換対応（最終的には正規キーへ寄せる）
  challenge_1: 'challenge',
  slide3_proposal: 'slide3_proposal',
  result: 'result_summary',
};

const PLACEHOLDER_KEY_MAP = {
  company_name: ['{{company_name}}'],
  catchcopy: ['{{catchcopy}}'],
  highlight_point: ['{{highlight_point}}'],
  website_url: ['{{website_url}}'],
  industry: ['{{industry}}'],
  employees: ['{{employees}}'],
  location: ['{{location}}'],
  challenge: ['{{challenge}}'],
  solution: ['{{solution}}'],
  result_summary: ['{{result_summary}}'],
  cc_product: ['{{cc_product}}'],
  slide3_problem: ['{{slide3_problem}}', '{{slide3_Problem}}'], // 旧テンプレ互換
  slide3_solution: ['{{slide3_solution}}'],
  kgi_text: ['{{kgi_text}}'],
  kpi1_label: ['{{kpi1_label}}'],
  kpi1_value: ['{{kpi1_value}}'],
  kpi2_label: ['{{kpi2_label}}'],
  kpi2_value: ['{{kpi2_value}}'],
  ai_usage: ['{{ai_usage}}'],
  ai_before: ['{{ai_before}}'],
  ai_after: ['{{ai_after}}'],
};

// ============================================================
// フォーム回答時に自動実行
// ============================================================

function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // 念のため、送信直後に回答シートを「回答」に揃えておく
    if (typeof normalizeSheet_ === 'function') {
      const responseSheet = normalizeSheet_(ss);
      if (typeof setupDisplaySheet_ === 'function' && responseSheet) {
        setupDisplaySheet_(ss, responseSheet);
      }
    }

    let data = buildDataObject_(e);
    // トリガーイベントに必要項目が乗らない場合は、回答シート最終行から補完
    if (!data.company_name || !data.challenge || !data.solution) {
      const fallback = buildDataFromLastRow_(ss);
      data = Object.assign({}, fallback, data);
    }
    writeLog_(ss, 'DEBUG', 'data keys', Object.keys(data).join(','));
    writeLog_(ss, 'DEBUG', 'company/challenge', (data.company_name || '') + ' / ' + (data.challenge || ''));
    const slideUrl = generateSlide_(data);
    updateStatus_(ss, '✅ 完了', slideUrl);
    writeLog_(ss, 'SUCCESS', data.company_name + ' 生成完了', slideUrl);
  } catch (err) {
    writeLog_(ss, 'ERROR', err.message, '');
    updateStatus_(ss, '❌ エラー', err.message);
  }
}

// ============================================================
// スライド生成メイン
// ============================================================

function generateSlide_(d) {
  const newFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID)
    .makeCopy(
      d.company_name + '様_事例紹介_' + formatDate_(new Date()),
      DriveApp.getFolderById(OUTPUT_FOLDER_ID)
    );
  const newId = newFile.getId();

  const kpi1 = splitKpi_(d.kpi1);
  const kpi2 = splitKpi_(d.kpi2);
  const problem = d.slide3_problem || d.challenge || '';
  const proposal = d.slide3_proposal || d.solution || '';

  const values = {
    company_name: d.company_name,
    catchcopy: d.catchcopy,
    highlight_point: d.highlight_point,
    website_url: d.website_url,
    industry: d.industry,
    employees: d.employees,
    location: d.location,
    challenge: d.challenge || '',
    solution: d.solution,
    result_summary: d.result_summary,
    cc_product: d.cc_product,
    slide3_problem: problem,
    slide3_solution: proposal,
    kgi_text: d.kgi_text,
    kpi1_label: kpi1.label,
    kpi1_value: kpi1.value,
    kpi2_label: kpi2.label,
    kpi2_value: kpi2.value,
    ai_usage: d.ai_usage || '',
    ai_before: d.ai_before || '',
    ai_after: d.ai_after || '',
  };

  const replacements = buildReplacements_(values);

  const presentation = SlidesApp.openById(newId);
  replacements.forEach(([from, to]) => {
    presentation.replaceAllText(from, to || ' ');
  });

  const hasAi = d.ai_usage || d.ai_before || d.ai_after;
  if (!hasAi) {
    const slides = presentation.getSlides();
    if (slides[4]) {
      slides[4].remove();
    }
  }
  presentation.saveAndClose();
  return 'https://docs.google.com/presentation/d/' + newId;
}

function buildReplacements_(v) {
  const replacements = [];
  Object.keys(PLACEHOLDER_KEY_MAP).forEach(key => {
    const placeholders = PLACEHOLDER_KEY_MAP[key];
    placeholders.forEach(ph => {
      replacements.push([ph, v[key] || '']);
    });
  });
  return replacements;
}

// ============================================================
// ユーティリティ
// ============================================================

function splitKpi_(kpiStr) {
  if (!kpiStr) return { label: '', value: '' };
  for (const sep of ['：', ':', '＝', '=']) {
    if (kpiStr.includes(sep)) {
      const parts = kpiStr.split(sep);
      return { label: parts[0].trim(), value: parts[1].trim() };
    }
  }
  return { label: kpiStr.trim(), value: '' };
}

function buildDataObject_(e) {
  const data = { email: '' };

  // Form trigger (イベントソース: フォームから) の payload
  if (e && e.response && typeof e.response.getItemResponses === 'function') {
    data.email = e.response.getRespondentEmail() || '';
    e.response.getItemResponses().forEach(ir => {
      const key = resolveFieldKey_(ir.getItem().getTitle());
      const val = ir.getResponse();
      data[key] = Array.isArray(val) ? val.join(', ') : (val || '');
    });
  // Spreadsheet trigger (イベントソース: スプレッドシートから) の payload
  } else if (e && e.namedValues) {
    Object.keys(e.namedValues).forEach(title => {
      const key = resolveFieldKey_(title);
      const val = e.namedValues[title];
      data[key] = Array.isArray(val) ? val.join(', ') : String(val || '');
    });
    data.email = data.email || '';
  } else {
    throw new Error('フォーム送信イベントの形式を解釈できませんでした。トリガー設定を確認してください。');
  }

  if (!data.slide3_problem) data.slide3_problem = data.challenge || '';
  if (!data.slide3_proposal) data.slide3_proposal = data.solution || '';
  return data;
}

function buildDataFromLastRow_(ss) {
  const sheet = getResponseSheet_(ss);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = {};
  headers.forEach((h, i) => {
    const rawKey = String(h || '').trim();
    const key = resolveFieldKey_(rawKey);
    if (key === 'ステータス' || key === 'status' || key === 'slide_url') return;
    data[key] = String(lastRow[i] || '');
  });

  if (!data.slide3_problem) data.slide3_problem = data.challenge || '';
  if (!data.slide3_proposal) data.slide3_proposal = data.solution || '';
  return data;
}

function resolveFieldKey_(rawKey) {
  const key = String(rawKey || '').trim();
  if (!key) return key;
  if (Object.prototype.hasOwnProperty.call(FIELD_MAP_SLIDES, key)) return FIELD_MAP_SLIDES[key];
  if (Object.prototype.hasOwnProperty.call(FIELD_ALIASES_SLIDES, key)) return FIELD_ALIASES_SLIDES[key];

  const normalized = key
    .toLowerCase()
    .replace(/[ 　\t\n\r]/g, '')
    .replace(/[()（）「」『』【】]/g, '');

  if (normalized.includes('抱えていた課題') && normalized.includes('詳')) return 'slide3_problem';
  if ((normalized.includes('解決手法') || normalized.includes('提案施策')) && normalized.includes('詳')) return 'slide3_proposal';
  if (normalized.includes('抱えていた課題')) return 'challenge';
  if (normalized.includes('解決手法') || normalized.includes('提案施策')) return 'solution';
  if (normalized.includes('得られた成果')) return 'result_summary';
  if (normalized.includes('注目ポイント')) return 'highlight_point';
  if (normalized.includes('どの部分にaiを使用') || normalized.includes('aiを使用')) return 'ai_usage';
  if (normalized.includes('ai活用前')) return 'ai_before';
  if (normalized.includes('ai活用後')) return 'ai_after';

  return key;
}

function updateStatus_(ss, status, slideUrl) {
  const sheet = getResponseSheet_(ss);
  if (!sheet || sheet.getLastRow() < 2) return;
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  let statusCol = headers.findIndex(h => String(h).trim() === 'ステータス' || String(h).trim() === 'status') + 1;
  if (statusCol === 0) {
    statusCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, statusCol).setValue('ステータス');
  }

  let slideUrlCol = headers.findIndex(h => String(h).trim() === 'slide_url') + 1;
  if (slideUrlCol === 0) {
    slideUrlCol = Math.max(sheet.getLastColumn(), statusCol) + 1;
    sheet.getRange(1, slideUrlCol).setValue('slide_url');
  }

  sheet.getRange(lastRow, statusCol).setValue(status);
  sheet.getRange(lastRow, slideUrlCol).setValue(slideUrl || '');
}

function writeLog_(ss, event, msg, detail) {
  const s = ss.getSheetByName(SHEET_LOG);
  if (s) s.appendRow([new Date().toLocaleString('ja-JP'), event, msg, detail || '']);
}

function formatDate_(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + m + d;
}

// ============================================================
// 手動テスト用（日本語ヘッダー・field_id 両対応）
// ============================================================

function testWithLastRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = buildDataFromLastRow_(ss);
  if (!data || Object.keys(data).length === 0) {
    Logger.log('回答シートにデータがありません。');
    return;
  }

  Logger.log('取得データ: ' + JSON.stringify(data));

  try {
    const slideUrl = generateSlide_(data);
    updateStatus_(ss, '✅ 完了（テスト）', slideUrl);
    writeLog_(ss, 'TEST', (data.company_name || 'テスト') + ' 生成完了', slideUrl);
    Logger.log('✅ 生成完了: ' + slideUrl);
  } catch (err) {
    Logger.log('❌ エラー: ' + err.message);
    writeLog_(ss, 'ERROR', err.message, '');
  }
}

function getResponseSheet_(ss) {
  const preferred = ss.getSheetByName(SHEET_RESPONSES);
  if (preferred && preferred.getLastRow() >= 2) return preferred;

  const sheets = ss.getSheets();
  const formNamedSheets = sheets.filter(s =>
    s.getName().startsWith('フォームの回答') ||
    s.getName().startsWith('Form Responses') ||
    s.getName().startsWith('Form_Responses') ||
    s.getName().startsWith('回答リスト')
  );
  if (formNamedSheets.length > 0) {
    // 複数ある場合は、回答行（2行目以降）が最も多いシートを優先
    const withRows = formNamedSheets
      .map(s => ({ sheet: s, dataRows: Math.max(0, s.getLastRow() - 1) }))
      .sort((a, b) => b.dataRows - a.dataRows);
    if (withRows[0] && withRows[0].dataRows > 0) return withRows[0].sheet;
    return formNamedSheets[formNamedSheets.length - 1];
  }

  // 最後のフォールバック
  if (preferred) return preferred;
  return ss.getSheets()[0] || null;
}
