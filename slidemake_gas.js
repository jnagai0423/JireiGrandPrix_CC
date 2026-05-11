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

    const data = buildDataObject_(e);
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
  const problem = d.slide3_problem || d.challenge_1 || d.challenge || '';
  const proposal = d.slide3_proposal || d.solution || '';

  const values = {
    company_name: d.company_name,
    catchcopy: d.catchcopy,
    highlight_point: d.highlight_point,
    website_url: d.website_url,
    industry: d.industry,
    employees: d.employees,
    location: d.location,
    challenge_1: d.challenge_1 || d.challenge || '',
    challenge_2: '',
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
  // 正式キーは英字 snake_case のみ。
  return [
    ['{{company_name}}', v.company_name],
    ['{{catchcopy}}', v.catchcopy],
    ['{{highlight_point}}', v.highlight_point],
    ['{{website_url}}', v.website_url],
    ['{{industry}}', v.industry],
    ['{{employees}}', v.employees],
    ['{{location}}', v.location],
    ['{{challenge}}', v.challenge_1],
    ['{{solution}}', v.solution],
    ['{{result_summary}}', v.result_summary],
    ['{{cc_product}}', v.cc_product],
    ['{{slide3_problem}}', v.slide3_problem],
    ['{{slide3_solution}}', v.slide3_solution],
    ['{{kgi_text}}', v.kgi_text],
    ['{{kpi1_label}}', v.kpi1_label],
    ['{{kpi1_value}}', v.kpi1_value],
    ['{{kpi2_label}}', v.kpi2_label],
    ['{{kpi2_value}}', v.kpi2_value],
    ['{{ai_usage}}', v.ai_usage],
    ['{{ai_before}}', v.ai_before],
    ['{{ai_after}}', v.ai_after],
  ];
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
  const data = { email: e.response.getRespondentEmail() || '' };
  e.response.getItemResponses().forEach(ir => {
    const key = FIELD_MAP_SLIDES[ir.getItem().getTitle()] || ir.getItem().getTitle();
    const val = ir.getResponse();
    data[key] = Array.isArray(val) ? val.join(', ') : (val || '');
  });
  if (!data.slide3_problem) data.slide3_problem = data.challenge_1 || data.challenge || '';
  if (!data.slide3_proposal) data.slide3_proposal = data.solution || '';
  return data;
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
  const sheet = getResponseSheet_(ss);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('回答シートにデータがありません。');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = {};
  headers.forEach((h, i) => {
    const rawKey = h.toString().trim();
    const key = FIELD_MAP_SLIDES[rawKey] || rawKey;
    if (key === 'ステータス' || key === 'status' || key === 'slide_url') return;
    data[key] = String(lastRow[i] || '');
  });

  if (!data.slide3_problem) data.slide3_problem = data.challenge_1 || data.challenge || '';
  if (!data.slide3_proposal) data.slide3_proposal = data.solution || '';

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
