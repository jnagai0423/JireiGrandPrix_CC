// ヘッダー名ベースで行データを取得（フォーム項目変更に強くする）
function getRowDataByHeader(sheet, rowNumber) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valueRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));
  /** 複数列に部分一致する場合は、最も長いキーに一致した列を採用（短い「url」だけの誤マッチを減らす） */
  const getByKey = (keys, excludeKeys) => {
    const normalizedKeys = (keys || []).map(k => normalizeHeader(k)).filter(Boolean);
    const normalizedExcludes = (excludeKeys || []).map(k => normalizeHeader(k)).filter(Boolean);

    let bestIdx = -1;
    let bestScore = 0;

    normalizedHeaders.forEach((h, i) => {
      if (!h) return;
      if (normalizedExcludes.some(ex => ex && h.includes(ex))) return;

      let score = 0;
      for (const k of normalizedKeys) {
        if (h.includes(k) && k.length > score) score = k.length;
      }
      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    });

    if (bestIdx === -1 || bestScore === 0) return '';
    return String(valueRow[bestIdx] || '').trim();
  };

  const personName = getByKey([
    '自分の名前を入力してください',
    '自分の名前を入力',
    '自分の名前',
    '担当者名',
    '氏名',
  ]);
  const clientName = getByKey([
    '募集企業名および製品サービス名',
    '募集企業名',
    '製品サービス名',
    '顧客企業名を正式名称で入力',
    '顧客企業名',
    '顧客名',
  ]);
  const industry = getByKey([
    '顧客企業の業種を以下から選択してください',
    '顧客企業の業種',
    '業種',
  ]);
  const products = getByKey(
    [
      '導入済みのcloudcircus製品があれば選択してください',
      '導入済みのcloudcircus製品',
      '導入済のcloudcircus製品',
      'cloudcircus製品',
      'cloudcircus',
      '導入済',
    ],
    ['顧客企業名', '顧客名', '募集企業名', '生成スライドurl']
  );
  const plan = getByKey([
    '期待するコンサルティングの成果',
    '運用中のコンサルティングプラン',
    'コンサルティングの成果',
    'コンサルティングプラン',
    '運用中',
    'プラン',
    'コンサルティング',
  ]);
  const siteUrl = getByKey(
    [
      '顧客企業の対象url',
      '対象url',
      'サイトurlを入力',
      'サイトurl',
      'webサイトurl',
      'ホームページurl',
      'url',
    ],
    ['生成スライドurl', 'メールアドレス']
  );
  const genre = getByKey([
    '成果事例のジャンルを以下から選択してください',
    '成果事例のジャンル',
    'ジャンル',
  ]);
  const feature = getByKey([
    '成果事例の成果を一言でアピール',
    '成果事例の成果を一言で',
    '成果事例の成果',
    '成果事例の特徴を一言で表すと',
    '成果事例の特徴',
    '15文字',
  ]);
  const detail = getByKey([
    '成果事例の内容をできるだけ詳細に記述してください',
    '成果事例の内容',
    '300文字',
    '詳細',
  ]);
  const metric30DayByHeader = getByKey([
    'kpi改善数',
    'KPI改善数',
    '30日内数値',
    '30日内',
  ]);
  const actualMetricByHeader = getByKey([
    'kpi実績数値',
    'KPI実績数値',
    '実績数値',
    '実績数値あれば',
  ]);
  // L/M 列を明示フォールバック（列追加や文言変更で見出しマッチしない時の保険）
  const metric30Day =
    metric30DayByHeader || (sheet.getLastColumn() >= 12 ? String(valueRow[11] || '').trim() : '');
  const actualMetric =
    actualMetricByHeader || (sheet.getLastColumn() >= 13 ? String(valueRow[12] || '').trim() : '');

  const data = {
    personName,
    clientName,
    industry,
    products,
    siteUrl,
    genre,
    feature,
    detail,
    plan,
    metric30Day,
    actualMetric,
  };

  if (CONFIG.DEBUG_SHEET_HEADERS) {
    Logger.log('[DEBUG] シート: ' + sheet.getName() + ' 行: ' + rowNumber);
    Logger.log('[DEBUG] ヘッダー: ' + JSON.stringify(headerRow));
    Logger.log('[DEBUG] 読み取り: ' + JSON.stringify(data));
  }

  return data;
}

function normalizeHeader(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[（）()！!：:・、,./]/g, '')
    .toLowerCase();
}

function ensureOutputColumns(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headerRow = headerRange.getValues()[0];
  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));

  const slideHeader = normalizeHeader(COL.SLIDE_URL_HEADER);
  const statusHeader = normalizeHeader(COL.STATUS_HEADER);

  let slideUrlCol = normalizedHeaders.findIndex(h => h === slideHeader) + 1;
  let statusCol = normalizedHeaders.findIndex(h => h === statusHeader) + 1;

  if (!slideUrlCol) {
    slideUrlCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, slideUrlCol).setValue(COL.SLIDE_URL_HEADER);
  }
  if (!statusCol) {
    statusCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, statusCol).setValue(COL.STATUS_HEADER);
  }

  return { slideUrlCol, statusCol };
}

function getTemplateSlideId() {
  const id = String(CONFIG.TEMPLATE_SLIDE_ID || '').trim();
  if (!id) {
    throw new Error('テンプレートID未設定です。CONFIG.TEMPLATE_SLIDE_ID を設定するか、メニューからテンプレート作成後に ID を反映してください。');
  }
  return extractSlideId(id);
}

function extractSlideId(value) {
  const text = String(value || '').trim();
  const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  return text;
}
