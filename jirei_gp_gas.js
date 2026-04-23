// 事例GP スライド自動生成 GAS（スプレッドシート紐付け）
// 設定エリア（初回はここを確認）
const CONFIG = {
  GEMINI_API_KEY   : 'AIzaSyCWQlTXU93MXYwSU3oXojiA4LjrY2Oo5PY',   // Gemini APIキー
  OUTPUT_FOLDER_ID : '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e',   // 完成スライドの保存先フォルダID
  NOTIFY_EMAIL     : 'jnagai0423@gmail.com',       // 完成通知メール
  TEMPLATE_SHEET   : 'CONFIG',                       // テンプレートIDを保存するシート名
  TEMPLATE_SLIDE_ID: '1vukTwLSPjNdbFrr89SfP7kGVlsPh50gwoAwlwdxcyh8', // 背景デザインとして使う既存テンプレートID
  ENABLE_UI_ALERT  : false,                          // ぐるぐる回避のため、通常は false 推奨
};

// スプレッドシート列番号（1始まり）
const COL = {
  SLIDE_URL_HEADER : '生成スライドURL',
  STATUS_HEADER    : 'ステータス',
};


// 初回のみ実行: テンプレートスライドを作成
function createTemplate() {
  // テンプレートを新規作成
  const pres = SlidesApp.create('【テンプレート】事例GP_スライド');
  const slide = pres.getSlides()[0];

  // デフォルトのプレースホルダーを削除
  slide.getPlaceholders().forEach(placeholder => placeholder.remove());

  // レイアウト定数（単位: pt、16:9）
  const W = 720, H = 405;

  // タイトル補助
  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, '#8B5A2B');

  // タイトル（顧客名）
  const clientBox = slide.insertTextBox('{{CLIENT_NAME}}', 20, 24, 460, 36);
  clientBox.getText().getTextStyle().setFontSize(21).setBold(true);

  // 発表者
  const personBox = slide.insertTextBox('担当者：{{PERSON_NAME}}', 505, 42, 195, 20);
  personBox.getText().getTextStyle().setFontSize(11);
  personBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // 上段メタ情報
  const industryBox = slide.insertTextBox('・業種：{{INDUSTRY}}', 20, 62, W - 40, 16);
  industryBox.getText().getTextStyle().setFontSize(9);
  const planBox = slide.insertTextBox('・運用プラン：{{PLAN}}', 20, 77, W - 40, 16);
  planBox.getText().getTextStyle().setFontSize(9);
  const productBox = slide.insertTextBox('・導入製品：{{PRODUCTS}}', 20, 92, W - 40, 16);
  productBox.getText().getTextStyle().setFontSize(9);
  const siteUrlBox = slide.insertTextBox('・サイトURL：{{SITE_URL}}', 20, 107, W - 40, 16);
  siteUrlBox.getText().getTextStyle().setFontSize(8);
  const genreBox = slide.insertTextBox('・成果ジャンル：{{GENRE}}', 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  // 成果特徴（目立たせる）
  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 220, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureBox = slide.insertTextBox('{{FEATURE}}', 20, 166, W - 40, 42);
  featureBox.getText().getTextStyle().setFontSize(28).setBold(true);

  // 成果内容（大きく表示）
  const detailLabel = slide.insertTextBox('事例の内容', 20, 212, 220, 20);
  detailLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const detailBox = slide.insertTextBox('{{DETAIL}}', 20, 232, W - 40, 58);
  detailBox.getText().getTextStyle().setFontSize(14);

  // AIコメント（右側に寄せる）
  const aiLabel = slide.insertTextBox('※', 360, 334, 20, 20);
  aiLabel.getText().getTextStyle().setFontSize(11).setBold(true);

  // AIコメント本文
  const aiBox = slide.insertTextBox('{{AI_COMMENT}}', 380, 334, 320, 34);
  aiBox.getText().getTextStyle().setFontSize(9);

  // ⑩ フッター（日付）
  const footerBox = slide.insertTextBox('{{FOOTER_DATE}}', 500, H - 24, 200, 18);
  footerBox.getText().getTextStyle().setFontSize(9);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // 作成直後はMyドライブ直下になるため、指定フォルダへ移動
  try {
    const outputFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    const templateFile = DriveApp.getFileById(pres.getId());
    outputFolder.addFile(templateFile);
    DriveApp.getRootFolder().removeFile(templateFile);
    Logger.log('テンプレートを保存先フォルダへ移動: ' + outputFolder.getName());
  } catch (e) {
    Logger.log('テンプレート移動に失敗（Myドライブに残ります）: ' + e.message);
  }

  // CONFIGシートにテンプレートIDを保存
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG.TEMPLATE_SHEET);
    configSheet.getRange('A1').setValue('TEMPLATE_SLIDE_ID');
  }
  configSheet.getRange('B1').setValue(pres.getId());

  const url = `https://docs.google.com/presentation/d/${pres.getId()}/edit`;
  Logger.log('テンプレート作成完了: ' + url);
  safeAlert(
    'テンプレート作成完了',
    `テンプレートスライドを作成しました。\n\n▼ 確認・編集はこちら\n${url}\n\nデザインはこのスライドを直接編集してカスタマイズできます。`
  );
}


// フォーム送信トリガー用（「フォーム送信時」に設定）
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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


// Gemini APIでAIコメント生成
function generateAIComment(data) {
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    return buildFallbackComment(data);
  }

  const prompt = `
あなたはBtoBデジタルマーケティングの専門家です。
以下の事例情報をもとに、イベント発表スライド用の印象的な一言コメントを80文字以内で生成してください。
数字・成果・ポジティブな変化を強調し、聴衆の共感を引く文章にしてください。
出力はコメント本文のみ（前置き・説明不要）。

顧客名: ${data.clientName}
業種: ${data.industry}
発表者: ${data.personName}
運用プラン: ${data.plan}
導入製品: ${data.products}
サイトURL: ${data.siteUrl}
成果ジャンル: ${data.genre}
成果特徴: ${data.feature}
成果内容: ${data.detail}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { maxOutputTokens: 150, temperature: 0.75 }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const status = res.getResponseCode();
    const body = res.getContentText();
    if (status < 200 || status >= 300) {
      Logger.log(`Gemini HTTPエラー: status=${status}, body=${body}`);
      return buildFallbackComment(data);
    }

    const json = JSON.parse(body);
    const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
    if (text && String(text).trim()) {
      return String(text).trim();
    }
    Logger.log('Gemini応答異常: ' + body);
    return buildFallbackComment(data);
  } catch (err) {
    Logger.log('Gemini呼び出しエラー: ' + err);
    return buildFallbackComment(data);
  }
}


// スライド生成（テンプレートコピー + 置換）
function createSlide(data, aiComment) {
  const templateId = getTemplateSlideId();

  // テンプレートをコピー
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const fileName = `事例GP_${data.personName || '担当者未入力'}_${date}`;

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    Logger.log('フォルダ取得成功: ' + folder.getName());
  } catch (e) {
    Logger.log('フォルダ取得失敗: ' + e.message);
    throw new Error(
      '保存先フォルダにアクセスできません。OUTPUT_FOLDER_ID と権限を確認してください。' +
      ' folderId=' + CONFIG.OUTPUT_FOLDER_ID
    );
  }

  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  Logger.log('生成ファイルID: ' + copy.getId());
  Logger.log('生成URL: https://docs.google.com/presentation/d/' + copy.getId() + '/edit');
  const pres = SlidesApp.openById(copy.getId());

  // プレースホルダー置換マップ
  const aiCommentForSlide = trimForSlide(aiComment, 95);
  const replacements = {
    '{{CLIENT_NAME}}' : data.clientName,
    '{{PERSON_NAME}}' : data.personName,
    '{{INDUSTRY}}'    : data.industry,
    '{{PLAN}}'        : data.plan,
    '{{PRODUCTS}}'    : data.products,
    '{{SITE_URL}}'    : data.siteUrl,
    '{{GENRE}}'       : data.genre,
    '{{FEATURE}}'     : data.feature,
    '{{DETAIL}}'      : data.detail,
    '{{AI_COMMENT}}'  : aiCommentForSlide,
    '{{DATE}}'        : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日'),
    '{{FOOTER_DATE}}' : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月') + '：事例グランプリ',
    // 旧テンプレート互換（古いプレースホルダーが残っていても置換する）
    '{{SUMMARY}}'     : data.feature,
    '{{INDUSTORY}}'   : data.industry || data.genre,
    '{{URL}}'         : data.siteUrl,
    '{{SCORE}}'       : '',
  };

  // プレゼンテーション全体で一括置換（シェイプ単位より大幅に高速）
  let replacedCoreCount = 0;
  Object.keys(replacements).forEach(key => {
    const replaced = pres.replaceAllText(key, replacements[key]);
    if (['{{CLIENT_NAME}}', '{{PERSON_NAME}}', '{{PLAN}}', '{{GENRE}}', '{{FEATURE}}', '{{DETAIL}}'].includes(key)) {
      replacedCoreCount += Number(replaced || 0);
    }
  });

  // 背景テンプレートにプレースホルダーが無い場合でも、必要情報を必ず描画する
  if (replacedCoreCount === 0) {
    renderContentFallback(pres, data, aiComment);
  }
  cleanupLegacyScoreElements(pres);
  insertSitePreviewImage(pres, data.siteUrl);

  return `https://docs.google.com/presentation/d/${copy.getId()}/edit`;
}

function getTemplateSlideId() {
  // 1) CONFIG定数で明示指定されていれば最優先
  const configTemplateId = String(CONFIG.TEMPLATE_SLIDE_ID || '').trim();
  if (configTemplateId) return extractSlideId(configTemplateId);

  // 2) 既存運用との互換でCONFIGシートのB1も参照
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!configSheet) {
    throw new Error('テンプレートIDが見つかりません。CONFIG.TEMPLATE_SLIDE_ID を設定してください。');
  }
  const sheetTemplateId = String(configSheet.getRange('B1').getValue() || '').trim();
  if (!sheetTemplateId) {
    throw new Error('テンプレートID未設定です。CONFIG.TEMPLATE_SLIDE_ID または CONFIGシートB1を設定してください。');
  }
  return extractSlideId(sheetTemplateId);
}

function extractSlideId(value) {
  const text = String(value || '').trim();
  const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  return text;
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
    .addItem('① テンプレートを作成（初回のみ）', 'createTemplate')
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
    onFormSubmit(null);
    const outCol = ensureOutputColumns(sheet);
    const url = sheet.getRange(lastRow, outCol.slideUrlCol).getValue();
    safeAlert('完了', `スライドが生成されました。\n\n${url}`);
  }
}


// ヘッダー名ベースで行データを取得（フォーム項目変更に強くする）
function getRowDataByHeader(sheet, rowNumber) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valueRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));
  const getByKey = (keys, excludeKeys) => {
    const normalizedKeys = (keys || []).map(k => normalizeHeader(k));
    const normalizedExcludes = (excludeKeys || []).map(k => normalizeHeader(k));
    const idx = normalizedHeaders.findIndex(h => {
      const hit = normalizedKeys.some(k => k && h.includes(k));
      if (!hit) return false;
      return !normalizedExcludes.some(ex => ex && h.includes(ex));
    });
    if (idx === -1) return '';
    return String(valueRow[idx] || '').trim();
  };

  const personName = getByKey(['自分の名前を入力', '自分の名前', '担当者名', '氏名']);
  const clientName = getByKey(['顧客企業名を正式名称で入力', '顧客企業名', '顧客名']);
  const industry = getByKey(['顧客企業の業種', '業種']);
  const products = getByKey(['導入済のcloudcircus製品', '導入済', 'cloudcircus製品'], ['顧客企業名', '顧客名']);
  const plan = getByKey(['運用中のコンサルティングプラン', '運用中', 'プラン']);
  const siteUrl = getByKey(['サイトurl', 'サイトurlを入力', 'webサイトurl', 'ホームページurl', 'url'], ['生成スライドurl']);
  const genre = getByKey(['成果事例のジャンル', 'ジャンル']);
  const feature = getByKey(['成果事例の特徴', '15文字']);
  const detail = getByKey(['成果事例の内容', '300文字', '詳細']);

  return {
    personName,
    clientName,
    industry,
    products,
    siteUrl,
    genre,
    feature,
    detail,
    plan
  };
}

function cleanupLegacyScoreElements(pres) {
  const targets = ['スコア', '点'];
  pres.getSlides().forEach(slide => {
    slide.getPageElements().forEach(el => {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const shape = el.asShape();
      const text = String(shape.getText().asString() || '').trim();
      if (targets.includes(text)) {
        el.remove();
      }
    });
  });
}

function renderContentFallback(pres, data, aiComment) {
  const slide = pres.getSlides()[0];
  const W = 720, H = 405;

  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, '#8B5A2B');

  const clientBox = slide.insertTextBox(data.clientName || '（顧客名未入力）', 20, 24, 460, 36);
  clientBox.getText().getTextStyle().setFontSize(21).setBold(true);

  const personBox = slide.insertTextBox(`担当者：${data.personName || '未入力'}`, 505, 42, 195, 20);
  personBox.getText().getTextStyle().setFontSize(11);
  personBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  const industryBox = slide.insertTextBox(`・業種：${data.industry || '未入力'}`, 20, 62, W - 40, 16);
  industryBox.getText().getTextStyle().setFontSize(9);

  const planBox = slide.insertTextBox(`・運用プラン：${data.plan || '未入力'}`, 20, 77, W - 40, 16);
  planBox.getText().getTextStyle().setFontSize(9);

  const productBox = slide.insertTextBox(`・導入製品：${data.products || '未入力'}`, 20, 92, W - 40, 16);
  productBox.getText().getTextStyle().setFontSize(9);

  const siteUrlBox = slide.insertTextBox(`・サイトURL：${data.siteUrl || '未入力'}`, 20, 107, W - 40, 16);
  siteUrlBox.getText().getTextStyle().setFontSize(8);

  const genreBox = slide.insertTextBox(`・成果ジャンル：${data.genre || '未入力'}`, 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 220, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureBox = slide.insertTextBox(data.feature || '（特徴未入力）', 20, 166, W - 40, 42);
  featureBox.getText().getTextStyle().setFontSize(28).setBold(true);

  const detailLabel = slide.insertTextBox('事例の内容', 20, 212, 220, 20);
  detailLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const detailBox = slide.insertTextBox(data.detail || '（詳細未入力）', 20, 232, W - 40, 58);
  detailBox.getText().getTextStyle().setFontSize(14);

  const aiLabel = slide.insertTextBox('※', 360, 334, 20, 20);
  aiLabel.getText().getTextStyle().setFontSize(11).setBold(true);
  const aiBox = slide.insertTextBox(trimForSlide(aiComment || '（AIコメントなし）', 95), 380, 334, 320, 34);
  aiBox.getText().getTextStyle().setFontSize(9);

  const footerText = `${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月')}：事例グランプリ`;
  const footerBox = slide.insertTextBox(footerText, 500, H - 24, 200, 18);
  footerBox.getText().getTextStyle().setFontSize(9);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}

function buildFallbackComment(data) {
  const feature = data.feature || '成果を創出';
  const genre = data.genre || '成果領域';
  return `「${feature}」を実現。${genre}で再現性のある運用成果が確認できる好事例です。`;
}

function trimForSlide(text, maxLength) {
  const str = String(text || '').trim();
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return `${str.slice(0, maxLength - 1)}…`;
}

function safeSetTextColor(textStyle, colorHex) {
  try {
    textStyle.setForegroundColor(colorHex);
  } catch (e) {
    Logger.log(`文字色の設定をスキップ: ${e}`);
  }
}

function insertSitePreviewImage(pres, siteUrl) {
  const imageBlob = fetchSitePreviewImageBlob(siteUrl);
  if (!imageBlob) return;

  try {
    const slide = pres.getSlides()[0];
    // 右側の空き領域にサイト画像を配置
    slide.insertImage(imageBlob, 500, 96, 190, 140);
  } catch (e) {
    Logger.log('サイト画像の挿入をスキップ: ' + e);
  }
}

function fetchSitePreviewImageBlob(siteUrl) {
  const url = String(siteUrl || '').trim();
  if (!url) return null;
  if (!/^https?:\/\//i.test(url)) return null;

  try {
    const htmlRes = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
    });

    if (htmlRes.getResponseCode() < 200 || htmlRes.getResponseCode() >= 300) {
      Logger.log(`サイトHTML取得失敗: status=${htmlRes.getResponseCode()} url=${url}`);
      return null;
    }

    const html = htmlRes.getContentText();
    const imageUrl =
      pickMetaContent(html, 'property', 'og:image') ||
      pickMetaContent(html, 'name', 'twitter:image') ||
      pickMetaContent(html, 'property', 'og:image:url');

    if (!imageUrl) return null;

    const absoluteImageUrl = toAbsoluteUrl(url, imageUrl);
    const imageRes = UrlFetchApp.fetch(absoluteImageUrl, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
    });
    if (imageRes.getResponseCode() < 200 || imageRes.getResponseCode() >= 300) {
      Logger.log(`OG画像取得失敗: status=${imageRes.getResponseCode()} url=${absoluteImageUrl}`);
      return null;
    }

    const blob = imageRes.getBlob();
    if (!String(blob.getContentType() || '').startsWith('image/')) {
      return null;
    }
    return blob;
  } catch (e) {
    Logger.log('サイト画像取得エラー: ' + e);
    return null;
  }
}

function pickMetaContent(html, attrName, attrValue) {
  const escaped = escapeRegex(attrValue);
  const pattern = new RegExp(
    `<meta[^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*content\\s*=\\s*["']([^"']+)["'][^>]*>|` +
    `<meta[^>]*content\\s*=\\s*["']([^"']+)["'][^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*>`,
    'i'
  );
  const m = String(html || '').match(pattern);
  return (m && (m[1] || m[2])) ? (m[1] || m[2]).trim() : '';
}

function toAbsoluteUrl(baseUrl, maybeRelative) {
  const raw = String(maybeRelative || '').trim();
  if (!raw) return '';
  if (/^https?:\/\//i.test(raw)) return raw;
  if (/^\/\//.test(raw)) {
    const scheme = /^https:\/\//i.test(baseUrl) ? 'https:' : 'http:';
    return `${scheme}${raw}`;
  }

  const hostMatch = String(baseUrl).match(/^(https?:\/\/[^\/?#]+)/i);
  if (!hostMatch) return raw;
  const origin = hostMatch[1];

  if (raw.startsWith('/')) return `${origin}${raw}`;

  const pathBase = String(baseUrl).replace(/[#?].*$/, '').replace(/\/[^/]*$/, '/');
  return `${pathBase}${raw}`;
}

function escapeRegex(str) {
  return String(str || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
