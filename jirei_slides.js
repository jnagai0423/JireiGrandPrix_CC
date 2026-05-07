/**
 * コピー元スライドを「一から」作る（スプレッドシートのメニューから実行）。
 * 既にドライブに jireiGp_SlideTemplate などがあり CONFIG.TEMPLATE_SLIDE_ID で指している場合は不要（別ファイルが増えるだけ）。
 * SlidesApp.create の引数は、作成される Google スライドのファイル名（タイトル）になる。
 */
function createTemplate() {
  const fileTitle = '事例GP_コピー元スライド（メニューから新規作成）';
  const pres = SlidesApp.create(fileTitle);
  const slide = pres.getSlides()[0];

  // デフォルトのプレースホルダーを削除
  slide.getPlaceholders().forEach(placeholder => placeholder.remove());

  // レイアウト定数（単位: pt、16:9）
  const W = 720, H = 405;

  // タイトル補助
  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, ACCENT_BROWN);

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
  siteUrlBox.getText().getTextStyle().setFontSize(9);
  const genreBox = slide.insertTextBox('・成果ジャンル：{{GENRE}}', 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  // 成果（本文は K列、KPIは見出し右横に表示）
  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 110, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const kpiArrowBox = slide.insertTextBox('{{KPI_ARROW}}', 94, 146, 260, 20);
  kpiArrowBox.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureBox = slide.insertTextBox('{{FEATURE}}', 20, 166, W - 40, 42);
  const featureStyle = featureBox.getText().getTextStyle();
  featureStyle.setFontSize(28).setBold(true);
  safeSetTextColor(featureStyle, ACCENT_BROWN);
  featureBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

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

  const url = `https://docs.google.com/presentation/d/${pres.getId()}/edit`;
  Logger.log('テンプレート作成完了: ' + url);
  safeAlert(
    'コピー元スライドの作成が完了',
    `コピー元となるスライドを新規作成しました。\n\n▼ 確認・編集はこちら\n${url}\n\nデザインはこのスライドを直接編集してカスタマイズできます。\n\nフォーム連携のコピー元に使うには、スクリプト先頭の CONFIG.TEMPLATE_SLIDE_ID を次の ID に更新してください。\n${pres.getId()}`
  );
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
  const featureForSlide = trimForSlide(buildFeatureForSlide(data), 200);
  const kpiArrowForSlide = trimForSlide(buildKpiArrowText(data), 80);
  const replacements = {
    '{{CLIENT_NAME}}' : data.clientName,
    '{{PERSON_NAME}}' : data.personName,
    '{{INDUSTRY}}'    : data.industry,
    '{{PLAN}}'        : data.plan,
    '{{PRODUCTS}}'    : data.products,
    '{{SITE_URL}}'    : data.siteUrl,
    '{{GENRE}}'       : data.genre,
    '{{FEATURE}}'     : featureForSlide,
    '{{KPI_ARROW}}'   : kpiArrowForSlide,
    '{{DETAIL}}'      : data.detail,
    '{{AI_COMMENT}}'  : aiCommentForSlide,
    '{{FOOTER_DATE}}' : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月') + '：事例グランプリ',
  };

  // プレゼンテーション全体で一括置換（シェイプ単位より大幅に高速）
  const corePlaceholderKeys = [
    '{{CLIENT_NAME}}',
    '{{PERSON_NAME}}',
    '{{PLAN}}',
    '{{GENRE}}',
    '{{FEATURE}}',
    '{{DETAIL}}',
  ];
  let replacedCoreCount = 0;
  Object.keys(replacements).forEach(key => {
    const n = pres.replaceAllText(key, replacements[key]);
    if (corePlaceholderKeys.includes(key)) {
      replacedCoreCount += Number(n || 0);
    }
  });

  applyFeatureHighlightTextColor(pres, featureForSlide);
  alignFeatureTextBox(pres, featureForSlide);
  insertKpiArrowNearFeatureLabel(pres, kpiArrowForSlide);

  // テンプレに {{...}} が無い（デザインのみ）場合は置換だけでは何も出ないため、テキストボックスで上書き描画する
  if (replacedCoreCount === 0) {
    Logger.log(
      'コア用プレースホルダーがテンプレ内に見つかりませんでした。テキストボックスで内容を描画します。' +
        ' デザインに埋め込む場合はスライドに {{CLIENT_NAME}} などを配置してください。'
    );
    renderContentFallback(pres, data, aiComment);
  }

  insertSitePreviewImage(pres, data.siteUrl);
  ensureThreeSameSlides(pres);

  return `https://docs.google.com/presentation/d/${copy.getId()}/edit`;
}

/** 1ページ目の見た目を基準に、同一構成の3ページへ揃える */
function ensureThreeSameSlides(presentation) {
  const slides = presentation.getSlides();
  if (!slides.length) return;

  const first = slides[0];

  // いったん1ページ目だけ残す（テンプレが複数ページでも同一構成3ページに統一）
  for (let i = slides.length - 1; i >= 1; i -= 1) {
    slides[i].remove();
  }

  // 1ページ目を複製して計3ページにする
  presentation.appendSlide(first);
  presentation.appendSlide(first);
}

/** テンプレに {{CLIENT_NAME}} 等が無いとき、1枚目にフォーム内容をテキストで描画する */
function renderContentFallback(pres, data, aiComment) {
  const slide = pres.getSlides()[0];
  const W = 720;
  const H = 405;

  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, ACCENT_BROWN);

  const clientBox = slide.insertTextBox(data.clientName || '（顧客名未入力）', 20, 24, 460, 36);
  clientBox.getText().getTextStyle().setFontSize(21).setBold(true);

  const personBox = slide.insertTextBox(`担当者：${data.personName || '未入力'}`, 505, 42, 195, 20);
  personBox.getText().getTextStyle().setFontSize(11);
  personBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  const industryBox = slide.insertTextBox(`・業種：${data.industry || '未入力'}`, 20, 62, W - 40, 16);
  industryBox.getText().getTextStyle().setFontSize(9);

  const planBox = slide.insertTextBox(`・期待する成果：${data.plan || '未入力'}`, 20, 77, W - 40, 16);
  planBox.getText().getTextStyle().setFontSize(9);

  const productBox = slide.insertTextBox(`・導入製品：${data.products || '未入力'}`, 20, 92, W - 40, 16);
  productBox.getText().getTextStyle().setFontSize(9);

  const siteUrlBox = slide.insertTextBox(`・サイトURL：${data.siteUrl || '未入力'}`, 20, 107, W - 40, 16);
  siteUrlBox.getText().getTextStyle().setFontSize(9);

  const genreBox = slide.insertTextBox(`・成果ジャンル：${data.genre || '未入力'}`, 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 110, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const kpiArrowText = trimForSlide(buildKpiArrowText(data), 80);
  const kpiArrowBox = slide.insertTextBox(kpiArrowText, 94, 146, 260, 20);
  kpiArrowBox.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureText =
    trimForSlide(buildFeatureForSlide(data), 200) || '（特徴未入力）';
  const featureBox = slide.insertTextBox(featureText, 20, 166, W - 40, 42);
  const featureBodyStyle = featureBox.getText().getTextStyle();
  featureBodyStyle.setFontSize(24).setBold(true);
  safeSetTextColor(featureBodyStyle, ACCENT_BROWN);
  featureBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

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

function safeSetTextColor(textStyle, colorHex) {
  try {
    textStyle.setForegroundColor(colorHex);
  } catch (e) {
    Logger.log(`文字色の設定をスキップ: ${e}`);
  }
}

/** 1枚目で本文が K列相当（featureForSlide）と一致するシェイプをこげ茶にする（{{FEATURE}} 置換後のテンプレ用） */
function applyFeatureHighlightTextColor(presentation, displayText) {
  const needle = String(displayText || '').trim();
  if (!needle) return;
  presentation.getSlides()[0].getPageElements().forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let full = '';
    try {
      full = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (full !== needle) return;
    safeSetTextColor(shape.getText().getTextStyle(), ACCENT_BROWN);
    shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  });
}

/** 1枚目で本文が K列相当（displayText）のシェイプを左寄せ・左位置に補正する */
function alignFeatureTextBox(presentation, displayText) {
  const needle = String(displayText || '').trim();
  if (!needle) return;
  presentation.getSlides()[0].getPageElements().forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let full = '';
    try {
      full = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (full !== needle) return;
    // 見出し「成果を一言で」と同じ左マージンに合わせる
    pe.setLeft(20);
    pe.setWidth(680);
    shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  });
}

/** 1枚目で「成果」見出しの右に KPI 文字列を黒で描画（本文には連結しない） */
function insertKpiArrowNearFeatureLabel(presentation, kpiText) {
  const text = String(kpiText || '').trim();
  if (!text) return;
  const slide = presentation.getSlides()[0];

  const hasSameText = slide.getPageElements().some(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return false;
    try {
      return String(pe.asShape().getText().asString() || '').trim() === text;
    } catch (e) {
      return false;
    }
  });
  if (hasSameText) return;

  let inserted = false;
  slide.getPageElements().forEach(pe => {
    if (inserted) return;
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let label = '';
    try {
      label = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (label !== '成果を一言で' && label !== '成果') return;

    const left = pe.getLeft();
    const top = pe.getTop();
    const width = pe.getWidth();
    const h = pe.getHeight();
    const kpiBox = slide.insertTextBox(text, left + width + 6, top, 300, h + 2);
    kpiBox.getText().getTextStyle().setFontSize(12);
    kpiBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    inserted = true;
  });
}
