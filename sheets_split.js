/**
 * マスターシートの全データを読み、送信者メールアドレスのドメインごとに
 * apple.com / toyota.co.jp のようなシートへ再出力します。
 *
 * 注意:
 * - ドメイン別シートは毎回 clearContents() 後に全再出力します。
 * - そのため、ドメイン別シートに手入力した値は保持されません。
 */
function splitMasterSheetBySenderDomain(ss, masterSheet) {
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  const values = masterSheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length).getValues();
  const rowsByDomain = {};

  values.forEach(row => {
    const email = String(row[2] || '').trim();
    const domain = extractDomain(email);
    if (!domain) return;
    if (!rowsByDomain[domain]) rowsByDomain[domain] = [];
    rowsByDomain[domain].push(row);
  });

  // 出力順を安定化させるためドメイン名でソート。
  Object.keys(rowsByDomain).sort().forEach(domain => {
    const sheetName = buildSafeSheetName(domain);
    const sheet = ensureSheetWithHeaders(ss, sheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    sheet
      .getRange(2, 1, rowsByDomain[domain].length, OUTPUT_HEADERS.length)
      .setValues(rowsByDomain[domain]);
    formatOutputSheet(sheet);
  });
}

/** メールアドレス末尾からドメイン部のみを抽出する。 */
function extractDomain(email) {
  const match = String(email || '').toLowerCase().match(/@([^@\s>]+)$/);
  return match ? match[1] : '';
}

/** マスターシートの存在・ヘッダー・表示書式を保証する。 */
function ensureMasterSheet(ss) {
  const sheet = ensureSheetWithHeaders(ss, CONFIG.MASTER_SHEET_NAME);
  formatOutputSheet(sheet);
  return sheet;
}

/** 指定名のシートを取得/作成し、1行目ヘッダーを保証する。 */
function ensureSheetWithHeaders(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const headerValues = sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).getValues()[0];
  const hasHeaders = OUTPUT_HEADERS.every((header, index) => headerValues[index] === header);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
  }
  return sheet;
}

/** シート名に使えない文字を除去し、長さ上限内に収める。 */
function buildSafeSheetName(domain) {
  return String(domain || 'unknown')
    .replace(/[\[\]\*\?\/\\:]/g, '_')
    .slice(0, 100);
}

/** 出力シートの見やすさを揃えるための共通フォーマット。 */
function formatOutputSheet(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);
  sheet.setColumnWidth(5, 500);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setWrap(true);
  }
}
