// 事例GP スライド自動生成 GAS（スプレッドシート紐付け）
// 設定エリア（初回はここを確認）
const CONFIG = {
  GEMINI_API_KEY   : 'AIzaSyCWQlTXU93MXYwSU3oXojiA4LjrY2Oo5PY',   // Gemini APIキー
  OUTPUT_FOLDER_ID : '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e',   // 完成スライドの保存先フォルダID
  NOTIFY_EMAIL     : 'jnagai0423@gmail.com',       // 完成通知メール
  /** コピー元: ドライブの「jireiGp_SlideTemplate」プレゼン（URL 全体でも可） */
  TEMPLATE_SLIDE_ID: '1vukTwLSPjNdbFrr89SfP7kGVlsPh50gwoAwlwdxcyh8',
  ENABLE_UI_ALERT  : false,                          // ぐるぐる回避のため、通常は false 推奨
  /** true のとき、フォーム送信ごとにヘッダーと読み取った各列の値を Logger に出す（列マッピング確認用） */
  DEBUG_SHEET_HEADERS: false,
  /**
   * 回答を読むシート名（完全一致）。空なら名前が「フォームの回答」で始まる先頭タブを使う。
   * 複数のフォーム回答シートがあるときは、ここに例: 「フォームの回答 4」を指定。
   */
  FORM_RESPONSE_SHEET_NAME: '',
};

// スプレッドシート列番号（1始まり）
const COL = {
  SLIDE_URL_HEADER : '生成スライドURL',
  STATUS_HEADER    : 'ステータス',
};

/** 「事例グランプリ」見出しと K列成果本文で揃える赤 */
const ACCENT_BROWN = '#8B5A2B';
