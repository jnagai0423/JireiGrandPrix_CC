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
成果（一言）: ${buildFeatureForSlide(data) || '（未入力）'}
${buildKpiArrowText(data) ? `成果KPI: ${buildKpiArrowText(data)}` : ''}
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

/** K列（成果の一言）本文のみを返す */
function buildFeatureForSlide(data) {
  return String(data.feature || '').trim();
}

/** L列・M列を見出し横に表示する文字列へ整形（例: （月間5件→月間8件）） */
function buildKpiArrowText(data) {
  const l = String(data.metric30Day || '').trim();
  const m = String(data.actualMetric || '').trim();
  if (!l && !m) return '';
  const arrowPart = l && m ? `${l}→${m}` : l ? `${l}→` : `→${m}`;
  return `（${arrowPart}）`;
}

function buildFallbackComment(data) {
  const feature = buildFeatureForSlide(data) || '成果を創出';
  const kpi = buildKpiArrowText(data);
  const genre = data.genre || '成果領域';
  return `「${feature}${kpi}」を実現。${genre}で再現性のある運用成果が確認できる好事例です。`;
}

function trimForSlide(text, maxLength) {
  const str = String(text || '').trim();
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return `${str.slice(0, maxLength - 1)}…`;
}
