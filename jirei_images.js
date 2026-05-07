function insertSitePreviewImage(pres, siteUrl) {
  const imageBlobs = fetchSitePreviewImageBlobs(siteUrl);
  if (!imageBlobs.length) return;

  try {
    const slide = pres.getSlides()[0];
    // 右側に最大3枚を重ねて配置（少しずつオフセット）
    const placements = [
      { x: 392, y: 64, w: 288, h: 211 },
      { x: 407, y: 78, w: 288, h: 211 },
      { x: 422, y: 92, w: 288, h: 211 },
    ];
    let insertedCount = 0;
    imageBlobs.slice(0, 12).forEach((blob, i) => {
      if (insertedCount >= 3) return;
      const p = placements[i] || placements[placements.length - 1];
      try {
        slide.insertImage(blob, p.x, p.y, p.w, p.h);
        insertedCount += 1;
      } catch (insertErr) {
        Logger.log('画像挿入をスキップ: ' + insertErr);
      }
    });
    // 挿入成功が足りない場合は、最初の画像を再利用して3枚に揃える
    while (insertedCount > 0 && insertedCount < 3) {
      const p = placements[insertedCount];
      slide.insertImage(imageBlobs[0], p.x, p.y, p.w, p.h);
      insertedCount += 1;
    }
    if (insertedCount < 3) {
      Logger.log(`画像挿入は${insertedCount}件でした（有効画像不足）`);
    }
  } catch (e) {
    Logger.log('サイト画像の挿入をスキップ: ' + e);
  }
}

function fetchSitePreviewImageBlobs(siteUrl) {
  const url = String(siteUrl || '').trim();
  if (!url) return [];
  if (!/^https?:\/\//i.test(url)) return [];

  try {
    const htmlRes = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
    });

    if (htmlRes.getResponseCode() < 200 || htmlRes.getResponseCode() >= 300) {
      Logger.log(`サイトHTML取得失敗: status=${htmlRes.getResponseCode()} url=${url}`);
      return [];
    }

    const html = htmlRes.getContentText();
    const candidates = [
      ...pickMetaContents(html, 'property', 'og:image'),
      ...pickMetaContents(html, 'name', 'twitter:image'),
      ...pickMetaContents(html, 'property', 'og:image:url'),
      ...pickImgSrcs(html),
      ...pickImgSrcsetCandidates(html),
    ];

    if (!candidates.length) return [];

    const uniqueUrls = [];
    candidates.forEach(v => {
      const abs = toAbsoluteUrl(url, v);
      if (abs && !uniqueUrls.includes(abs)) uniqueUrls.push(abs);
    });

    const blobs = [];
    uniqueUrls.slice(0, 60).forEach(imageUrl => {
      try {
        const imageRes = UrlFetchApp.fetch(imageUrl, {
          method: 'get',
          muteHttpExceptions: true,
          followRedirects: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
        });
        if (imageRes.getResponseCode() < 200 || imageRes.getResponseCode() >= 300) return;
        const blob = imageRes.getBlob();
        const contentType = String(blob.getContentType() || '').toLowerCase();
        if (!contentType.startsWith('image/')) return;
        // Slides に挿入しやすい形式を優先（svg や ico は除外）
        if (!/image\/(png|jpeg|jpg|gif|webp|bmp)/.test(contentType)) return;
        if (blob.getBytes().length < 4000) return; // 小さすぎるアイコン画像を除外
        const dim = getImageDimensions(blob, contentType);
        if (!dim) return;
        // 縦長画像は除外（正方形・横長のみ採用）
        if (dim.width < dim.height) return;
        blobs.push(blob);
      } catch (fetchErr) {
        Logger.log(`画像取得をスキップ: ${imageUrl} err=${fetchErr}`);
      }
    });
    return blobs;
  } catch (e) {
    Logger.log('サイト画像取得エラー: ' + e);
    return [];
  }
}

/** ページ内の <img src="..."> を抽出（data URI は除外） */
function pickImgSrcs(html) {
  const pattern = /<img[^>]*\ssrc\s*=\s*["']([^"']+)["'][^>]*>/ig;
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const v = String(m[1] || '').trim();
    if (!v) continue;
    if (/^data:/i.test(v)) continue;
    out.push(v);
  }
  return out;
}

function pickImgSrcsetCandidates(html) {
  const pattern = /<img[^>]*\ssrcset\s*=\s*["']([^"']+)["'][^>]*>/ig;
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const srcset = String(m[1] || '');
    srcset.split(',').forEach(part => {
      const first = String(part || '').trim().split(/\s+/)[0];
      if (!first) return;
      if (/^data:/i.test(first)) return;
      out.push(first);
    });
  }
  return out;
}

function getImageDimensions(blob, contentType) {
  try {
    const bytes = blob.getBytes();
    if (!bytes || bytes.length < 24) return null;

    if (/image\/png/.test(contentType)) {
      return {
        width: readUInt32BE(bytes, 16),
        height: readUInt32BE(bytes, 20),
      };
    }
    if (/image\/gif/.test(contentType)) {
      return {
        width: readUInt16LE(bytes, 6),
        height: readUInt16LE(bytes, 8),
      };
    }
    if (/image\/bmp/.test(contentType)) {
      return {
        width: Math.abs(readInt32LE(bytes, 18)),
        height: Math.abs(readInt32LE(bytes, 22)),
      };
    }
    if (/image\/jpe?g/.test(contentType)) {
      return readJpegDimensions(bytes);
    }
    if (/image\/webp/.test(contentType)) {
      return readWebpDimensions(bytes);
    }
    return null;
  } catch (e) {
    Logger.log('画像サイズ取得失敗: ' + e);
    return null;
  }
}

function readUInt16LE(bytes, i) {
  return (bytes[i] & 0xff) | ((bytes[i + 1] & 0xff) << 8);
}

function readInt32LE(bytes, i) {
  const b0 = bytes[i] & 0xff;
  const b1 = (bytes[i + 1] & 0xff) << 8;
  const b2 = (bytes[i + 2] & 0xff) << 16;
  const b3 = (bytes[i + 3] & 0xff) << 24;
  return (b0 | b1 | b2 | b3);
}

function readUInt32BE(bytes, i) {
  return ((bytes[i] & 0xff) << 24) | ((bytes[i + 1] & 0xff) << 16) | ((bytes[i + 2] & 0xff) << 8) | (bytes[i + 3] & 0xff);
}

function readJpegDimensions(bytes) {
  let i = 2;
  while (i + 9 < bytes.length) {
    if ((bytes[i] & 0xff) !== 0xff) {
      i += 1;
      continue;
    }
    const marker = bytes[i + 1] & 0xff;
    const length = ((bytes[i + 2] & 0xff) << 8) | (bytes[i + 3] & 0xff);
    if (length < 2) return null;
    if (marker >= 0xc0 && marker <= 0xc3 && i + 8 < bytes.length) {
      const height = ((bytes[i + 5] & 0xff) << 8) | (bytes[i + 6] & 0xff);
      const width = ((bytes[i + 7] & 0xff) << 8) | (bytes[i + 8] & 0xff);
      return { width, height };
    }
    i += 2 + length;
  }
  return null;
}

function readWebpDimensions(bytes) {
  if (bytes.length < 30) return null;
  const riff = String.fromCharCode(bytes[0] & 0xff, bytes[1] & 0xff, bytes[2] & 0xff, bytes[3] & 0xff);
  const webp = String.fromCharCode(bytes[8] & 0xff, bytes[9] & 0xff, bytes[10] & 0xff, bytes[11] & 0xff);
  if (riff !== 'RIFF' || webp !== 'WEBP') return null;
  const chunk = String.fromCharCode(bytes[12] & 0xff, bytes[13] & 0xff, bytes[14] & 0xff, bytes[15] & 0xff);

  if (chunk === 'VP8X' && bytes.length >= 30) {
    const w = 1 + ((bytes[24] & 0xff) | ((bytes[25] & 0xff) << 8) | ((bytes[26] & 0xff) << 16));
    const h = 1 + ((bytes[27] & 0xff) | ((bytes[28] & 0xff) << 8) | ((bytes[29] & 0xff) << 16));
    return { width: w, height: h };
  }
  if (chunk === 'VP8L' && bytes.length >= 25) {
    const b0 = bytes[21] & 0xff;
    const b1 = bytes[22] & 0xff;
    const b2 = bytes[23] & 0xff;
    const b3 = bytes[24] & 0xff;
    const w = 1 + (b0 | ((b1 & 0x3f) << 8));
    const h = 1 + ((b1 >> 6) | (b2 << 2) | ((b3 & 0x0f) << 10));
    return { width: w, height: h };
  }
  if (chunk === 'VP8 ' && bytes.length >= 30) {
    const w = readUInt16LE(bytes, 26) & 0x3fff;
    const h = readUInt16LE(bytes, 28) & 0x3fff;
    return { width: w, height: h };
  }
  return null;
}

function pickMetaContents(html, attrName, attrValue) {
  const escaped = escapeRegex(attrValue);
  const pattern = new RegExp(
    `<meta[^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*content\\s*=\\s*["']([^"']+)["'][^>]*>|` +
    `<meta[^>]*content\\s*=\\s*["']([^"']+)["'][^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*>`,
    'ig'
  );
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const v = (m[1] || m[2] || '').trim();
    if (v) out.push(v);
  }
  return out;
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
