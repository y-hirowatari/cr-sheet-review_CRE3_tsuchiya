/**
 * ネイルホリック CRシート AIレビュー — 社内Webアプリ版（スタンドアロンGAS）
 *
 * 方針（担当者の要望）：
 *   ・CRシート上には作り込まない。納品フロー（レビュー全OK→格納＆貼り付け）は変えない。
 *   ・レビュー専用の独立Webアプリとして構築する。
 *   ・CRシートは【読み取り専用】。コメントやテキストボックスを書き込まない（納品物を汚さない）。
 *
 * 仕組み：CRシートのURLを受け取り、
 *   正解＝スライド上の指示テキストの色番号（文字データなので正確）
 *   提出物＝貼り付けクリエイティブ画像を Drive OCR で読取
 * を機械照合し、結果をこのアプリ画面に表示する（AIに正誤判断はさせない）。
 *
 * デプロイ：デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *   次のユーザーとして実行: 「ウェブアプリにアクセスしているユーザー」
 *   アクセスできるユーザー: 「(社内ドメイン) 内の全員」
 * セットアップ詳細は gas-webapp/README.md を参照（Drive 詳細サービスの有効化が必要）。
 */

var CONFIG = {
  codePattern: /[A-Z]{2}\d{3}[A-Z]?/g,   // 色番号 例) BL920, WT045R, PU076D
  ocrLanguage: 'ja'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CRシート AIレビュー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- URL/ID ---------- */

function idFromUrl(url) {
  if (!url) throw new Error('URLが空です');
  var m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim(); // 生ID
  throw new Error('スライドのURL/IDを認識できませんでした');
}

function openPres(url) {
  try { return SlidesApp.openById(idFromUrl(url)); }
  catch (e) { throw new Error('スライドを開けませんでした（URL誤り、またはアクセス権がありません）'); }
}

/* ---------- テキスト・色番号ユーティリティ ---------- */

function norm(s) {
  if (!s) return '';
  s = s.normalize('NFKC').replace(/　/g, ' ');
  return s.replace(/[ \t]+/g, ' ').trim();
}

function extractCodes(text) {
  var m = norm(text).match(CONFIG.codePattern) || [];
  var seen = {}, out = [];
  m.forEach(function (c) { if (!seen[c]) { seen[c] = 1; out.push(c); } });
  return out;
}

function slideText(slide) {
  var parts = [];
  slide.getShapes().forEach(function (sh) {
    try { var t = sh.getText().asString(); if (t) parts.push(t); } catch (e) {}
  });
  slide.getTables().forEach(function (tbl) {
    for (var r = 0; r < tbl.getNumRows(); r++)
      for (var c = 0; c < tbl.getNumColumns(); c++)
        try { parts.push(tbl.getCell(r, c).getText().asString()); } catch (e) {}
  });
  return parts.join('\n');
}

function postIdsOf(slide) {
  var text = norm(slideText(slide));
  var re = /投稿no,\s*(\d+)_\s*(\d+)月\s*(\d+)日/g, m, ids = [], seen = {};
  while ((m = re.exec(text)) !== null) {
    var key = m[1] + '_' + m[2] + '/' + m[3];
    if (!seen[key]) { seen[key] = 1; ids.push(key); }
  }
  var t = text.match(/【([^】]+)】/);
  return { ids: ids, title: t ? t[1] : '' };
}

function classify(slide) {
  var text = norm(slideText(slide));
  if (/xxx|テンプレ/.test(text)) return { type: 'TEMPLATE' };
  if (/投稿日一覧/.test(text)) return { type: 'LIST' };
  if (/^(IG|X)$/.test(text)) return { type: 'LIST' };
  var pid = postIdsOf(slide);
  if (pid.ids.length >= 2) return { type: 'LIST' };
  if (pid.ids.length === 1) return { type: 'CONTENT', key: pid.ids[0], title: pid.title };
  return { type: 'CONT' };
}

function keyedSlides(pres) {
  var slides = pres.getSlides();
  var out = [], lastKey = null, lastTitle = '';
  for (var i = 0; i < slides.length; i++) {
    var cl = classify(slides[i]);
    var key = null, title = '';
    if (cl.type === 'CONTENT') { key = cl.key; title = cl.title; lastKey = key; lastTitle = title; }
    else if (cl.type === 'CONT') { key = lastKey; title = lastTitle; }
    out.push({ index: i, slide: slides[i], type: cl.type, key: key, title: title });
  }
  return out;
}

/* ---------- OCR（Drive 詳細サービス・読み取りのみ） ---------- */

function ocrBlob(blob) {
  var tmp = Drive.Files.insert(
    { title: '__ocr_tmp__', mimeType: 'application/vnd.google-apps.document' },
    blob, { ocr: true, ocrLanguage: CONFIG.ocrLanguage });
  var text = '';
  try { text = DocumentApp.openById(tmp.id).getBody().getText(); }
  finally { try { Drive.Files.remove(tmp.id); } catch (e) {} }
  return text;
}

function ocrImagesOf(slide, counter) {
  var token = ScriptApp.getOAuthToken();
  slide.getImages().forEach(function (img) {
    var blob;
    try {
      var url = img.getContentUrl();
      blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } }).getBlob();
    } catch (e) { return; }
    extractCodes(ocrBlob(blob)).forEach(function (c) { counter[c] = (counter[c] || 0) + 1; });
  });
}

/* ---------- 照合（構造データを返す。シートには書き込まない） ---------- */

function truthOfGroup(group) {
  var truth = [], seen = {};
  group.forEach(function (g) {
    extractCodes(slideText(g.slide)).forEach(function (c) { if (!seen[c]) { seen[c] = 1; truth.push(c); } });
  });
  return truth;
}

function ocrOfSlides(slides) {
  var counter = {};
  slides.forEach(function (s) { ocrImagesOf(s, counter); });
  return counter;
}

function buildResult(truth, ocrCodes, meta) {
  var truthSet = {}; truth.forEach(function (c) { truthSet[c] = 1; });
  var ocrSet = {}; ocrCodes.forEach(function (c) { ocrSet[c] = 1; });
  var matched = truth.filter(function (c) { return ocrSet[c]; });
  var missing = truth.filter(function (c) { return !ocrSet[c]; });
  var unknown = ocrCodes.filter(function (c) { return !truthSet[c]; });
  var ok = truth.length > 0 && !missing.length && !unknown.length;
  return {
    scope: meta.scope, pageIndex: meta.pageIndex || null,
    key: meta.key || '', title: meta.title || '',
    truth: truth, ocr: ocrCodes,
    matched: matched, missing: missing, unknown: unknown,
    ok: ok, hasIssue: unknown.length > 0 || missing.length > 0,
    truthEmpty: truth.length === 0
  };
}

/* ---------- Webアプリ用API（google.script.run から呼ぶ） ---------- */

/** CRシートを読み込み、投稿一覧を返す（照合はしない＝軽い） */
function api_load(url) {
  var pres = openPres(url);
  var ks = keyedSlides(pres);
  var groups = {}, order = [];
  ks.forEach(function (g) {
    if (!g.key) return;
    if (!groups[g.key]) { groups[g.key] = { key: g.key, title: g.title, pages: [], images: 0 }; order.push(g.key); }
    groups[g.key].pages.push(g.index + 1);
    groups[g.key].images += g.slide.getImages().length;
  });
  return {
    title: pres.getName(),
    slideCount: ks.length,
    posts: order.map(function (k) { return groups[k]; })
  };
}

/** 投稿キー単位で照合 */
function api_reviewPost(url, key) {
  var pres = openPres(url);
  var ks = keyedSlides(pres);
  var group = ks.filter(function (g) { return g.key === key; });
  if (!group.length) return [{ skipped: true, key: key }];
  var truth = truthOfGroup(group);
  var ocrCodes = Object.keys(ocrOfSlides(group.map(function (g) { return g.slide; })));
  return [buildResult(truth, ocrCodes, { scope: 'post', key: group[0].key, title: group[0].title })];
}

/** ページ番号指定で照合（各ページ個別。正解はそのページが属する投稿から） */
function api_reviewPages(url, spec) {
  var pres = openPres(url);
  var ks = keyedSlides(pres);
  var idxs = parsePageSpec(spec, ks.length);
  var results = [];
  idxs.forEach(function (i) {
    var item = ks[i];
    if (!item || !item.key) { results.push({ skipped: true, pageIndex: i + 1 }); return; }
    var group = ks.filter(function (g) { return g.key === item.key; });
    var truth = truthOfGroup(group);
    var ocrCodes = Object.keys(ocrOfSlides([item.slide]));
    results.push(buildResult(truth, ocrCodes, { scope: 'page', pageIndex: item.index + 1, key: item.key, title: item.title }));
  });
  return results;
}

/** 全投稿を照合（貼り付け画像のある投稿のみ） */
function api_reviewAll(url) {
  var pres = openPres(url);
  var ks = keyedSlides(pres);
  var groups = {}, order = [];
  ks.forEach(function (g) { if (g.key) { if (!groups[g.key]) { groups[g.key] = []; order.push(g.key); } groups[g.key].push(g); } });
  var results = [];
  order.forEach(function (key) {
    var group = groups[key];
    if (!group.some(function (g) { return g.slide.getImages().length; })) return;
    var truth = truthOfGroup(group);
    var ocrCodes = Object.keys(ocrOfSlides(group.map(function (g) { return g.slide; })));
    results.push(buildResult(truth, ocrCodes, { scope: 'post', key: group[0].key, title: group[0].title }));
  });
  return results;
}

function parsePageSpec(spec, total) {
  var set = {}, out = [];
  norm(spec).split(',').forEach(function (part) {
    part = part.trim(); if (!part) return;
    var m = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) {
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10);
      for (var n = Math.min(a, b); n <= Math.max(a, b); n++) addPage(n);
    } else if (/^\d+$/.test(part)) { addPage(parseInt(part, 10)); }
  });
  function addPage(n) { var i = n - 1; if (i >= 0 && i < total && !set[i]) { set[i] = 1; out.push(i); } }
  out.sort(function (a, b) { return a - b; });
  return out;
}
