/**
 * CRシート AIレビュー — 社内Webアプリ版（スタンドアロンGAS）
 *
 * フロー：
 *   1. アプリにCRシートのURLを入れて読み込む → 投稿一覧＋各投稿の「正解の色番号」を取得
 *   2. レビューしたい投稿を選ぶ（指示＝正解は常にCRシート内にある）
 *   3. 貼り付ける"前"の完成クリエイティブ画像をアプリにアップロード
 *   4. アップ画像をOCR → 正解と機械照合 → 結果を画面表示
 *
 * 方針：CRシートは【読み取り専用】（書き込まない・納品フローを変えない）。
 *       AIに正誤判断はさせない。OCRは「読む」だけ、一致判定は機械が文字単位で行う。
 *
 * セットアップ：
 *   1. GASエディタ →「サービス」→「+」→「Google Slides API」を追加
 *   2. GASエディタ →「サービス」→「+」→「Drive API」を追加（OCR用、既存）
 *   3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      次のユーザーとして実行: 「ウェブアプリにアクセスしているユーザー」
 *      アクセスできるユーザー: 「(社内ドメイン) 内の全員」
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
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim();
  throw new Error('スライドのURL/IDを認識できませんでした');
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

/* ---------- Slides API JSON からテキスト抽出 ---------- */

/** textElements 配列からプレーンテキストを結合する */
function textContent(textElements) {
  if (!textElements) return '';
  var out = '';
  for (var i = 0; i < textElements.length; i++) {
    if (textElements[i].textRun) out += textElements[i].textRun.content;
  }
  return out;
}

/** JSON形式のスライドオブジェクトから全テキストを抽出する */
function slideTextJson(slide) {
  var parts = [];
  var elems = slide.pageElements || [];
  for (var i = 0; i < elems.length; i++) {
    var el = elems[i];
    // Shape（テキストボックス等）
    if (el.shape && el.shape.text) {
      parts.push(textContent(el.shape.text.textElements));
    }
    // Table
    if (el.table && el.table.tableRows) {
      var rows = el.table.tableRows;
      for (var r = 0; r < rows.length; r++) {
        var cells = rows[r].tableCells || [];
        for (var c = 0; c < cells.length; c++) {
          if (cells[c].text) {
            parts.push(textContent(cells[c].text.textElements));
          }
        }
      }
    }
    // Group（グループ化された要素を再帰的に処理）
    if (el.elementGroup && el.elementGroup.children) {
      var children = el.elementGroup.children;
      for (var j = 0; j < children.length; j++) {
        if (children[j].shape && children[j].shape.text) {
          parts.push(textContent(children[j].shape.text.textElements));
        }
      }
    }
  }
  return parts.join('\n');
}

/* ---------- スライド分類（テキストベース） ---------- */

function postIdsOfText(text) {
  var n = norm(text);
  var re = /投稿no,\s*(\d+)_\s*(\d+)月\s*(\d+)日/g, m, ids = [], seen = {};
  while ((m = re.exec(n)) !== null) {
    var key = m[1] + '_' + m[2] + '/' + m[3];
    if (!seen[key]) { seen[key] = 1; ids.push(key); }
  }
  var t = n.match(/【([^】]+)】/);
  return { ids: ids, title: t ? t[1] : '' };
}

function classifyText(text) {
  var n = norm(text);
  if (/xxx|テンプレ/.test(n)) return { type: 'TEMPLATE' };
  if (/投稿日一覧/.test(n)) return { type: 'LIST' };
  if (/^(IG|X)$/.test(n)) return { type: 'LIST' };
  var pid = postIdsOfText(text);
  if (pid.ids.length >= 2) return { type: 'LIST' };
  if (pid.ids.length === 1) return { type: 'CONTENT', key: pid.ids[0], title: pid.title };
  return { type: 'CONT' };
}

/* ---------- OCR（アップロード画像。Drive 詳細サービス使用） ---------- */

function ocrBlob(blob) {
  var tmp = Drive.Files.insert(
    { title: '__ocr_tmp__', mimeType: 'application/vnd.google-apps.document' },
    blob, { ocr: true, ocrLanguage: CONFIG.ocrLanguage });
  var text = '';
  try { text = DocumentApp.openById(tmp.id).getBody().getText(); }
  finally { try { Drive.Files.remove(tmp.id); } catch (e) {} }
  return text;
}

/* ---------- 照合結果（構造データ） ---------- */

function buildResult(truth, ocrCodes, meta) {
  var truthSet = {}; truth.forEach(function (c) { truthSet[c] = 1; });
  var ocrSet = {}; ocrCodes.forEach(function (c) { ocrSet[c] = 1; });
  var matched = truth.filter(function (c) { return ocrSet[c]; });
  var missing = truth.filter(function (c) { return !ocrSet[c]; });
  var unknown = ocrCodes.filter(function (c) { return !truthSet[c]; });
  var ok = truth.length > 0 && !missing.length && !unknown.length;
  return {
    scope: meta.scope || 'upload',
    key: meta.key || '', title: meta.title || '', images: meta.images || 0,
    truth: truth, ocr: ocrCodes,
    matched: matched, missing: missing, unknown: unknown,
    ok: ok, hasIssue: unknown.length > 0 || missing.length > 0,
    truthEmpty: truth.length === 0
  };
}

/* ---------- Webアプリ用API ---------- */

/**
 * STEP1：CRシートを読み込み、投稿一覧＋各投稿の正解色番号を返す。
 * Slides Advanced Service (Slides.Presentations.get) で全スライドのデータを
 * 1回のAPI呼び出しで一括取得し、ローカルでJSON解析する。
 * 旧方式（SlidesApp）では72ページで6分超→タイムアウトしていたが、
 * この方式では数秒で完了する。
 */
function api_load(url) {
  var id = idFromUrl(url);
  var pres;
  try { pres = Slides.Presentations.get(id); }
  catch (e) { throw new Error('スライドを開けませんでした（URL誤り、またはアクセス権がありません）'); }
  var slides = pres.slides || [];

  // 全スライドのテキストをローカルで一括抽出（API呼び出しなし）
  var slideTexts = [];
  for (var i = 0; i < slides.length; i++) {
    slideTexts.push(slideTextJson(slides[i]));
  }

  // スライドを分類してキー付け
  var lastKey = null, lastTitle = '';
  var groups = {}, order = [];
  for (var i = 0; i < slideTexts.length; i++) {
    var cl = classifyText(slideTexts[i]);
    var key = null, title = '';
    if (cl.type === 'CONTENT') { key = cl.key; title = cl.title; lastKey = key; lastTitle = title; }
    else if (cl.type === 'CONT') { key = lastKey; title = lastTitle; }
    if (!key) continue;
    if (!groups[key]) { groups[key] = { key: key, title: title, pages: [], texts: [] }; order.push(key); }
    groups[key].pages.push(i + 1);
    groups[key].texts.push(slideTexts[i]);
  }

  // 各投稿グループから正解の色番号を抽出
  var posts = order.map(function (k) {
    var grp = groups[k];
    var truth = [], seen = {};
    grp.texts.forEach(function (t) {
      extractCodes(t).forEach(function (c) { if (!seen[c]) { seen[c] = 1; truth.push(c); } });
    });
    return { key: grp.key, title: grp.title, pages: grp.pages, truth: truth };
  });

  return { title: pres.title, slideCount: slides.length, posts: posts };
}

/**
 * STEP3-4：アップロードされた完成画像をOCRし、選んだ投稿の正解と照合する。
 * data = { key, title, truth:[...], files:[{name, mime, b64}] }
 * 正解(truth)はSTEP1で取得済みのものを渡す（再読込不要で高速）。
 */
function api_reviewUpload(data) {
  data = data || {};
  var truth = data.truth || [];
  var files = data.files || [];
  if (!files.length) throw new Error('画像がアップロードされていません');

  var counter = {};
  files.forEach(function (f) {
    var blob;
    try { blob = Utilities.newBlob(Utilities.base64Decode(f.b64), f.mime || 'image/png', f.name || 'image'); }
    catch (e) { return; }
    extractCodes(ocrBlob(blob)).forEach(function (c) { counter[c] = (counter[c] || 0) + 1; });
  });
  var ocrCodes = Object.keys(counter);
  return buildResult(truth, ocrCodes, { scope: 'upload', key: data.key, title: data.title, images: files.length });
}
