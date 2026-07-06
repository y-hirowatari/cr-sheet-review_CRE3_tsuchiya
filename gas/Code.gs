/**
 * ネイルホリック CRシート AIレビュー — Google Slides ワンクリック照合（プロトタイプ）
 *
 * 前提（実物の7月CRシートで確認した構造）：
 *   ・1つの投稿は「IG_fe＿投稿no,N_◯月◯日(◯) 17:00 【タイトル】」で識別される
 *   ・1投稿につき複数スライド（指示ページ／貼り付けページ／FIX）が並ぶ
 *   ・貼り付けページには提出クリエイティブ画像が貼られている（＝OCR対象）
 *   ・投稿日一覧ページ（複数の投稿IDが並ぶ）やテンプレページは照合対象外
 *
 * 本スクリプトは、同じ投稿IDのスライド群をまとめて
 *   正解＝スライド上のテキストの色番号（文字データなので正確）
 *   提出物＝貼り付け画像を Drive OCR で読取
 * を機械照合し、結果をスライド上のコメント（テキストボックス）とノートに書き出す。
 *
 * 設計方針：AIに正誤を判断させない。OCRは「読む」だけ、一致判定は機械が文字単位で行う。
 *
 * ★まず「AIレビュー → 構造を診断」を実行し、出力されたGoogleドキュメントのURLを共有してください。
 *   実際のスライド構造に合わせて、正解テキストと貼り付け画像の対応づけを最終調整します。
 *
 * セットアップは gas/README.md を参照（Drive 詳細サービスの有効化が必要）。
 */

var CONFIG = {
  codePattern: /[A-Z]{2}\d{3}[A-Z]?/g,   // 色番号 例) BL920, WT045R, PU076D
  ocrLanguage: 'ja',
  writeComment: true,
  writeSpeakerNotes: true,
  commentTitle: '🔎 AI一次チェック（照合結果）'
};

function onOpen() {
  SlidesApp.getUi()
    .createMenu('AIレビュー')
    .addItem('▶ レビューパネルを開く', 'showSidebar')
    .addToUi();
}

/** サイドバー（レビューパネル）を開く */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('AIレビュー');
  SlidesApp.getUi().showSidebar(html);
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

/** スライド上の投稿IDを抽出。返り値 {ids:[key...], title} */
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

/** スライドの種別を判定：TEMPLATE / LIST / CONTENT / CONT(継続) */
function classify(slide) {
  var text = norm(slideText(slide));
  if (/xxx|テンプレ/.test(text)) return { type: 'TEMPLATE' };
  if (/投稿日一覧/.test(text)) return { type: 'LIST' };       // 「IG/X 投稿日一覧」ページ
  if (/^(IG|X)$/.test(text)) return { type: 'LIST' };         // 媒体の区切りスライド
  var pid = postIdsOf(slide);
  if (pid.ids.length >= 2) return { type: 'LIST' };          // 投稿日一覧（複数ID）
  if (pid.ids.length === 1) return { type: 'CONTENT', key: pid.ids[0], title: pid.title };
  return { type: 'CONT' };                                    // ID無し（画像のみ等）→ 直前に従属
}

/** 全スライドに投稿キーを割り当てる（CONTは直前のキーを継承） */
function keyedSlides() {
  var slides = SlidesApp.getActivePresentation().getSlides();
  var out = [], lastKey = null, lastTitle = '';
  for (var i = 0; i < slides.length; i++) {
    var cl = classify(slides[i]);
    var key = null, title = '';
    if (cl.type === 'CONTENT') { key = cl.key; title = cl.title; lastKey = key; lastTitle = title; }
    else if (cl.type === 'CONT') { key = lastKey; title = lastTitle; }
    // TEMPLATE / LIST は key=null（対象外）
    out.push({ index: i, slide: slides[i], type: cl.type, key: key, title: title });
  }
  return out;
}

/* ---------- OCR（Drive 詳細サービス） ---------- */

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

/* ---------- 照合 ---------- */

/** 投稿グループ全体の正解色番号（テキストから）を返す */
function truthOfGroup(group) {
  var truth = [], seen = {};
  group.forEach(function (g) {
    extractCodes(slideText(g.slide)).forEach(function (c) { if (!seen[c]) { seen[c] = 1; truth.push(c); } });
  });
  return truth;
}

/** 指定スライド配列の画像をOCRし、色番号→検出回数を返す */
function ocrOfSlides(slides) {
  var counter = {};
  slides.forEach(function (s) { ocrImagesOf(s, counter); });
  return counter;
}

/** 正解とOCR結果から照合レポート（構造データ）を組み立てる。サイドバーはこの構造をそのまま描画する */
function buildResult(truth, ocrCodes, meta) {
  var truthSet = {}; truth.forEach(function (c) { truthSet[c] = 1; });
  var ocrSet = {}; ocrCodes.forEach(function (c) { ocrSet[c] = 1; });

  var matched = truth.filter(function (c) { return ocrSet[c]; });
  var missing = truth.filter(function (c) { return !ocrSet[c]; });
  var unknown = ocrCodes.filter(function (c) { return !truthSet[c]; });
  var ok = truth.length > 0 && !missing.length && !unknown.length;

  return {
    scope: meta.scope,                 // 'page' or 'post'
    pageIndex: meta.pageIndex || null, // 1始まり
    key: meta.key || '',
    title: meta.title || '',
    truth: truth, ocr: ocrCodes,
    matched: matched, missing: missing, unknown: unknown,
    ok: ok, hasIssue: unknown.length > 0 || missing.length > 0,
    truthEmpty: truth.length === 0
  };
}

/** 構造データ → スライド上コメント/ノート用のテキスト */
function resultToText(r) {
  var head = (r.scope === 'page' ? ('ページ p' + r.pageIndex) : '投稿') +
    ' / ' + (r.key || '?') + ' 【' + (r.title || '') + '】';
  var lines = [CONFIG.commentTitle, head];
  if (r.truthEmpty) lines.push('※ 正解の色番号を取得できませんでした（指示テキストが見当たらない）');
  lines.push('正解(指示): ' + (r.truth.join(', ') || '（なし）'));
  lines.push('OCR(提出物): ' + (r.ocr.join(', ') || '（なし）'));
  lines.push('一致: ' + r.matched.length + '/' + r.truth.length);
  if (r.unknown.length) lines.push('❌ 取り違え疑い（正解に無い）: ' + r.unknown.join(', '));
  if (r.missing.length) lines.push('⚠ 未検出（正解にあるが読めず/未反映）: ' + r.missing.join(', '));
  if (r.ok) lines.push('✅ 色番号は正解と一致');
  return lines.join('\n');
}

/** 投稿キー単位で照合（投稿内の全ページを対象にOCR） */
function reviewGroup(group) {
  var truth = truthOfGroup(group);
  var ocrCodes = Object.keys(ocrOfSlides(group.map(function (g) { return g.slide; })));
  return buildResult(truth, ocrCodes,
    { scope: 'post', key: group[0].key, title: group[0].title });
}

/** 1ページだけ照合（正解はその投稿から引き、OCRはこのページの画像のみ） */
function reviewSinglePage(item, group) {
  var truth = truthOfGroup(group);
  var ocrCodes = Object.keys(ocrOfSlides([item.slide]));
  return buildResult(truth, ocrCodes,
    { scope: 'page', pageIndex: item.index + 1, key: item.key, title: item.title });
}

function writeFindings(slide, result) {
  var text = resultToText(result);
  if (CONFIG.writeSpeakerNotes) {
    try { slide.getNotesPage().getSpeakerNotesShape().getText().setText(text); } catch (e) {}
  }
  if (CONFIG.writeComment) {
    clearFindingsOnSlide(slide);
    var box = slide.insertTextBox(text, 12, 12, 340, 130);
    box.setTitle('__ai_review__');
    box.getText().getTextStyle().setFontSize(9)
      .setForegroundColor(result.hasIssue ? '#C0392B' : '#1E8449');
  }
}

function clearFindingsOnSlide(slide) {
  slide.getShapes().forEach(function (sh) {
    try { if (sh.getTitle() === '__ai_review__') sh.remove(); } catch (e) {}
  });
}

/* ---------- メニュー実行 ---------- */

function activeSlideIndex() {
  var pres = SlidesApp.getActivePresentation();
  var sel = pres.getSelection();
  var page = sel && sel.getCurrentPage();
  var slides = pres.getSlides();
  if (page) for (var i = 0; i < slides.length; i++)
    if (slides[i].getObjectId() === page.getObjectId()) return i;
  return 0;
}

/** いま選択されているスライドの index 配列を返す（サムネイルで複数選択に対応） */
function selectedSlideIndexes() {
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  var idById = {};
  slides.forEach(function (s, i) { idById[s.getObjectId()] = i; });

  var sel = pres.getSelection();
  var ids = [];
  if (sel) {
    var pr = sel.getPageRange();                 // サムネイルで選択中のページ群
    if (pr) pr.getPages().forEach(function (p) { ids.push(p.getObjectId()); });
    if (!ids.length && sel.getCurrentPage()) ids.push(sel.getCurrentPage().getObjectId());
  }
  var idxs = [];
  ids.forEach(function (id) { if (idById[id] !== undefined) idxs.push(idById[id]); });
  idxs.sort(function (a, b) { return a - b; });
  return idxs;
}

/** index配列を照合し、結果オブジェクトの配列を返す（コメントも書き出す）。UIには依存しない */
function reviewIndexesCore(idxs) {
  var ks = keyedSlides();
  var results = [];
  idxs.forEach(function (i) {
    var item = ks[i];
    if (!item || !item.key) { results.push({ skipped: true, pageIndex: i + 1 }); return; }
    var group = ks.filter(function (g) { return g.key === item.key; });
    var result = reviewSinglePage(item, group);
    writeFindings(item.slide, result);
    results.push(result);
  });
  return results;
}

/** メニュー：サムネイルで選択したページを照合（サイドバー未使用時のフォールバック） */
function reviewSelectedPages() {
  var results = reviewIndexesCore(selectedSlideIndexes());
  SlidesApp.getUi().alert(summarize(results));
}

/** メニュー：ページ番号を指定して照合 */
function reviewByNumbers() {
  var ui = SlidesApp.getUi();
  var res = ui.prompt('レビューするページ番号',
    '例: 7 ／ 5,7 ／ 9-11（カンマと範囲を使えます）', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var total = SlidesApp.getActivePresentation().getSlides().length;
  var results = reviewIndexesCore(parsePageSpec(res.getResponseText(), total));
  ui.alert(summarize(results));
}

function summarize(results) {
  var done = 0, issues = 0, skipped = 0, last = '';
  results.forEach(function (r) {
    if (r.skipped) { skipped++; return; }
    done++; if (r.hasIssue) issues++; last = resultToText(r);
  });
  var msg = '照合完了：' + done + 'ページ（要確認 ' + issues + '）';
  if (skipped) msg += ' / 対象外 ' + skipped + 'ページ';
  if (done === 1) msg += '\n\n' + last;
  return msg;
}

/* ---------- サイドバー用API（google.script.run から呼ぶ。構造データを返す） ---------- */

function api_reviewSelected() { return reviewIndexesCore(selectedSlideIndexes()); }

function api_reviewByNumbers(spec) {
  var total = SlidesApp.getActivePresentation().getSlides().length;
  return reviewIndexesCore(parsePageSpec(spec, total));
}

function api_reviewActivePost() {
  var ks = keyedSlides();
  var idx = activeSlideIndex();
  var key = ks[idx] && ks[idx].key;
  if (!key) return [{ skipped: true, pageIndex: idx + 1 }];
  var group = ks.filter(function (g) { return g.key === key; });
  var result = reviewGroup(group);
  group.forEach(function (g) { if (g.slide.getImages().length) writeFindings(g.slide, result); });
  return [result];
}

function api_reviewAll() {
  var ks = keyedSlides();
  var groups = {}, order = [];
  ks.forEach(function (g) { if (g.key) { if (!groups[g.key]) { groups[g.key] = []; order.push(g.key); } groups[g.key].push(g); } });
  var results = [];
  order.forEach(function (key) {
    var group = groups[key];
    if (!group.some(function (g) { return g.slide.getImages().length; })) return;
    var result = reviewGroup(group);
    group.forEach(function (g) { if (g.slide.getImages().length) writeFindings(g.slide, result); });
    results.push(result);
  });
  return results;
}

/** いま選択されているページの情報（サイドバー表示用） */
function api_selectionInfo() {
  var ks = keyedSlides();
  var idxs = selectedSlideIndexes();
  return idxs.map(function (i) {
    var it = ks[i];
    return { page: i + 1, type: it.type, key: it.key || '', title: it.title || '', images: it.slide.getImages().length };
  });
}

function api_clear() {
  var n = 0;
  SlidesApp.getActivePresentation().getSlides().forEach(function (s) {
    s.getShapes().forEach(function (sh) { try { if (sh.getTitle() === '__ai_review__') { sh.remove(); n++; } } catch (e) {} });
  });
  return n;
}

function api_diagnoseUrl() { return diagnose(true); }

/** "5,7,9-11" のような指定を 0始まり index 配列に変換 */
function parsePageSpec(spec, total) {
  var set = {}, out = [];
  norm(spec).split(',').forEach(function (part) {
    part = part.trim(); if (!part) return;
    var m = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) {
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10);
      for (var n = Math.min(a, b); n <= Math.max(a, b); n++) addPage(n);
    } else if (/^\d+$/.test(part)) {
      addPage(parseInt(part, 10));
    }
  });
  function addPage(n) { var i = n - 1; if (i >= 0 && i < total && !set[i]) { set[i] = 1; out.push(i); } }
  out.sort(function (a, b) { return a - b; });
  return out;
}

function reviewActivePost() {
  var ks = keyedSlides();
  var idx = activeSlideIndex();
  var key = ks[idx].key;
  if (!key) { SlidesApp.getUi().alert('このスライドは照合対象外です（一覧/テンプレ、または投稿IDなし）。'); return; }
  var group = ks.filter(function (g) { return g.key === key; });
  var result = reviewGroup(group);
  group.forEach(function (g) { if (g.slide.getImages().length) writeFindings(g.slide, result); });
  writeFindings(ks[idx].slide, result); // アクティブスライドには必ず出す
  SlidesApp.getUi().alert(resultToText(result));
}

function reviewAllPosts() {
  var ks = keyedSlides();
  var groups = {};
  ks.forEach(function (g) { if (g.key) (groups[g.key] = groups[g.key] || []).push(g); });
  var reviewed = 0, issues = 0;
  Object.keys(groups).forEach(function (key) {
    var group = groups[key];
    var hasImg = group.some(function (g) { return g.slide.getImages().length; });
    if (!hasImg) return;                     // 貼り付け画像が無い投稿はスキップ
    var result = reviewGroup(group);
    group.forEach(function (g) { if (g.slide.getImages().length) writeFindings(g.slide, result); });
    reviewed++; if (result.hasIssue) issues++;
  });
  SlidesApp.getUi().alert('照合完了：' + reviewed + '投稿中 ' + issues + '投稿で要確認');
}

function clearFindings() {
  SlidesApp.getActivePresentation().getSlides().forEach(clearFindingsOnSlide);
  SlidesApp.getUi().alert('結果コメントを削除しました');
}

/* ---------- 構造診断（最初に実行） ---------- */

function diagnose(returnUrl) {
  var ks = keyedSlides();
  var doc = DocumentApp.create('AIレビュー_構造診断_' + new Date().toISOString().slice(0, 16));
  var body = doc.getBody();
  body.appendParagraph('ネイルホリック CRシート 構造診断').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('スライド数: ' + ks.length);
  body.appendParagraph('凡例: type=種別 / key=投稿ID / T=テキストshape数 / IMG=画像数 / codes=テキストから読めた色番号');
  body.appendParagraph('');

  var counts = { CONTENT: 0, CONT: 0, LIST: 0, TEMPLATE: 0 };
  ks.forEach(function (g) {
    counts[g.type] = (counts[g.type] || 0) + 1;
    var slide = g.slide;
    var nText = slide.getShapes().filter(function (sh) { try { return !!sh.getText().asString().trim(); } catch (e) { return false; } }).length;
    var nImg = slide.getImages().length;
    var codes = extractCodes(slideText(slide));
    var first = norm(slideText(slide)).split('\n')[0] || '';
    body.appendParagraph(
      'p' + (g.index + 1) + ' [' + g.type + ']' +
      ' key=' + (g.key || '-') +
      ' T=' + nText + ' IMG=' + nImg +
      ' codes=' + (codes.join(',') || '-') +
      '  | ' + first.slice(0, 40)
    );
  });

  body.appendParagraph('');
  body.appendParagraph('種別集計: CONTENT=' + counts.CONTENT + ' CONT=' + counts.CONT +
    ' LIST=' + counts.LIST + ' TEMPLATE=' + counts.TEMPLATE);
  doc.saveAndClose();
  if (returnUrl) return doc.getUrl();
  SlidesApp.getUi().alert('診断を出力しました。このドキュメントのURLを共有してください:\n' + doc.getUrl());
}
