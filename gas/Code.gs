/**
 * ネイルホリック CRシート AIレビュー — Google Slides ワンクリック照合（プロトタイプ）
 *
 * 前提（元ブリーフより）：CRシート(Googleスライド)の中に
 *   ・「文字入れ指示内容」テキスト  … 正解（＝文字データなので正確に取れる）
 *   ・「貼り付けページ」の提出クリエイティブ … 画像（OCRで文字を読む）
 * が同居している。本スクリプトはスライド上でワンクリック照合し、
 * 結果を「コメント」（テキストボックス）とスピーカーノートに書き出す。
 *
 * 設計方針：AIに正誤を判断させない。OCRは「文字を読む」だけ。
 *           一致判定は機械が文字単位で照合する（照合であって判断ではない）。
 *
 * セットアップは gas/README.md を参照（Drive 詳細サービスの有効化が必要）。
 */

var CONFIG = {
  // 色番号の書式： 英大文字2 + 数字3 + 任意の英大文字1  例) BL920, WT045R, PU076D
  codePattern: /[A-Z]{2}\d{3}[A-Z]?/g,

  ocrLanguage: 'ja',          // Drive OCR の言語
  writeComment: true,         // スライド上にテキストボックスで結果を出す
  writeSpeakerNotes: true,    // スピーカーノートに詳細を出す

  // 正解テキストの探し方：
  //  'activeThenPrev' … アクティブスライドのテキストを正解にする。
  //                      テキストが無ければ直前スライド（指示ページ）を使う。
  truthSource: 'activeThenPrev',

  commentTitle: '🔎 AI一次チェック（照合結果）'
};

/** メニュー登録（ファイルを開くと「AIレビュー」メニューが出る） */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('AIレビュー')
    .addItem('このスライドを照合', 'reviewActiveSlide')
    .addItem('全スライドを照合', 'reviewAllSlides')
    .addSeparator()
    .addItem('直近の結果コメントを消す', 'clearFindings')
    .addToUi();
}

/** 文字列正規化：全角半角・空白のゆれのみ吸収（色番号の1文字差は保持） */
function norm(s) {
  if (!s) return '';
  s = s.normalize('NFKC').replace(/　/g, ' ');
  return s.replace(/[ \t]+/g, ' ').trim();
}

/** テキストから色番号の集合（重複なし・出現順）を取り出す */
function extractCodes(text) {
  var m = norm(text).match(CONFIG.codePattern) || [];
  var seen = {}, out = [];
  m.forEach(function (c) { if (!seen[c]) { seen[c] = 1; out.push(c); } });
  return out;
}

/** スライド上の全テキスト（シェイプ内テキスト）を連結して返す */
function slideText(slide) {
  var parts = [];
  slide.getShapes().forEach(function (sh) {
    try {
      var t = sh.getText().asString();
      if (t) parts.push(t);
    } catch (e) { /* テキストを持たないシェイプは無視 */ }
  });
  // 表の中のテキストも拾う
  slide.getTables().forEach(function (tbl) {
    for (var r = 0; r < tbl.getNumRows(); r++) {
      for (var c = 0; c < tbl.getNumColumns(); c++) {
        try { parts.push(tbl.getCell(r, c).getText().asString()); } catch (e) {}
      }
    }
  });
  return parts.join('\n');
}

/** 画像 Blob を Drive の OCR に通して文字列を得る（Drive 詳細サービスを使用） */
function ocrBlob(blob) {
  var tmp = Drive.Files.insert(
    { title: '__ocr_tmp__', mimeType: 'application/vnd.google-apps.document' },
    blob,
    { ocr: true, ocrLanguage: CONFIG.ocrLanguage }
  );
  var text = '';
  try {
    text = DocumentApp.openById(tmp.id).getBody().getText();
  } finally {
    try { Drive.Files.remove(tmp.id); } catch (e) {}
  }
  return text;
}

/** スライド上の全画像をOCRし、読めた色番号(確信の代わりに重複回数)を集計 */
function ocrCodesOnSlide(slide) {
  var images = slide.getImages();
  var count = {};
  var token = ScriptApp.getOAuthToken();
  images.forEach(function (img) {
    var blob;
    try {
      var url = img.getContentUrl();
      blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } }).getBlob();
    } catch (e) { return; }
    var text = ocrBlob(blob);
    extractCodes(text).forEach(function (c) { count[c] = (count[c] || 0) + 1; });
  });
  return count; // {code: 検出画像数}
}

/** アクティブスライド（未選択時は先頭）を返す */
function getActiveSlide() {
  var pres = SlidesApp.getActivePresentation();
  var sel = pres.getSelection();
  var page = sel && sel.getCurrentPage();
  if (page && page.getPageType && page.getPageType() === SlidesApp.PageType.SLIDE) {
    return page.asSlide();
  }
  var slides = pres.getSlides();
  return slides.length ? slides[0] : null;
}

/** 正解の色番号集合を求める（CONFIG.truthSource に従う） */
function getTruthCodes(slide, index, slides) {
  var codes = extractCodes(slideText(slide));
  if (codes.length === 0 && CONFIG.truthSource === 'activeThenPrev' && index > 0) {
    codes = extractCodes(slideText(slides[index - 1])); // 直前の「指示ページ」を正解に
  }
  return codes;
}

/** 1スライドを照合してレポート文字列と判定を返す */
function reviewSlide(slide, index, slides) {
  var truth = getTruthCodes(slide, index, slides);
  var ocr = ocrCodesOnSlide(slide);                 // {code: count}
  var ocrCodes = Object.keys(ocr);

  var truthSet = {}; truth.forEach(function (c) { truthSet[c] = 1; });
  var ocrSet = {}; ocrCodes.forEach(function (c) { ocrSet[c] = 1; });

  var matched = truth.filter(function (c) { return ocrSet[c]; });
  var missing = truth.filter(function (c) { return !ocrSet[c]; });     // 正解にあるがOCR未検出
  var unknown = ocrCodes.filter(function (c) { return !truthSet[c]; }); // OCRにあるが正解に無い＝取り違え疑い

  var ok = (truth.length > 0) && missing.length === 0 && unknown.length === 0;

  var lines = [];
  lines.push(CONFIG.commentTitle);
  if (truth.length === 0) {
    lines.push('※ このスライドから正解の色番号を取得できませんでした（指示テキストが見つからない）。');
  }
  lines.push('正解(指示): ' + (truth.join(', ') || '（なし）'));
  lines.push('OCR(提出物): ' + (ocrCodes.join(', ') || '（なし）'));
  lines.push('一致: ' + matched.length + '/' + truth.length);
  if (unknown.length) lines.push('❌ 取り違え疑い（正解に無い）: ' + unknown.join(', '));
  if (missing.length) lines.push('⚠ 未検出（正解にあるが読めず/未反映）: ' + missing.join(', '));
  if (ok) lines.push('✅ 色番号は正解と一致');

  return { ok: ok, hasIssue: unknown.length > 0 || missing.length > 0, text: lines.join('\n') };
}

/** 結果をスライドへ書き出す（テキストボックス＋スピーカーノート） */
function writeFindings(slide, result) {
  if (CONFIG.writeSpeakerNotes) {
    try { slide.getNotesPage().getSpeakerNotesShape().getText().setText(result.text); } catch (e) {}
  }
  if (CONFIG.writeComment) {
    clearFindingsOnSlide(slide);
    var box = slide.insertTextBox(result.text, 12, 12, 320, 120);
    box.setTitle('__ai_review__'); // 後で消せるように印を付ける
    var tr = box.getText();
    tr.getTextStyle().setFontSize(9).setForegroundColor(result.hasIssue ? '#C0392B' : '#1E8449');
  }
}

/** 印付きの結果テキストボックスを1スライドから消す */
function clearFindingsOnSlide(slide) {
  slide.getShapes().forEach(function (sh) {
    try { if (sh.getTitle() === '__ai_review__') sh.remove(); } catch (e) {}
  });
}

/** メニュー：アクティブスライドを照合 */
function reviewActiveSlide() {
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  var slide = getActiveSlide();
  if (!slide) { SlidesApp.getUi().alert('スライドが見つかりません'); return; }
  var index = 0;
  for (var i = 0; i < slides.length; i++) { if (slides[i].getObjectId() === slide.getObjectId()) { index = i; break; } }
  var result = reviewSlide(slide, index, slides);
  writeFindings(slide, result);
  SlidesApp.getUi().alert(result.text);
}

/** メニュー：全スライドを照合 */
function reviewAllSlides() {
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  var issues = 0, reviewed = 0;
  for (var i = 0; i < slides.length; i++) {
    if (slides[i].getImages().length === 0) continue; // 画像の無いページ（目次・指示のみ）はスキップ
    var result = reviewSlide(slides[i], i, slides);
    writeFindings(slides[i], result);
    reviewed++;
    if (result.hasIssue) issues++;
  }
  SlidesApp.getUi().alert('照合完了：' + reviewed + 'ページ中 ' + issues + 'ページで要確認');
}

/** メニュー：全スライドの結果コメントを消す */
function clearFindings() {
  SlidesApp.getActivePresentation().getSlides().forEach(clearFindingsOnSlide);
  SlidesApp.getUi().alert('結果コメントを削除しました');
}
