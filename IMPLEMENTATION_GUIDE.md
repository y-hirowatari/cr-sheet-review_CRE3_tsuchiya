# 実装ガイド ― CRシート AIレビュー（指示書 × 提出物の非判断型・機械照合）

このMD **1つで、読んだ人がゼロから同じツールを再現・実装できる**ことを目的にしています。
全ソースコード・セットアップ・デプロイ・カスタマイズ・トラブルシュートを自己完結でまとめています。

- 事例としての解説（背景・設計思想・汎用性）は [PORTFOLIO.md](PORTFOLIO.md) を参照
- 本ガイドは「**動くものを作る**」ことに特化

---

## 0. これは何か

**制作指示書（Googleスライド＝CRシート）に書かれた「色番号（正解）」と、完成クリエイティブ画像をOCRして得た「色番号（実物）」を、AIに判断させず“機械で照合”して食い違いを検出する社内Webアプリ。**

- 実行基盤：Google Apps Script（Webアプリとしてデプロイ）
- 使い方：CRシートのURLを読み込む → 投稿を選ぶ → 完成画像をアップ → 照合
- 出力：❌ 取り違え疑い ／ ⚠ 未検出 ／ ✅ 一致
- 非破壊：CRシートには一切書き込まない

### データフロー
```
① CRシート(Googleスライド) のテキスト → 色番号を抽出 = 正解
② 完成クリエイティブ画像をアップロード → Drive OCR で文字化 = 実物
③ 正規化して集合比較（AIは判断しない）
④ 取り違え / 未検出 / 一致 を証拠つきで表示
```

**設計の核心**：AIの役割は「文字を読む(OCR)」までに限定し、正誤の判定は**機械の集合比較**が行う。これにより生成AIの「自信ありげな誤判定」を回避する。表記ゆれ（全角半角・空白・記号）は正規化で吸収するが、**色番号そのものの1文字差は吸収しない**（取り違えを見逃さないため）。

---

## 1. 前提条件

- **Google Workspace アカウント**（`"access": "DOMAIN"` を使うため。個人gmailは不可）
- 照合したい **CRシートが「ネイティブGoogleスライド」**であること（アップロードしたPowerPoint(.pptx)のままは不可 → 「ファイル→Googleスライドとして保存」で変換）
- CRシートの各投稿ページに、**色番号がテキストとして**入っていること（画像の中の文字ではなく、編集可能テキスト）
- 色番号の書式：英大文字2 + 数字3 + 任意の英大文字1（例 `BL920` `WT045R` `PU076D`）。異なる体系なら `CONFIG.codePattern` を変更

---

## 2. ファイル構成（Apps Scriptプロジェクト内）

```
（スタンドアロンGASプロジェクト）
├─ appsscript.json   … マニフェスト（Webアプリ設定・スコープ・Drive詳細サービス）
├─ コード.gs (Code.gs) … サーバー側（指示書読取・OCR・照合API）
└─ Index.html        … クライアント側（URL入力→投稿選択→画像アップ→結果表示）
```

---

## 3. セットアップ手順（ゼロから）

1. 対象ドメインのアカウントで **https://script.google.com** を開く → **新しいプロジェクト**
2. 既定の `Code.gs` の中身を、後述「4-2 Code.gs」で全置換 → 保存
3. 左「ファイル ＋ → HTML」で **`Index`** を作成 → 「4-3 Index.html」を貼り付け → 保存
   - ⚠ ファイル名は必ず `Index`（`doGet` が `Index` を読む）
4. 左「⚙ プロジェクトの設定 → 『appsscript.json』マニフェストをエディタで表示」をON
5. `appsscript.json` を「4-1 appsscript.json」で全置換 → 保存
   - これで **Drive 詳細サービス(v2)** と Webアプリ設定・スコープが有効化
6. **デプロイ → 新しいデプロイ → 種類「ウェブアプリ」**
   - 次のユーザーとして実行：**ウェブアプリにアクセスしているユーザー**
   - アクセスできるユーザー：**（自ドメイン）内の全員**
7. 権限承認（初回のみ。未確認アプリ警告は自分のスクリプトなので「詳細→移動」で進む）
8. 発行された **ウェブアプリURL** を開く／共有

> 更新を反映：**デプロイ → デプロイを管理 → 編集(鉛筆) → バージョン「新バージョン」→ デプロイ**

---

## 4. ソースコード（全文）

### 4-1. `appsscript.json`

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": [
      { "userSymbol": "Drive", "version": "v2", "serviceId": "drive" }
    ]
  },
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "DOMAIN"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

> `"access"` は所有アカウントのドメインに自動で紐づく。別ドメインの人にも使わせたい場合は
> `"ANYONE_WITH_GOOGLE_ACCOUNT"` に変更（URLを知る全Googleユーザーが対象になる点に注意）。

### 4-2. `Code.gs`（サーバー側）

```javascript
/**
 * ネイルホリック CRシート AIレビュー — 社内Webアプリ版（スタンドアロンGAS / B案）
 *
 * フロー（担当者決定）：
 *   1. アプリにCRシートのURLを入れて読み込む → 投稿一覧＋各投稿の「正解の色番号」を取得
 *   2. レビューしたい投稿を選ぶ（指示＝正解は常にCRシート内にある）
 *   3. 貼り付ける“前”の完成クリエイティブ画像をアプリにアップロード
 *   4. アップ画像をOCR → 正解と機械照合 → 結果を画面表示
 *
 * 方針：CRシートは【読み取り専用】（書き込まない・納品フローを変えない）。
 *       AIに正誤判断はさせない。OCRは「読む」だけ、一致判定は機械が文字単位で行う。
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
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim();
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

function truthOfGroup(group) {
  var truth = [], seen = {};
  group.forEach(function (g) {
    extractCodes(slideText(g.slide)).forEach(function (c) { if (!seen[c]) { seen[c] = 1; truth.push(c); } });
  });
  return truth;
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

/** STEP1：CRシートを読み込み、投稿一覧＋各投稿の正解色番号を返す（OCRしないので軽い） */
function api_load(url) {
  var pres = openPres(url);
  var ks = keyedSlides(pres);
  var groups = {}, order = [];
  ks.forEach(function (g) {
    if (!g.key) return;
    if (!groups[g.key]) { groups[g.key] = { key: g.key, title: g.title, pages: [], slides: [] }; order.push(g.key); }
    groups[g.key].pages.push(g.index + 1);
    groups[g.key].slides.push(g);
  });
  var posts = order.map(function (k) {
    var grp = groups[k];
    return { key: grp.key, title: grp.title, pages: grp.pages, truth: truthOfGroup(grp.slides) };
  });
  return { title: pres.getName(), slideCount: ks.length, posts: posts };
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
```

### 4-3. `Index.html`（クライアント側UI）

```html
<!DOCTYPE html>
<html>
<head>
<base target="_top">
<meta charset="utf-8">
<style>
  :root{
    --ink:#241c22; --ink2:#6c5c64; --line:#ecdde4; --bg:#fbf7f8; --card:#fff;
    --rose:#c63d6e;
    --good:#2e7d5b; --goodbg:#e5f1eb;
    --warn:#8a5d10; --warnbg:#f6ecd6;
    --crit:#b23a34; --critbg:#f7e0dd;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:"Hiragino Kaku Gothic ProN","Yu Gothic","Noto Sans JP",system-ui,sans-serif;
    color:var(--ink);background:var(--bg);font-size:14px;line-height:1.7}
  .wrap{max-width:720px;margin:0 auto;padding:24px 18px 60px}
  h1{font-size:19px;margin:0 0 2px;font-weight:800}
  .sub{font-size:12px;color:var(--ink2);margin:0 0 18px}
  .step{font-size:11px;font-weight:800;letter-spacing:.08em;color:var(--rose);margin:0 0 8px}
  .panel{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px 18px;margin-bottom:14px}
  label{font-size:12px;font-weight:700;color:var(--ink2);display:block;margin-bottom:6px}
  select,input[type=text],button{font-family:inherit;font-size:13.5px;border-radius:9px;border:1px solid var(--line);background:#fff;color:var(--ink);padding:10px 11px}
  input[type=text],select{width:100%}
  select{font-weight:700;cursor:pointer}
  .rowflex{display:flex;gap:8px}
  .rowflex input{flex:1}
  .rowflex button{flex:none}
  button.go{font-weight:800;background:var(--rose);border-color:var(--rose);color:#fff;cursor:pointer;width:100%;padding:12px}
  button.go:hover{filter:brightness(1.06)}
  button.go:disabled{opacity:.55;cursor:default}
  button.ghost{background:#fff;color:var(--ink);cursor:pointer;font-weight:700}
  button.ghost:hover{border-color:var(--rose);color:var(--rose)}
  .field{margin-bottom:13px}
  .field:last-child{margin-bottom:0}
  .meta{font-size:12px;color:var(--ink2);margin:0 0 12px}
  .meta b{color:var(--ink)}
  .truthbox{font-size:12px;color:var(--ink2);background:#faf6f8;border:1px solid var(--line);border-radius:9px;padding:9px 11px;margin-top:4px}
  .drop{border:1.5px dashed var(--line);border-radius:11px;padding:18px;text-align:center;color:var(--ink2);font-size:12.5px;cursor:pointer;background:#fff}
  .drop.hl{border-color:var(--rose);background:#fdf3f7}
  .files{font-size:12px;color:var(--ink);margin-top:8px}
  .files span{display:inline-block;background:#f2e8ee;border-radius:6px;padding:2px 8px;margin:3px 4px 0 0}
  .status{font-size:12.5px;color:var(--ink2);margin:6px 0 12px;min-height:18px}
  .spin{display:inline-block;width:12px;height:12px;border:2px solid var(--line);border-top-color:var(--rose);border-radius:50%;animation:s .7s linear infinite;vertical-align:-2px;margin-right:7px}
  @keyframes s{to{transform:rotate(360deg)}}
  .hide{display:none}

  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px;margin-bottom:10px}
  .chead{display:flex;align-items:center;gap:9px;margin-bottom:8px}
  .badge{font-size:11px;font-weight:800;padding:2px 9px;border-radius:999px;flex:none}
  .badge.ng{background:var(--critbg);color:var(--crit)}
  .badge.ok{background:var(--goodbg);color:var(--good)}
  .badge.info{background:#eee;color:#555}
  .ctitle{font-size:13px;font-weight:700;min-width:0;word-break:break-word}
  .ctitle small{display:block;font-weight:400;color:var(--ink2);font-size:11px}
  .count{font-size:12px;color:var(--ink2);margin:2px 0 9px}
  .grp{margin-top:8px}
  .glabel{font-size:12px;font-weight:700;margin-bottom:5px}
  .glabel.crit{color:var(--crit)} .glabel.warn{color:var(--warn)} .glabel.good{color:var(--good)}
  .chips{display:flex;flex-wrap:wrap;gap:6px}
  .chip{font-family:"DejaVu Sans Mono",monospace;font-size:12px;font-weight:600;padding:3px 8px;border-radius:6px;border:1px solid var(--line);background:#faf6f8}
  .chip.crit{background:var(--critbg);color:var(--crit);border-color:transparent}
  .chip.warn{background:var(--warnbg);color:var(--warn);border-color:transparent}
  .chip.good{background:var(--goodbg);color:var(--good);border-color:transparent}
  .note{font-size:12px;color:var(--warn);background:var(--warnbg);border-radius:8px;padding:7px 10px;margin-top:7px}
  .empty{font-size:13px;color:var(--ink2);text-align:center;padding:20px 8px}
  .legend{font-size:11.5px;color:var(--ink2);margin-top:14px;line-height:1.8}
  .legend b.c{color:var(--crit)} .legend b.w{color:var(--warn)} .legend b.g{color:var(--good)}
</style>
</head>
<body>
<div class="wrap">
  <h1>CRシート AIレビュー</h1>
  <p class="sub">指示（正解）はCRシートから読み込み、貼り付け前の完成画像をアップして色番号を照合します。シートには書き込みません。</p>

  <!-- STEP 1 -->
  <div class="panel">
    <div class="step">STEP 1 ・ CRシートを読み込む</div>
    <label>CRシート（Googleスライド）のURL</label>
    <div class="rowflex">
      <input id="url" type="text" placeholder="https://docs.google.com/presentation/d/........./edit" />
      <button class="ghost" id="load" onclick="load()">読み込み</button>
    </div>
  </div>

  <!-- STEP 2 + 3 -->
  <div class="panel hide" id="step2">
    <div class="step">STEP 2 ・ 投稿を選ぶ</div>
    <div class="meta" id="sheetMeta"></div>
    <div class="field">
      <label>レビューする投稿</label>
      <select id="post" onchange="onPost()"></select>
      <div class="truthbox" id="truthbox"></div>
    </div>

    <div class="step" style="margin-top:16px">STEP 3 ・ 完成画像をアップして照合</div>
    <div class="field">
      <div class="drop" id="drop">クリックで画像を選択（複数可）／ここにドラッグ＆ドロップ<br><span style="font-size:11px">PNG / JPG。貼り付け“前”の完成クリエイティブ</span></div>
      <input id="file" type="file" accept="image/*" multiple class="hide" />
      <div class="files" id="files"></div>
    </div>
    <button class="go" id="go" onclick="go()" disabled>▶ 照合する</button>
  </div>

  <div class="status" id="status"></div>
  <div id="results"></div>

  <div class="legend hide" id="legend">
    <b class="c">❌ 取り違え疑い</b>：画像にあるが正解に無い色番号<br>
    <b class="w">⚠ 未検出</b>：正解にあるが画像で読めず／載せ忘れ<br>
    <b class="g">✅ 一致</b>：正解と画像で一致
  </div>
</div>

<script>
  var POSTS = [], FILES = [];
  function esc(s){return String(s).replace(/[&<>]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;'}[c];});}
  function el(id){return document.getElementById(id);}
  function busy(b,msg){el('go').disabled=b||!FILES.length;el('load').disabled=b;
    el('status').innerHTML=b?('<span class="spin"></span>'+esc(msg||'処理中…')):esc(msg||'');}
  function url(){return el('url').value.trim();}

  /* STEP1 */
  function load(){
    if(!url()){busy(false,'URLを入力してください');return;}
    el('results').innerHTML=''; busy(true,'CRシートを読み込み中…');
    google.script.run.withSuccessHandler(onLoaded).withFailureHandler(fail).api_load(url());
  }
  function onLoaded(info){
    busy(false,'');
    POSTS=info.posts||[];
    el('step2').classList.remove('hide');
    el('legend').classList.remove('hide');
    el('sheetMeta').innerHTML='<b>'+esc(info.title)+'</b>（'+info.slideCount+'ページ / 投稿 '+POSTS.length+'件）';
    var sel=el('post'); sel.innerHTML='';
    POSTS.forEach(function(p,i){
      var o=document.createElement('option'); o.value=i;
      o.textContent=p.key+'  '+(p.title||'');
      sel.appendChild(o);
    });
    onPost();
  }
  function onPost(){
    var p=POSTS[el('post').value]; if(!p)return;
    el('truthbox').innerHTML='この投稿の正解の色番号（'+p.truth.length+'）：'+
      (p.truth.length?p.truth.map(function(c){return '<span class="chip">'+esc(c)+'</span>';}).join(' '):'（テキストから取得できず）');
  }

  /* STEP3: ファイル選択 */
  el('drop').onclick=function(){el('file').click();};
  el('file').onchange=function(){pick(this.files);};
  ['dragenter','dragover'].forEach(function(ev){el('drop').addEventListener(ev,function(e){e.preventDefault();el('drop').classList.add('hl');});});
  ['dragleave','drop'].forEach(function(ev){el('drop').addEventListener(ev,function(e){e.preventDefault();el('drop').classList.remove('hl');});});
  el('drop').addEventListener('drop',function(e){pick(e.dataTransfer.files);});

  function pick(fileList){
    var arr=Array.prototype.slice.call(fileList).filter(function(f){return /^image\//.test(f.type);});
    if(!arr.length)return;
    FILES=[]; var done=0;
    el('files').innerHTML='読み込み中…';
    arr.forEach(function(f){
      var r=new FileReader();
      r.onload=function(){FILES.push({name:f.name,mime:f.type||'image/png',b64:r.result.split(',')[1]}); if(++done===arr.length)filesReady();};
      r.onerror=function(){if(++done===arr.length)filesReady();};
      r.readAsDataURL(f);
    });
  }
  function filesReady(){
    el('files').innerHTML=FILES.map(function(f){return '<span>'+esc(f.name)+'</span>';}).join('');
    el('go').disabled=false;
  }

  /* 照合 */
  function go(){
    if(!FILES.length){busy(false,'画像を選んでください');return;}
    var p=POSTS[el('post').value];
    el('results').innerHTML=''; busy(true,'照合中…（OCRに数十秒かかることがあります）');
    google.script.run.withSuccessHandler(render).withFailureHandler(fail)
      .api_reviewUpload({key:p.key,title:p.title,truth:p.truth,files:FILES});
  }
  function fail(e){busy(false,'');el('results').innerHTML='<div class="note">エラー: '+esc(e&&e.message||e)+'</div>';}

  function chips(a,cls){if(!a||!a.length)return '';return '<div class="chips">'+a.map(function(c){return '<span class="chip '+cls+'">'+esc(c)+'</span>';}).join('')+'</div>';}
  function card(r){
    var badge=r.truthEmpty?'<span class="badge info">情報</span>':(r.hasIssue?'<span class="badge ng">要確認</span>':'<span class="badge ok">一致</span>');
    var h='<div class="card"><div class="chead">'+badge+'<div class="ctitle">投稿 '+esc(r.key)+'<small>'+(r.title?'【'+esc(r.title)+'】 ':'')+'アップ画像 '+r.images+'枚</small></div></div>';
    h+='<div class="count">一致 '+r.matched.length+' / 正解 '+r.truth.length+'（画像OCR検出 '+r.ocr.length+'）</div>';
    if(r.truthEmpty)h+='<div class="note">正解の色番号を取得できませんでした。指示テキストのある投稿か確認してください。</div>';
    if(r.unknown.length)h+='<div class="grp"><div class="glabel crit">❌ 取り違え疑い</div>'+chips(r.unknown,'crit')+'</div>';
    if(r.missing.length)h+='<div class="grp"><div class="glabel warn">⚠ 未検出</div>'+chips(r.missing,'warn')+'</div>';
    if(r.matched.length)h+='<div class="grp"><div class="glabel good">✅ 一致</div>'+chips(r.matched,'good')+'</div>';
    return h+'</div>';
  }
  function render(r){busy(false,'完了');el('results').innerHTML=card(r);}
</script>
</body>
</html>
```

---

## 5. 設定とカスタマイズ（`CONFIG`）

`Code.gs` 冒頭の `CONFIG` を変えるだけで挙動を調整できる。

| 項目 | 既定 | 説明 |
| :- | :- | :- |
| `codePattern` | `/[A-Z]{2}\d{3}[A-Z]?/g` | 検出する識別子の書式。型番・SKU等に合わせて変更 |
| `ocrLanguage` | `'ja'` | Drive OCR の言語 |

投稿の識別・グルーピングは `postIdsOf()` / `classify()` を、正解の抽出範囲は `truthOfGroup()` を、
照合結果の組み立ては `buildResult()` を編集する。**照合の考え方（正規化するが1文字差は残す）は `norm()` / `extractCodes()` に集約**。

---

## 6. 動作の仕組み（照合ロジック）

1. `api_load(url)`：スライドを開き、`keyedSlides()` で「投稿」単位にグルーピング。目次(LIST)・テンプレは除外。各投稿の指示テキストから `extractCodes()` で**正解の色番号**を得る。
2. `api_reviewUpload(data)`：アップされた画像(base64)を `ocrBlob()`（Drive OCR）で文字化 → `extractCodes()` で**実物の色番号**を得る。
3. `buildResult()`：正解と実物を集合比較し、
   - `unknown`（実物にあるが正解に無い）＝ **取り違え疑い**
   - `missing`（正解にあるが実物に無い）＝ **未検出（載せ忘れ/OCR読み落とし）**
   - `matched` ＝ 一致
4. 結果はクライアント(`Index.html`)が色分けチップで描画。**シートには書き込まない**。

---

## 7. 拡張ポイント

- **照合対象を増やす**：色番号だけでなく色名・素材名・必須文言（CTA/価格/発売日/SPF等）。`extractCodes()` を「対象トークンの抽出」に拡張し、`buildResult()` を項目別に。
- **OCR精度向上**：`ocrBlob()` を Google Drive OCR から **Cloud Vision API** に差し替え（装飾フォント・写真上の白文字に強い。GCP設定が必要）。
- **格納物の一致確認（先祖還り対策）**：貼り付け画像と、Drive格納フォルダの最終ファイルをハッシュ/OCRで突き合わせ。
- **上流照合**：クライアント発注書（別Slides/スプシ）を正解ソースにし、CRシート指示との食い違い（記入ミス）も検出。
- **他ドメインへ転用**：`codePattern` と抽出/グルーピングを差し替えるだけで、EC型番・医薬表現・帳票検収などに応用可能。

---

## 8. トラブルシュート

| 症状 | 原因 / 対処 |
| :- | :- |
| `スライドを開けませんでした` | ①ログイン中アカウントに**閲覧権限が無い** → 同アカウントでURLを直接開けるか確認／共有する。②**.pptxのまま**（ネイティブSlidesでない）→ Googleスライドに変換。③URL/IDが不正 |
| `Ui.showSidebar 権限` 系（旧・メニュー版） | Webアプリ版では発生しない。旧版は `script.container.ui` スコープが必要 |
| `Drive is not defined` | Drive 詳細サービス未有効 → マニフェストの `enabledAdvancedServices`、または「サービス → ＋ → Drive API(v2, 識別子 Drive)」 |
| `HTML file not found` | HTMLファイル名が `Index` でない |
| OCRが遅い/一部読めない | 画像枚数を絞る／解像度を上げる／Cloud Visionへ差し替え。写真上の細い白文字は誤読しやすい（Phase 1の重点検証点） |
| 未確認アプリ警告 | 自作スクリプトのため「詳細→(プロジェクト名)に移動」で承認 |

---

## 9. （付録）Python PoC ― 提出物が無くても核心を先行検証したい人向け

本番はGASだが、**照合の考え方は言語非依存**。手元でロジックを検証したい場合の参考実装。

### 9-1. `cr_check.py`（CRシート単体の整合チェック／PDF入力）

依存：`pip install pymupdf`

```python
#!/usr/bin/env python3
"""
ネイルホリック CRシート 整合チェック PoC (v0)

提出物（クリエイティブ画像）が無くても、CRシート単体で検証できる範囲の
「機械的な照合」を実行する。設計方針は docs/AIレビューの仕組み_設計書_チーム共有用.md、
仕様は docs/照合ロジック仕様_たたき台.md を参照。

v0 でやること（提出物なしで検証可能）:
  1. 色番号↔色ライン名の整合チェック（シート全体）
       同じ色番号が別のライン名で書かれていないか＝取り違え/誤記の検出
  2. キャプション使用色 ⊆（文字入れ指示 ∪ 対象商材）の包含チェック（投稿単位）
       キャプションにあるのにシートのどこにも無い色番号＝キャプション側の取り違え
  3. インベントリ出力（投稿一覧・抽出色番号・対象商材カバレッジ）
  --demo-typo: 正解に1文字だけ違う色番号を混ぜ、機械照合が検出できることを実演

v0 でやらないこと:
  - 提出画像のOCR照合（v1。実際の提出画像が必要）
  - 配色/アート出典の判定（参考どまり）

使い方:
  python3 cr_check.py <CRシート.pdf> [<CRシート.pdf> ...] [--demo-typo]
"""
import sys
import re
import unicodedata
from collections import defaultdict

CODE_RE = re.compile(r"[A-Z]{2}\d{3}[A-Z]?")
POST_RE = re.compile(r"投稿no,\s*(\d+)_\s*(\d+)月\s*(\d+)日")
TITLE_RE = re.compile(r"投稿no,.*?【(.+?)】")
CODE_LINE_RE = re.compile(r"^([A-Z]{2}\d{3}[A-Z]?)[ 　]+(.+)$")

# ライン名として扱わない語（対象商材欄の定型句など）
NON_LINE = ("商材画像", "格納先", "色玉", "必要素材", "背景画像")


def norm(s: str) -> str:
    """記号・空白・全角半角の表記ゆれのみ吸収（色番号そのものは変えない）。"""
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("　", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def norm_line(name: str) -> str:
    """ライン名の正規化：囲み記号・前後の記号を落とす。"""
    name = norm(name)
    name = name.strip("()（）「」[]【】 　:：・")
    return name.strip()


def codes_in(text: str):
    return CODE_RE.findall(norm(text))


def parse(pdf_path):
    import fitz
    doc = fitz.open(pdf_path)
    posts = {}
    medium = None

    def get_post(medium, text):
        m = POST_RE.search(norm(text))
        if not m:
            return None
        key = (medium, int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if key not in posts:
            posts[key] = {
                "key": key, "medium": medium, "title": None,
                "instr_codes": set(), "shohin_codes": set(), "caption_codes": set(),
                "code_line": {}, "has_caption": False, "has_instr": False,
            }
        return posts[key]

    for pg in doc:
        raw = pg.get_text()
        n = norm(raw)
        stripped = n.strip()

        # 媒体の切り替え（"IG" / "X" の区切りページ）
        if stripped in ("IG", "X"):
            medium = stripped
            continue
        if medium is None:
            continue

        is_cap = "投稿キャプション" in n
        is_ins = "文字入れ指示内容" in n
        if not (is_cap or is_ins):      # 目次・一覧ページは投稿化しない
            continue

        post = get_post(medium, raw)
        if post is None:
            continue

        tm = TITLE_RE.search(n)
        if tm and not post["title"]:
            post["title"] = tm.group(1)

        if is_cap:
            post["has_caption"] = True
            body = n.split("【修正版】", 1)[1] if "【修正版】" in n else n
            post["caption_codes"] |= set(codes_in(body))
        elif is_ins:
            post["has_instr"] = True
            instr_part = n.split("デザイン", 1)[0] if "デザイン" in n else n
            shohin_part = n.split("対象商材", 1)[1] if "対象商材" in n else ""
            for raw_line in instr_part.splitlines():
                line = norm(raw_line)
                m = CODE_LINE_RE.match(line)
                if m:
                    code, name = m.group(1), norm_line(m.group(2))
                    post["instr_codes"].add(code)
                    if name and not any(w in name for w in NON_LINE):
                        post["code_line"].setdefault(code, set()).add(name)
                else:
                    post["instr_codes"] |= set(codes_in(line))
            post["shohin_codes"] |= set(codes_in(shohin_part))

    return posts


def collect(pdfs):
    all_posts, gmap, gsrc = [], defaultdict(set), defaultdict(lambda: defaultdict(set))
    for pdf in pdfs:
        posts = parse(pdf)
        src = pdf.split("/")[-1]
        # ファイル名から "6月"/"7月" を推定できないので通し番号ラベル
        for key in sorted(posts, key=lambda k: (k[0], k[2], k[3], k[1])):
            p = posts[key]
            p["src"] = src
            all_posts.append(p)
            lbl = f'{p["medium"]} no,{key[1]} {key[2]}/{key[3]}'
            for code, names in p["code_line"].items():
                for nm in names:
                    gmap[code].add(nm)
                    gsrc[code][nm].add(lbl)
    return all_posts, gmap, gsrc


def report(all_posts, gmap, gsrc):
    line = "=" * 72
    print(line)
    print(" ネイルホリック CRシート 整合チェック PoC (v0)")
    print(line)

    # 【1】色番号↔ライン名
    print("\n【1】色番号↔色ライン名の整合チェック（シート全体）")
    conflicts = {c: v for c, v in gmap.items() if len(v) > 1}
    if not conflicts:
        print(f"  OK: 色番号 {len(gmap)} 種すべてでライン名の矛盾なし")
    else:
        print(f"  NG: {len(conflicts)} 件の色番号でライン名が食い違い（取り違えの可能性）")
        for c in sorted(conflicts):
            print(f"    - {c}:")
            for nm in sorted(conflicts[c]):
                print(f"        「{nm}」 ← {', '.join(sorted(gsrc[c][nm]))}")

    # 【2】キャプション ⊆ 指示∪商材
    print("\n【2】キャプション使用色 ⊆（文字入れ指示 ∪ 対象商材）（投稿単位）")
    print("     ※キャプションにあるがシートのどこにも無い色番号＝キャプション側の取り違え疑い")
    issues = 0
    for p in all_posts:
        if not p["has_caption"]:
            continue
        sheet = p["instr_codes"] | p["shohin_codes"]
        if not sheet:                       # 照合対象なし（X の転載キャプション等）
            continue
        missing = p["caption_codes"] - sheet
        if missing:
            issues += 1
            k = p["key"]
            print(f'  NG: {p["medium"]} no,{k[1]} {k[2]}/{k[3]} 【{p["title"]}】')
            print(f"        シートに無い色番号: {', '.join(sorted(missing))}")
    if not issues:
        print("  OK: 照合対象の全投稿で、キャプションの色番号はシート側に存在")

    # 【3】インベントリ
    print("\n【3】投稿インベントリ")
    print(f"  {'媒体':<4}{'投稿':<26}{'指示':>4}{'商材':>4}{'ｷｬﾌﾟ':>5}  商材カバレッジ")
    for p in all_posts:
        k = p["key"]
        lbl = f'no,{k[1]} {k[2]}/{k[3]} {p["title"] or ""}'[:25]
        instr = p["instr_codes"]
        cov = f"{len(instr & p['shohin_codes'])}/{len(instr)}" if instr else "-"
        print(f"  {p['medium']:<4}{lbl:<26}{len(instr):>4}{len(p['shohin_codes']):>4}"
              f"{len(p['caption_codes']):>5}  {cov}")
    print(f"\n  ユニーク色番号（全シート合計）: {len(gmap)} 種")
    print(line)


def demo_typo(all_posts):
    """正解セットの1つを1文字だけ書き換えた『提出物/キャプション』を作り、検出できることを実演。"""
    target = max((p for p in all_posts if p["has_instr"]),
                 key=lambda p: len(p["instr_codes"]), default=None)
    if not target:
        return
    truth = sorted(target["instr_codes"] | target["shohin_codes"])
    victim = truth[0]
    # 末尾の数字を1つずらす（BL920 -> BL921 のような一字違いを模擬）
    m = re.match(r"^([A-Z]{2})(\d{3})([A-Z]?)$", victim)
    d = int(m.group(2))
    typo = f"{m.group(1)}{(d+1)%1000:03d}{m.group(3)}"
    submitted = set(truth) - {victim} | {typo}

    k = target["key"]
    print("\n" + "-" * 72)
    print(" 【デモ】正解を1文字だけ書き換えたら検出できるか")
    print("-" * 72)
    print(f"  対象投稿: {target['medium']} no,{k[1]} {k[2]}/{k[3]} 【{target['title']}】")
    print(f"  模擬ミス: 提出物側で {victim} → {typo} と1文字取り違え")
    missing = submitted - (target["instr_codes"] | target["shohin_codes"])
    extra = (target["instr_codes"] | target["shohin_codes"]) - submitted
    if missing or extra:
        print("  ✅ 検出成功:")
        if missing:
            print(f"       提出物にある不明な色番号: {', '.join(sorted(missing))}（正解に存在しない）")
        if extra:
            print(f"       提出物に欠けている色番号: {', '.join(sorted(extra))}（正解に存在）")
        print(f"       → 正しい表記(コピペ用): {victim}")
    else:
        print("  ❌ 検出できませんでした")
    print("-" * 72)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    if not args:
        print(__doc__); sys.exit(1)
    all_posts, gmap, gsrc = collect(args)
    report(all_posts, gmap, gsrc)
    if "--demo-typo" in flags:
        demo_typo(all_posts)


if __name__ == "__main__":
    main()
```

### 9-2. `ocr_match.py`（提出画像OCR照合）

依存：`pip install pytesseract pillow` ＋ システムに `tesseract-ocr`（日本語データ `tesseract-ocr-jpn`）

```python
#!/usr/bin/env python3
"""
ネイルホリック 提出画像 OCR照合 PoC (v1)

提出クリエイティブ画像をOCRで文字化し、CRシート由来の「正解」色番号セットと
機械照合する。設計方針どおり、AIには正誤を判断させず、OCRは「文字を読む」だけ、
一致判定は機械が文字単位で行う。

やること:
  1. 画像をOCR（tesseract, 日本語＋英語）で文字抽出
  2. 正規化（記号・空白・全角半角。色番号の1文字差は吸収しない）
  3. 色番号の抽出と、正解セットとの集合照合（過不足＝取り違え）
  4. 指定した必須文字列（注釈・CTA等）の有無チェック
  5. 証拠つきレポート

使い方:
  python3 ocr_match.py <画像> --truth BL920,BL921,WT045R [--require "SPF50+,PA++++"]
  python3 ocr_match.py <画像> --truth-file codes.txt

依存: tesseract-ocr(+jpn), pytesseract, pillow
"""
import sys
import re
import argparse
import unicodedata

CODE_RE = re.compile(r"[A-Z]{2}\d{3}[A-Z]?")


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("　", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def ocr_passes(image_path):
    """複数パスでOCRし、(全文テキスト, 色番号→{count, conf}) を返す。
    色番号は複数パスでの検出回数と最大確信度を持たせ、OCRノイズ(1パスだけ・低確信)を選別できるようにする。"""
    from PIL import Image, ImageOps
    import pytesseract
    img = Image.open(image_path).convert("RGB")
    w, h = img.size
    scale = 2 if max(w, h) < 1600 else 1
    if scale > 1:
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
    g = ImageOps.autocontrast(ImageOps.grayscale(img))

    variants = [g, img]
    texts, code_stat = [], {}
    n_pass = 0
    for src in variants:
        for psm in (11, 6, 3):
            n_pass += 1
            cfg = f"--oem 1 --psm {psm}"
            texts.append(pytesseract.image_to_string(src, lang="jpn+eng", config=cfg))
            data = pytesseract.image_to_data(src, lang="jpn+eng", config=cfg,
                                             output_type=pytesseract.Output.DICT)
            for word, conf in zip(data["text"], data["conf"]):
                try:
                    c = float(conf)
                except (TypeError, ValueError):
                    c = -1
                for code in CODE_RE.findall(norm(word)):
                    st = code_stat.setdefault(code, {"count": 0, "conf": 0.0})
                    st["count"] += 1
                    st["conf"] = max(st["conf"], c)
    return "\n".join(texts), code_stat, n_pass


def confident_codes(code_stat, min_conf=60.0, min_count=2):
    """確信度・検出回数のどちらかを満たす色番号を『読めた』とみなす。"""
    return {c for c, st in code_stat.items()
            if st["conf"] >= min_conf or st["count"] >= min_count}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("--truth", default="", help="正解の色番号(カンマ区切り)")
    ap.add_argument("--truth-file", default="", help="正解の色番号(1行1件 or カンマ区切り)")
    ap.add_argument("--require", default="", help="必須文字列(カンマ区切り)")
    ap.add_argument("--show-ocr", action="store_true", help="OCR生テキストを表示")
    args = ap.parse_args()

    truth = set()
    if args.truth:
        truth |= {norm(x) for x in args.truth.split(",") if x.strip()}
    if args.truth_file:
        with open(args.truth_file, encoding="utf-8") as f:
            for line in f:
                truth |= {norm(x) for x in re.split(r"[,\s]+", line) if x.strip()}

    raw, code_stat, n_pass = ocr_passes(args.image)
    norm_text = norm(raw)
    found = confident_codes(code_stat)

    line = "=" * 72
    print(line)
    print(" ネイルホリック 提出画像 OCR照合 PoC (v1)")
    print(line)
    print(f"  画像: {args.image}    OCRパス数: {n_pass}")

    if args.show_ocr:
        print("\n--- OCR生テキスト（正規化前）---")
        print(raw.strip()[:2000])
        print("--- ここまで ---")

    def stat(c):
        st = code_stat.get(c, {"count": 0, "conf": 0})
        return f"{c}(conf{int(st['conf'])}/{st['count']}回)"

    print(f"\n【OCRで読めた色番号(確信)】 {len(found)} 件: "
          f"{', '.join(stat(c) for c in sorted(found)) or '（なし）'}")
    low = set(code_stat) - found
    if low:
        print(f"  （参考）低確信で除外: {', '.join(stat(c) for c in sorted(low))}")

    if truth:
        print("\n【色番号の照合（正解 vs OCR）】")
        missing = truth - found           # 正解にあるがOCRで見つからない（未反映 or OCR読み落とし）
        unknown = found - truth           # OCRにあるが正解に無い（取り違え疑い）
        ok = truth & found
        print(f"  一致: {len(ok)}/{len(truth)}  {', '.join(sorted(ok)) or ''}")
        if missing:
            print(f"  ⚠ 正解にあるがOCR未検出: {', '.join(sorted(missing))}")
            print("     → 未反映の可能性、またはOCRの読み落とし（写真上の細字など）。要目視。")
        if unknown:
            print(f"  ❌ OCRにあるが正解に無い(取り違え疑い): "
                  f"{', '.join(stat(c) for c in sorted(unknown))}")
        if not missing and not unknown:
            print("  ✅ 色番号は正解と完全一致")

    if args.require:
        print("\n【必須文字列の有無】")
        for token in [t for t in args.require.split(",") if t.strip()]:
            nt = norm(token)
            hit = nt in norm_text
            print(f"  {'✅' if hit else '❌'} {token}")

    print(line)


if __name__ == "__main__":
    main()
```

---

## 10. ライセンス / 注意

- 社内利用を想定。クライアント素材（CRシート・提出画像）はリポジトリに含めない運用（`.gitignore` 済み）。
- OCRは Google Drive / 各自の Google アカウント権限で実行される（外部送信しない設計）。Cloud Vision 等の外部APIを使う場合はデータ取り扱いポリシーを確認すること。

