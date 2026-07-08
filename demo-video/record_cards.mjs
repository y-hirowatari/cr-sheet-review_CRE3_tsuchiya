import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';

const b64 = f => fs.readFileSync(f).toString('base64');
const pop400 = b64('fonts/fontsource-poppins/files/poppins-latin-400-normal.woff2');
const pop500 = b64('fonts/fontsource-poppins/files/poppins-latin-500-normal.woff2');
const MARK = 'M0,0 Q50,40 100,0 Q60,50 100,100 Q50,60 0,100 Q40,50 0,0 Z M50,16 Q78,22 84,50 Q78,78 50,84 Q22,78 16,50 Q22,22 50,16 Z';

const FONT_CSS = `
@font-face{font-family:Poppins;font-weight:400;src:url(data:font/woff2;base64,${pop400}) format('woff2')}
@font-face{font-family:Poppins;font-weight:500;src:url(data:font/woff2;base64,${pop500}) format('woff2')}`;

// DAOX logo: "DAO" wordmark + star-X mark
const logo = (h, color = '#fff') => `
<div style="display:flex;align-items:center;gap:${h * 0.13}px;line-height:1">
  <span style="font-family:Poppins,sans-serif;font-weight:400;font-size:${h}px;letter-spacing:.015em;color:${color};transform:translateY(${-h * 0.06}px)">DAO</span>
  <svg width="${h * 0.94}" height="${h * 0.94}" viewBox="0 0 100 100" style="display:block"><path d="${MARK}" fill="${color}" fill-rule="evenodd"/></svg>
</div>`;

const VDIR = 'video60/segments-cards';
fs.rmSync(VDIR, { recursive: true, force: true });
fs.mkdirSync(VDIR, { recursive: true });

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const ctx = await browser.newContext({ viewport: { width: 1280, height: 720 }, recordVideo: { dir: VDIR, size: { width: 1280, height: 720 } } });
const sleep = ms => new Promise(r => setTimeout(r, ms));
const meta = [];

async function segment(name, html, dur) {
  const page = await ctx.newPage();
  const t0 = Date.now();
  await page.setContent(html);
  await page.evaluate(() => document.fonts.ready);
  const tStart = Date.now() - t0;
  await sleep(dur);
  const tEnd = Date.now() - t0;
  const video = page.video();
  await page.close();
  meta.push({ name, path: await video.path(), trimStart: tStart / 1000, trimEnd: tEnd / 1000 });
  console.log(`segment ${name}: keep ${tStart}..${tEnd}ms`);
}

const shell = body => `<!doctype html><html><head><style>${FONT_CSS}
@keyframes up{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:none}}
.a{animation:up .5s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.12s}.d2{animation-delay:.3s}.d3{animation-delay:.45s}
body{margin:0;font-family:Poppins,sans-serif}</style></head><body>${body}</body></html>`;

// title (2.3s)
await segment('title', shell(`
<div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:30px;background:radial-gradient(1200px 700px at 50% 35%, #262450 0%, #101018 70%)">
  <div class="a">${logo(96)}</div>
  <div class="a d1" style="color:#c9c8ee;font-size:26px;font-weight:700;font-family:sans-serif">地域コミュニティDAOプラットフォーム</div>
  <div class="a d2" style="color:#8886b8;font-size:17px;font-family:sans-serif">チェックイン ・ タスク ・ 投票 ・ トークン ・ モデレーション</div>
</div>`), 2300);

// section cards (1.2s each)
const sec = (no, title, sub, grad) => shell(`
<div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;background:${grad}">
  <svg class="a" width="40" height="40" viewBox="0 0 100 100" style="opacity:.9"><path d="${MARK}" fill="#fff" fill-rule="evenodd"/></svg>
  <div class="a" style="color:rgba(255,255,255,.55);font-size:19px;font-weight:800;letter-spacing:.3em;font-family:sans-serif">${no}</div>
  <div class="a d1" style="color:#fff;font-size:50px;font-weight:800;font-family:sans-serif">${title}</div>
  <div class="a d1" style="color:rgba(255,255,255,.75);font-size:21px;font-weight:600;font-family:sans-serif">${sub}</div>
</div>`);
await segment('sec-user', sec('01', 'ユーザー画面', '住民・メンバーの体験', 'linear-gradient(135deg,#1a1a30,#3730a3)'), 1200);
await segment('sec-admin', sec('02', '管理者画面', '運営チームのためのダッシュボード', 'linear-gradient(135deg,#2a1224,#9d2960)'), 1200);

// end card (3.0s)
await segment('end', shell(`
<div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:24px;background:radial-gradient(1200px 700px at 50% 35%, #262450 0%, #101018 70%)">
  <div class="a">${logo(78)}</div>
  <div class="a d1" style="color:#c9c8ee;font-size:25px;font-weight:700;font-family:sans-serif">地域コミュニティを、DAOでなめらかに。</div>
  <div class="a d2" style="display:flex;gap:10px;margin-top:4px;font-family:sans-serif">
    ${['QRチェックイン','タスク受発注','投票','トークン','モデレーション'].map(t => `<span style="background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);color:#dedbf5;padding:7px 16px;border-radius:999px;font-size:14.5px;font-weight:600">${t}</span>`).join('')}
  </div>
  <div class="a d3" style="color:#8886b8;font-size:16px;margin-top:6px;font-family:sans-serif">デモ: daox-zeta.vercel.app/home ｜ 管理者: /admin</div>
</div>`), 3000);

await ctx.close();
await browser.close();
fs.writeFileSync('video60/cards_meta.json', JSON.stringify(meta, null, 2));
console.log('DONE');
