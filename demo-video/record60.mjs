import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import { installFetchRoute } from './fetchproxy.mjs';
import fs from 'fs';

const BASE = 'https://daox-zeta.vercel.app';
const VDIR = 'video60/segments';
fs.rmSync('video60', { recursive: true, force: true });
fs.mkdirSync(VDIR, { recursive: true });

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const ctx = await browser.newContext({
  viewport: { width: 1280, height: 720 },
  recordVideo: { dir: VDIR, size: { width: 1280, height: 720 } },
});
await installFetchRoute(ctx);

await ctx.addInitScript(() => {
  const mk = () => {
    if (document.getElementById('__pwcursor')) return;
    const d = document.createElement('div');
    d.id = '__pwcursor';
    d.style.cssText = 'position:fixed;z-index:2147483647;width:26px;height:26px;border-radius:50%;background:rgba(85,80,230,.30);border:2.5px solid rgba(85,80,230,.95);box-shadow:0 2px 10px rgba(85,80,230,.5);pointer-events:none;transform:translate(-50%,-50%);left:-60px;top:-60px;transition:width .1s,height .1s';
    document.documentElement.appendChild(d);
    addEventListener('mousemove', e => { d.style.left = e.clientX + 'px'; d.style.top = e.clientY + 'px'; }, true);
    addEventListener('mousedown', () => { d.style.width = '15px'; d.style.height = '15px'; }, true);
    addEventListener('mouseup', () => { d.style.width = '26px'; d.style.height = '26px'; }, true);
  };
  document.readyState === 'loading' ? addEventListener('DOMContentLoaded', mk) : mk();
});

const sleep = ms => new Promise(r => setTimeout(r, ms));
const meta = [];

// mode badge (top-left) + feature caption (bottom, big title + sub)
async function badge(page, text, color) {
  await page.evaluate(([text, color]) => {
    let b = document.getElementById('__pwbadge');
    if (!b) {
      b = document.createElement('div');
      b.id = '__pwbadge';
      b.style.cssText = 'position:fixed;top:14px;left:14px;z-index:2147483600;color:#fff;padding:7px 16px;border-radius:999px;font-family:sans-serif;font-size:14px;font-weight:700;letter-spacing:.05em;box-shadow:0 4px 14px rgba(0,0,0,.25)';
      document.documentElement.appendChild(b);
    }
    b.textContent = text;
    b.style.background = color;
  }, [text, color]);
}
async function feature(page, title, sub, color = '#5550e6') {
  await page.evaluate(([title, sub, color]) => {
    let c = document.getElementById('__pwcap');
    if (!c) {
      c = document.createElement('div');
      c.id = '__pwcap';
      c.style.cssText = 'position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(14px);display:flex;flex-direction:column;align-items:center;gap:2px;background:rgba(14,14,20,.90);backdrop-filter:blur(8px);color:#fff;padding:12px 30px;border-radius:18px;font-family:sans-serif;z-index:2147483600;box-shadow:0 8px 30px rgba(0,0,0,.4);white-space:nowrap;transition:opacity .25s,transform .25s;opacity:0';
      document.documentElement.appendChild(c);
    }
    c.innerHTML = `<div style="font-size:19px;font-weight:800;letter-spacing:.03em"><span style="color:${color};margin-right:8px">●</span>${title}</div><div style="font-size:13.5px;color:#c8c8d8;font-weight:500">${sub}</div>`;
    c.style.opacity = '0'; c.style.transform = 'translateX(-50%) translateY(14px)';
    requestAnimationFrame(() => requestAnimationFrame(() => { c.style.opacity = '1'; c.style.transform = 'translateX(-50%) translateY(0)'; }));
  }, [title, sub, color]);
}

async function click(page, locator) {
  try {
    const el = locator.locator('visible=true').first();
    const box = await el.boundingBox({ timeout: 4000 });
    if (!box) return false;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2, { steps: 12 });
    await sleep(140);
    await el.click({ timeout: 4000 });
    return true;
  } catch (e) { console.log('click skip:', e.message.split('\n')[0]); return false; }
}

function sectionCard(no, title, sub, grad) {
  return `<!doctype html><html><body style="margin:0"><div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;background:${grad};font-family:sans-serif">
  <style>@keyframes up{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:none}}.a{animation:up .4s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.12s}</style>
  <div class="a" style="color:rgba(255,255,255,.55);font-size:20px;font-weight:800;letter-spacing:.3em">${no}</div>
  <div class="a" style="color:#fff;font-size:50px;font-weight:800">${title}</div>
  <div class="a d1" style="color:rgba(255,255,255,.75);font-size:21px;font-weight:600">${sub}</div>
  </div></body></html>`;
}

async function segment(name, fn) {
  const page = await ctx.newPage();
  const t0 = Date.now();
  let r = { tStart: 0, tEnd: 0 };
  const mark = () => Date.now() - t0;
  try { r = await fn(page, mark); }
  finally {
    const video = page.video();
    await page.close();
    meta.push({ name, path: await video.path(), trimStart: r.tStart / 1000, trimEnd: r.tEnd / 1000 });
    console.log(`segment ${name}: keep ${r.tStart}..${r.tEnd}ms`);
  }
}

const U = '#5550e6', A = '#e6558a';

// ---- A: title (2.2s)
await segment('title', async (page, mark) => {
  await page.setContent(`<!doctype html><html><body style="margin:0"><div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:22px;background:radial-gradient(1200px 700px at 50% 35%, #262450 0%, #101018 70%);font-family:sans-serif">
  <style>@keyframes up{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:none}}.a{animation:up .5s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.12s}.d2{animation-delay:.3s}</style>
  <div class="a" style="display:flex;align-items:center;gap:20px">
    <div style="width:80px;height:80px;border-radius:20px;background:#fff;color:#101018;display:flex;align-items:center;justify-content:center;font-size:48px;font-weight:800">D</div>
    <div style="color:#fff;font-size:60px;font-weight:800;letter-spacing:.04em">DAOX</div>
  </div>
  <div class="a d1" style="color:#c9c8ee;font-size:26px;font-weight:700">地域コミュニティDAOプラットフォーム</div>
  <div class="a d2" style="color:#8886b8;font-size:17px">チェックイン ・ タスク ・ 投票 ・ トークン ・ モデレーション</div>
  </div></body></html>`);
  const tStart = mark(); await sleep(2200); return { tStart, tEnd: mark() };
});

// ---- B: user section card (1.2s)
await segment('sec-user', async (page, mark) => {
  await page.setContent(sectionCard('01', 'ユーザー画面', '住民・メンバーの体験', 'linear-gradient(135deg,#1a1a30,#3730a3)'));
  const tStart = mark(); await sleep(1200); return { tStart, tEnd: mark() };
});

// ---- C: user tour
await segment('user', async (page, mark) => {
  await page.goto(BASE + '/home', { waitUntil: 'load', timeout: 60000 });
  await sleep(1100);
  await badge(page, 'ユーザー画面', U);
  await feature(page, 'ホーム & 掲示板', 'お知らせ・地域の話題がひと目で', U);
  const tStart = mark();
  await sleep(2300);
  // 投票
  await click(page, page.locator('main button:has-text("投票"), button:has-text("投票")'));
  await feature(page, '投票', 'コミュニティの意思決定に1タップで参加', U);
  await sleep(1400);
  await click(page, page.getByText('中央公園'));
  await sleep(600);
  await click(page, page.locator('button:has-text("投票する")'));
  await sleep(1300);
  // チェックイン
  await click(page, page.locator('a[href="/checkin"]'));
  await feature(page, 'QRチェックイン', 'お店・施設で読み取ってポイント獲得', U);
  await sleep(2100);
  // タスク
  await click(page, page.locator('a[href="/tasks"]'));
  await feature(page, 'タスク受発注', '地域の仕事に応募して DAO トークンを獲得', U);
  await sleep(800);
  await click(page, page.locator('main').getByText('商店街の朝清掃'));
  await sleep(2400);
  // ウォレット
  await click(page, page.locator('a[href="/wallet"]'));
  await feature(page, 'ウォレット', '残高・取引履歴・クーポン交換', U);
  await sleep(2500);
  // DM
  await click(page, page.locator('a[href="/dm"]'));
  await feature(page, 'DM & トークン送付', 'お礼のトークンをメッセージに添えて', U);
  await sleep(2500);
  // メンバー
  await click(page, page.locator('a[href="/members"]'));
  await feature(page, 'メンバー', '保有トークン・XPで活動が見える', U);
  await sleep(2200);
  return { tStart, tEnd: mark() };
});

// ---- D: admin section card (1.2s)
await segment('sec-admin', async (page, mark) => {
  await page.setContent(sectionCard('02', '管理者画面', '運営チームのためのダッシュボード', 'linear-gradient(135deg,#2a1224,#9d2960)'));
  const tStart = mark(); await sleep(1200); return { tStart, tEnd: mark() };
});

// ---- E: admin tour
await segment('admin', async (page, mark) => {
  await page.goto(BASE + '/admin', { waitUntil: 'load', timeout: 60000 });
  await sleep(1100);
  await badge(page, '管理者画面', A);
  await feature(page, 'ダッシュボード', 'KPI・要対応アラートをひと目で', A);
  const tStart = mark();
  await sleep(2000);
  await page.mouse.move(660, 400, { steps: 10 });
  await page.mouse.wheel(0, 520);
  await sleep(1400);
  // 掲示板モデレーション
  await click(page, page.locator('a[href="/admin/board"]'));
  await feature(page, 'モデレーション', '通報された投稿にワンクリックで対応', A);
  await sleep(2500);
  // メンバー管理
  await click(page, page.locator('a[href="/admin/members"]'));
  await feature(page, 'メンバー管理', 'ロール変更・一斉DM・CSV書き出し', A);
  await sleep(2500);
  // 店舗・チェックイン
  await click(page, page.locator('a[href="/admin/shops"]'));
  await feature(page, '店舗QR発行', '店舗を登録してチェックインQRを発行', A);
  await sleep(2600);
  // 投票管理
  await click(page, page.locator('a[href="/admin/votes"]'));
  await feature(page, '投票管理', '作成・締め切り・結果をリアルタイムに', A);
  await sleep(2300);
  // トークン・ランク
  await click(page, page.locator('a[href="/admin/tokens"]'));
  await feature(page, 'トークン経済', '報酬テーブルと発行をコントロール', A);
  await sleep(2200);
  await click(page, page.locator('button:has-text("ランク条件")'));
  await feature(page, 'ランク設計', '貢献に応じたランク条件を自由に設定', A);
  await sleep(2000);
  return { tStart, tEnd: mark() };
});

// ---- F: end card (3.0s)
await segment('end', async (page, mark) => {
  await page.setContent(`<!doctype html><html><body style="margin:0"><div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;background:radial-gradient(1200px 700px at 50% 35%, #262450 0%, #101018 70%);font-family:sans-serif">
  <style>@keyframes up{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:none}}.a{animation:up .5s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.12s}.d2{animation-delay:.28s}.d3{animation-delay:.45s}</style>
  <div class="a" style="display:flex;align-items:center;gap:18px">
    <div style="width:72px;height:72px;border-radius:18px;background:#fff;color:#101018;display:flex;align-items:center;justify-content:center;font-size:44px;font-weight:800">D</div>
    <div style="color:#fff;font-size:54px;font-weight:800;letter-spacing:.04em">DAOX</div>
  </div>
  <div class="a d1" style="color:#c9c8ee;font-size:25px;font-weight:700">地域コミュニティを、DAOでなめらかに。</div>
  <div class="a d2" style="display:flex;gap:10px;margin-top:6px">
    ${['QRチェックイン','タスク受発注','投票','トークン','モデレーション'].map(t=>`<span style="background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);color:#dedbf5;padding:7px 16px;border-radius:999px;font-size:14.5px;font-weight:600">${t}</span>`).join('')}
  </div>
  <div class="a d3" style="color:#8886b8;font-size:16px;margin-top:8px">デモ: daox-zeta.vercel.app/home ｜ 管理者: /admin</div>
  </div></body></html>`);
  const tStart = mark(); await sleep(3000); return { tStart, tEnd: mark() };
});

await ctx.close();
await browser.close();
fs.writeFileSync('video60/meta.json', JSON.stringify(meta, null, 2));
const total = meta.reduce((s, m) => s + m.trimEnd - m.trimStart, 0);
console.log('DONE, kept total', total.toFixed(2), 's');
