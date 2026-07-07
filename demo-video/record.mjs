import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import { installFetchRoute } from './fetchproxy.mjs';
import fs from 'fs';

const BASE = 'https://daox-zeta.vercel.app';
const VDIR = 'video/segments';
fs.rmSync(VDIR, { recursive: true, force: true });

const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const ctx = await browser.newContext({
  viewport: { width: 1280, height: 720 },
  recordVideo: { dir: VDIR, size: { width: 1280, height: 720 } },
});
await installFetchRoute(ctx);

// visible cursor overlay
await ctx.addInitScript(() => {
  const mk = () => {
    if (document.getElementById('__pwcursor')) return;
    const d = document.createElement('div');
    d.id = '__pwcursor';
    d.style.cssText = 'position:fixed;z-index:2147483647;width:26px;height:26px;border-radius:50%;background:rgba(85,80,230,.30);border:2.5px solid rgba(85,80,230,.95);box-shadow:0 2px 10px rgba(85,80,230,.5);pointer-events:none;transform:translate(-50%,-50%);left:-60px;top:-60px;transition:width .12s,height .12s';
    document.documentElement.appendChild(d);
    addEventListener('mousemove', e => { d.style.left = e.clientX + 'px'; d.style.top = e.clientY + 'px'; }, true);
    addEventListener('mousedown', () => { d.style.width = '16px'; d.style.height = '16px'; }, true);
    addEventListener('mouseup', () => { d.style.width = '26px'; d.style.height = '26px'; }, true);
  };
  document.readyState === 'loading' ? addEventListener('DOMContentLoaded', mk) : mk();
});

const sleep = ms => new Promise(r => setTimeout(r, ms));
const meta = [];

async function caption(page, chip, text, chipColor = '#5550e6') {
  await page.evaluate(([chip, text, chipColor]) => {
    let c = document.getElementById('__pwcap');
    if (!c) {
      c = document.createElement('div');
      c.id = '__pwcap';
      c.style.cssText = 'position:fixed;bottom:26px;left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:10px;background:rgba(16,16,22,.88);backdrop-filter:blur(6px);color:#fff;padding:9px 20px 9px 9px;border-radius:999px;font-family:sans-serif;font-size:15.5px;font-weight:600;letter-spacing:.02em;z-index:2147483600;box-shadow:0 6px 24px rgba(0,0,0,.35);white-space:nowrap;transition:opacity .35s;opacity:0';
      document.documentElement.appendChild(c);
    }
    c.innerHTML = `<span style="background:${chipColor};padding:4px 12px;border-radius:999px;font-size:13px">${chip}</span><span>${text}</span>`;
    c.style.opacity = '0';
    requestAnimationFrame(() => requestAnimationFrame(() => { c.style.opacity = '1'; }));
  }, [chip, text, chipColor]);
}

async function moveAndClick(page, locator) {
  const el = locator.first();
  try {
    const box = await el.boundingBox({ timeout: 4000 });
    if (box) {
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2, { steps: 25 });
      await sleep(250);
      await el.click({ timeout: 4000 });
      return true;
    }
  } catch (e) { console.log('click skip:', e.message.split('\n')[0]); }
  return false;
}

function card(title, sub, small) {
  return `<!doctype html><html><body style="margin:0"><div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:26px;background:radial-gradient(1200px 700px at 50% 35%, #262450 0%, #101018 70%);font-family:sans-serif">
  <style>@keyframes up{from{opacity:0;transform:translateY(26px)}to{opacity:1;transform:none}}
  .a{animation:up .7s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.15s}.d2{animation-delay:.35s}</style>
  <div class="a" style="display:flex;align-items:center;gap:22px">
    <div style="width:84px;height:84px;border-radius:22px;background:#fff;color:#101018;display:flex;align-items:center;justify-content:center;font-size:52px;font-weight:800">D</div>
    <div style="color:#fff;font-size:64px;font-weight:800;letter-spacing:.04em">${title}</div>
  </div>
  <div class="a d1" style="color:#c9c8ee;font-size:27px;font-weight:600">${sub}</div>
  ${small ? `<div class="a d2" style="color:#8886b8;font-size:18px">${small}</div>` : ''}
  </div></body></html>`;
}

function transitionCard(title, sub) {
  return `<!doctype html><html><body style="margin:0"><div style="width:100vw;height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:20px;background:linear-gradient(135deg,#17172a,#2b2960);font-family:sans-serif">
  <style>@keyframes up{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:none}}.a{animation:up .55s cubic-bezier(.2,.7,.2,1) both}.d1{animation-delay:.18s}</style>
  <div class="a" style="color:#fff;font-size:46px;font-weight:800">${title}</div>
  <div class="a d1" style="color:#b9b7e8;font-size:22px;font-weight:600">${sub}</div>
  </div></body></html>`;
}

async function segment(name, fn) {
  const page = await ctx.newPage();
  const t0 = Date.now();
  let tStart = 0, tEnd = 0;
  const mark = () => Date.now() - t0;
  try {
    ({ tStart, tEnd } = await fn(page, mark));
  } finally {
    const video = page.video();
    await page.close();
    const p = await video.path();
    meta.push({ name, path: p, trimStart: tStart / 1000, trimEnd: tEnd / 1000 });
    console.log(`segment ${name}: file=${p} keep ${tStart}ms..${tEnd}ms`);
  }
}

// ---- Segment A: title card (~2.7s)
await segment('title', async (page, mark) => {
  await page.setContent(card('DAOX', '地域コミュニティDAOプラットフォーム', '商店街・地域の活動をトークンでなめらかに'));
  const tStart = mark();
  await sleep(2700);
  return { tStart, tEnd: mark() };
});

// ---- Segment B: user tour (~13.5s)
await segment('user', async (page, mark) => {
  await page.goto(BASE + '/home', { waitUntil: 'load', timeout: 60000 });
  await sleep(1200); // settle fonts/render
  await caption(page, 'ユーザー画面', 'ホーム — お知らせ・掲示板・投票をひと目で');
  const tStart = mark();
  await sleep(1700);
  // switch to 投票 tab
  await moveAndClick(page, page.locator('main button:has-text("投票"), button:has-text("投票")').locator('visible=true'));
  await sleep(1700);
  // checkin
  await moveAndClick(page, page.locator('a[href="/checkin"]').locator('visible=true'));
  await caption(page, 'ユーザー画面', 'QRチェックインで来店ポイントを獲得');
  await sleep(2100);
  // tasks
  await moveAndClick(page, page.locator('a[href="/tasks"]').locator('visible=true'));
  await caption(page, 'ユーザー画面', '地域のタスクに参加して報酬を受け取る');
  await sleep(700);
  await moveAndClick(page, page.locator('main').getByText('商店街の朝清掃').locator('visible=true'));
  await sleep(2100);
  // wallet
  await moveAndClick(page, page.locator('a[href="/wallet"]').locator('visible=true'));
  await caption(page, 'ユーザー画面', 'ウォレット — 残高・取引履歴・クーポン交換');
  await sleep(2200);
  await page.mouse.wheel(0, 300);
  await sleep(1200);
  return { tStart, tEnd: mark() };
});

// ---- Segment C: transition (~1.6s)
await segment('transition', async (page, mark) => {
  await page.setContent(transitionCard('管理者画面', 'コミュニティ運営のためのダッシュボード'));
  const tStart = mark();
  await sleep(1600);
  return { tStart, tEnd: mark() };
});

// ---- Segment D: admin tour (~9.5s)
await segment('admin', async (page, mark) => {
  await page.goto(BASE + '/admin', { waitUntil: 'load', timeout: 60000 });
  await sleep(1200);
  await caption(page, '管理者画面', 'ダッシュボード — コミュニティの動きを一望', '#e6558a');
  const tStart = mark();
  await sleep(1700);
  await page.mouse.move(700, 400, { steps: 15 });
  await page.mouse.wheel(0, 500);
  await sleep(1500);
  // board moderation
  await moveAndClick(page, page.locator('a[href="/admin/board"]').locator('visible=true'));
  await caption(page, '管理者画面', '通報・承認待ちもワンクリックで対応', '#e6558a');
  await sleep(2300);
  // tokens
  await moveAndClick(page, page.locator('a[href="/admin/tokens"]').locator('visible=true'));
  await caption(page, '管理者画面', 'トークン経済の設計・発行・ランク管理', '#e6558a');
  await sleep(2400);
  return { tStart, tEnd: mark() };
});

// ---- Segment E: end card (~2.9s)
await segment('end', async (page, mark) => {
  await page.setContent(card('DAOX', '地域コミュニティを、DAOでなめらかに。', 'デモ: daox-zeta.vercel.app/home ｜ 管理者: /admin'));
  const tStart = mark();
  await sleep(2900);
  return { tStart, tEnd: mark() };
});

await ctx.close();
await browser.close();
fs.writeFileSync('video/meta.json', JSON.stringify(meta, null, 2));
console.log('DONE');
