// 债立方 · 经纪商行情（成交行情下框）每日提取
// 路由：quote-web/#/bond/broker-market
// 输出：3 Sheet Excel（Sheet1=下框源数据；Sheet2=当天汇总；Sheet3=历史累积）+ 每日 json
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const USERNAME = process.env.INNODEALING_USERNAME || 'wangle6';
const PASSWORD = process.env.INNODEALING_PASSWORD || '123456';
const LOGIN_URL = 'https://web.innodealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innodealing.com/quote-web/#/bond/broker-market';

// 阈值：上午(<=12:00) >=2.4，下午 >=2.8
const hourBJ = new Date(Date.now() + 8 * 3600 * 1000).getUTCHours();
const MIN_VOLUME = hourBJ < 12 ? 2.4 : 2.8;

const HEADLESS = process.env.HEADLESS === 'true' || process.env.CI === 'true';
const BROWSER_CHANNEL = process.env.BROWSER_CHANNEL != null
  ? process.env.BROWSER_CHANNEL
  : (HEADLESS ? '' : 'msedge');

const WORKSPACE = __dirname;
const DOWNLOAD_DIR = path.join(WORKSPACE, 'downloads');
const BROKER_DIR = process.env.BROKER_DIR || path.join(WORKSPACE, 'data', 'broker');
const SCREENSHOT_DIR = path.join(BROKER_DIR, 'screenshots');
const HISTORY_DIR = path.join(BROKER_DIR, 'daily');
[ DOWNLOAD_DIR, SCREENSHOT_DIR, BROKER_DIR, HISTORY_DIR ].forEach(d => fs.mkdirSync(d, { recursive: true }));

const BJ_DATE = new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10);

function ts() { return new Date().toLocaleTimeString('zh-CN', { hour12: false }); }
function log(...args) { console.log(`[${ts()}]`, ...args); }
async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
async function screenshot(page, name) {
  const p = path.join(SCREENSHOT_DIR, `${name}.png`);
  try {
    await page.screenshot({ path: p, fullPage: false, timeout: 15000, animations: 'disabled' });
    log(`截图: ${name}`);
  } catch (e) {
    log(`截图跳过(${name}): ${e.message.split('\n')[0]}`);
  }
}

async function main() {
  log('=== 启动浏览器 ===');
  log(`阈值(北京时间): ${hourBJ}点 → 提取 >= ${MIN_VOLUME}`);
  const userDataDir = path.join(WORKSPACE, '.chrome-data-broker-' + Date.now());
  fs.mkdirSync(userDataDir, { recursive: true });
  const launchOpts = {
    headless: HEADLESS,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    viewport: { width: 3800, height: 1080 },
    acceptDownloads: true
  };
  if (BROWSER_CHANNEL) launchOpts.channel = BROWSER_CHANNEL;
  const context = await chromium.launchPersistentContext(userDataDir, launchOpts);
  const page = await context.newPage();
  try { await page.setViewportSize({ width: 3800, height: 1080 }); await sleep(800); } catch (e) {}

  let targetFrame = null;

  async function closeAllModals(maxRounds = 5) {
    if (!targetFrame) return;
    for (let i = 0; i < maxRounds; i++) {
      const info = await targetFrame.evaluate(() => {
        const result = { clicked: 0, details: [] };
        const modalContainers = document.querySelectorAll(
          '.ant-modal, .dmuiv4-modal, [class*="ant-modal"]:not([class*="tab"]), [class*="dmuiv4-modal"], [class*="Modal"], [class*="adModal"], [class*="AdModal"], [class*="banner-modal"], [class*="popup"], [class*="dialog"]'
        );
        const visibleModals = [];
        for (const m of modalContainers) {
          const r = m.getBoundingClientRect();
          if (r.width > 200 && r.height > 100 && r.x >= 0 && r.y >= 0) visibleModals.push(m);
        }
        for (const m of visibleModals) {
          const allBtns = m.querySelectorAll('button, [role="button"], span[aria-label]');
          for (const btn of allBtns) {
            const r = btn.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) continue;
            const mRect = m.getBoundingClientRect();
            const isTopRight = r.x > mRect.x + mRect.width * 0.6 && r.y < mRect.y + 80;
            const text = (btn.textContent || btn.getAttribute('aria-label') || '').toLowerCase();
            const aria = (btn.getAttribute('aria-label') || '').toLowerCase();
            const cls = (btn.className || '').toLowerCase();
            const isClose = text.includes('close') || text.includes('关闭') || text === '×' || text === '✕' ||
              aria.includes('close') || aria.includes('关闭') || cls.includes('close');
            if ((isTopRight && r.width < 60 && r.height < 60) || isClose) {
              btn.click(); result.clicked++; result.details.push(`topRight ${Math.round(r.x)},${Math.round(r.y)}`);
            }
          }
        }
        return result;
      });
      if (info.clicked > 0) log(`  关弹窗第${i+1}轮: ${info.clicked}个`);
      if (info.clicked === 0) break;
      await sleep(500);
    }
    try { await targetFrame.keyboard.press('Escape'); } catch (e) {}
  }

  try {
    // ===== 1. 登录 =====
    log('1. 登录');
    await page.goto(LOGIN_URL, { waitUntil: 'networkidle', timeout: 30000 });
    await sleep(2000);
    await screenshot(page, '01-login-page');

    if (!page.url().includes('signin')) {
      log('已有会话，跳过登录');
    } else {
      await page.fill('input[placeholder*="手机号"]', USERNAME);
      await page.fill('input[type="password"]', PASSWORD);
      const termsLabel = page.locator('label:has-text("我已阅读并同意相关服务条款和政策")');
      if (await termsLabel.count() > 0) { await termsLabel.click(); log('已勾选服务条款'); }
      await sleep(500);
      await page.click('button:has-text("登录")');
      let logged = false;
      for (let i = 0; i < 45; i++) { await sleep(2000); if (!page.url().includes('signin')) { logged = true; break; } }
      if (!logged) throw new Error('登录失败：90s内未离开signin页');
      log('登录成功:', page.url());
      await sleep(2000);
    }
    await screenshot(page, '02-after-login');

    // ===== 2. 导航到经纪商行情 =====
    log('2. 导航到经纪商行情');
    await page.goto('https://web.innodealing.com/quote-web/', { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(5000);

    const allFrames = page.frames();
    log(`  共${allFrames.length}个frame`);
    targetFrame = allFrames.find(f => f.url().includes('quote-web') && !f.url().includes('auth-service'));
    if (!targetFrame) throw new Error('未找到quote-web iframe');
    log(`  iframe: ${targetFrame.url()}`);

    await targetFrame.evaluate(() => {
      window.location.hash = '#/bond/broker-market';
      window.dispatchEvent(new HashChangeEvent('hashchange'));
    });
    await sleep(8000);
    log(`  当前hash: ${await targetFrame.evaluate(() => window.location.hash)}`);
    await screenshot(page, '03-broker-market');

    // 关推广弹窗
    await closeAllModals(5);
    await screenshot(page, '04-after-close-modal');

    // ===== 3. 下框提取 =====
    log('3. 提取下框（成交行情）');

    // 先找到"成交行情"tab并点击（确保它在激活状态）
    await targetFrame.evaluate(() => {
      const tabs = Array.from(document.querySelectorAll('*')).filter(el => {
        const t = (el.textContent || '').trim();
        return t === '成交行情' && el.getBoundingClientRect().width > 0;
      });
      // 点最深的叶子
      tabs.sort((a, b) => {
        const da = (() => { let d = 0, p = a; while (p) { d++; p = p.parentElement; } return d; })();
        const db = (() => { let d = 0, p = b; while (p) { d++; p = p.parentElement; } return d; })();
        return db - da;
      });
      if (tabs.length) tabs[0].click();
    });
    await sleep(2000);

    // 找下框表头 y（含"平均成交"的最大 y）
    const lowerMeta = await targetFrame.evaluate(() => {
      let hy = 0;
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0 && (el.textContent || '').trim().includes('平均成交')) {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && r.y > hy) hy = r.y;
        }
      });
      return { headerY: hy };
    });
    log(`  下框表头y=${lowerMeta.headerY}`);

    // 切换到"成交行情"Tab（默认可能停留在"成交明细"，那是全市场列表，不是关注集）
    async function clickChengjiaoTab() {
      const info = await targetFrame.evaluate(() => {
        const els = Array.from(document.querySelectorAll('*')).filter(el => {
          const t = (el.textContent || '').trim();
          return t === '成交行情' && el.children.length === 0;
        });
        if (!els.length) return null;
        els.sort((a, b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
        const el = els[0];
        let p = el.parentElement, active = false;
        while (p) { if (p.className && p.className.toString && /active/i.test(p.className.toString())) { active = true; break; } p = p.parentElement; }
        try { el.click(); } catch (e) {}
        return { active };
      });
      if (!info) { log('  未找到"成交行情"Tab'); return; }
      log(`  点击"成交行情"Tab 当前active=${info.active}`);
      await sleep(2000);
    }
    await clickChengjiaoTab();
    await screenshot(page, '04b-after-chengjiao-tab');

    // 双击"平均成交"表头实现降序（两次单击）
    async function sortByAvgDesc() {
      const info = await targetFrame.evaluate(() => {
        const els = Array.from(document.querySelectorAll('*')).filter(el => {
          const t = (el.textContent || '').trim();
          return t.includes('平均成交') && el.children.length === 0 && el.getBoundingClientRect().width > 0;
        });
        if (!els.length) return null;
        els.sort((a, b) => b.getBoundingClientRect().y - a.getBoundingClientRect().y);
        const r = els[0].getBoundingClientRect();
        return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
      });
      if (!info) { log('  未找到平均成交表头'); return; }
      log(`  点平均成交 (${Math.round(info.x)},${Math.round(info.y)})`);
      await page.mouse.click(info.x, info.y);
      await sleep(1500);
      await page.mouse.click(info.x, info.y);
      await sleep(3000);
    }
    await sortByAvgDesc();
    await screenshot(page, '05-after-sort');

    // [RAW 调试] 排序后 dump 下框顶部区域所有 leaf cell（定位特定债为何未归行）
    try {
      const raw = await targetFrame.evaluate((headerY) => {
        let hdr = null;
        document.querySelectorAll('*').forEach(el => { if (el.children.length === 0 && (el.textContent || '').trim().includes('平均成交')) { const r = el.getBoundingClientRect(); if (r.width > 0 && Math.abs(r.y - headerY) <= 6) hdr = el; } });
        let box = null, p = hdr; while (p) { if (p.className && p.className.toString && /dmui-vt-background/.test(p.className.toString())) { box = p; break; } p = p.parentElement; }
        const out = { headerY, boxFound: !!box };
        if (box) {
          const br = box.getBoundingClientRect();
          out.boxRect = { y: Math.round(br.y), h: Math.round(br.h || br.height) };
          out.cells = [];
          box.querySelectorAll('*').forEach(el => {
            if (el.children.length === 0) {
              const t = (el.textContent || '').trim();
              const r = el.getBoundingClientRect();
              if (t && r.y > headerY && r.y < headerY + 280) out.cells.push({ x: Math.round(r.x), y: Math.round(r.y), t: t.slice(0, 40) });
            }
          });
          out.cells.sort((a, b) => a.y - b.y || a.x - b.x);
        }
        return out;
      }, lowerMeta.headerY);
      fs.writeFileSync(path.join(HISTORY_DIR, `_rawcells_debug_${BJ_DATE}.json`), JSON.stringify(raw, null, 1), 'utf8');
      log('  [RAW调试] 已写 _rawcells_debug_' + BJ_DATE + '.json, top cells=' + (raw.cells ? raw.cells.length : 0));
    } catch (e) { log('  [RAW调试] 失败: ' + e.message); }

    // [DOM 调试] 定位下框真实结构（是否 AG-Grid、滚动容器、成交行情 Tab）
    try {
      const domInfo = await targetFrame.evaluate(() => {
        const out = { agGridCount: document.querySelectorAll('[class*="ag-"]').length, agRows: document.querySelectorAll('.ag-row').length };
        out.tabs = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0) {
            const t = (el.textContent || '').trim();
            if (t === '成交行情' || t === '经纪商报价' || t === '报价' || t === '成交明细') {
              const r = el.getBoundingClientRect();
              out.tabs.push({ text: t, x: Math.round(r.x), y: Math.round(r.y), cls: (el.className && el.className.toString ? el.className.toString().slice(0, 60) : '') });
            }
          }
        });
        const els = Array.from(document.querySelectorAll('*')).filter(el => {
          const t = (el.textContent || '').trim();
          return t.includes('平均成交') && el.children.length === 0 && el.getBoundingClientRect().width > 0;
        });
        els.sort((a, b) => b.getBoundingClientRect().y - a.getBoundingClientRect().y);
        if (els[0]) {
          const r = els[0].getBoundingClientRect();
          const chain = []; let p = els[0];
          for (let k = 0; k < 5 && p; k++) { chain.push({ tag: p.tagName, cls: (p.className && p.className.toString ? p.className.toString().slice(0, 70) : ''), y: Math.round(p.getBoundingClientRect().y), h: Math.round(p.getBoundingClientRect().height) }); p = p.parentElement; }
          out.avgHeader = { x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height), chain };
        }
        let cont = null; document.querySelectorAll('*').forEach(el => { if (el.__lowerBox) cont = el; });
        if (cont) { const r = cont.getBoundingClientRect(); out.scrollBox = { tag: cont.tagName, cls: (cont.className && cont.className.toString ? cont.className.toString().slice(0, 70) : ''), sh: cont.scrollHeight, ch: cont.clientHeight, y: Math.round(r.y), h: Math.round(r.height) }; }
        return out;
      });
      fs.writeFileSync(path.join(HISTORY_DIR, `_dom_debug_${BJ_DATE}.json`), JSON.stringify(domInfo, null, 1), 'utf8');
      log('  [DOM调试] 已写 _dom_debug_' + BJ_DATE + '.json: ' + JSON.stringify(domInfo).slice(0, 400));
    } catch (e) { log('  [DOM调试] 失败: ' + e.message); }

    // 滚动下框：找滚动容器，按 px 步进滚动多次，失败则用鼠标滚轮兜底
    let lowerScrollContainer = null;
    async function findLowerScrollContainer() {
      return await targetFrame.evaluate((headerY) => {
        // 先找之前标记过的容器
        let marked = null;
        document.querySelectorAll('*').forEach(el => { if (el.__lowerBox) marked = el; });
        if (marked) return { found: true, top: Math.round(marked.getBoundingClientRect().top), sh: marked.scrollHeight, ch: marked.clientHeight };
        // 找"平均成交"表头，沿父链找滚动容器
        let headerEl = null;
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0 && (el.textContent || '').trim().includes('平均成交')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && Math.abs(r.y - headerY) <= 6) headerEl = el;
          }
        });
        if (!headerEl) return { found: false };
        let el = headerEl.parentElement;
        while (el && el.parentElement) {
          if (el.scrollHeight > el.clientHeight + 60 && el.clientHeight > 60) { el.__lowerBox = true; return { found: true, top: Math.round(el.getBoundingClientRect().top), sh: el.scrollHeight, ch: el.clientHeight }; }
          el = el.parentElement;
        }
        return { found: false };
      }, lowerMeta.headerY);
    }
    const scInfo = await findLowerScrollContainer();
    log(`  下框滚动容器: ${scInfo.found ? `top=${scInfo.top} sh=${scInfo.sh} ch=${scInfo.ch}` : '未找到'}`);

    async function scrollLowerDown(px) {
      // 仅用原生 scrollTop 推进虚拟滚动（不再叠加鼠标滚轮，避免滚错容器）
      return await targetFrame.evaluate((px) => {
        let cont = null;
        document.querySelectorAll('*').forEach(el => { if (el.__lowerBox) cont = el; });
        if (!cont) return false;
        const old = cont.scrollTop; cont.scrollTop += px;
        return cont.scrollTop !== old;
      }, px);
    }
    async function getLowerScrollInfo() {
      return await targetFrame.evaluate(() => {
        let cont = null;
        document.querySelectorAll('*').forEach(el => { if (el.__lowerBox) cont = el; });
        if (!cont) return null;
        return { top: cont.scrollTop, sh: cont.scrollHeight, ch: cont.clientHeight,
                 rowApprox: cont.scrollHeight > 0 && cont.children.length ? Math.round(cont.scrollHeight / Math.max(cont.children.length,1)) : 0 };
      });
    }

    // 单次快照：使用固定的列定义，只收集数据 cell
    async function detectLowerColumns(headerY) {
      return await targetFrame.evaluate((headerY) => {
        const heads = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length !== 0) return;
          const raw = (el.textContent || '').trim();
          if (!raw || raw.length > 30) return;
          const r = el.getBoundingClientRect();
          if (r.x < 100 || r.x > 4200) return;
          if (r.width < 5 || r.height < 5 || r.height > 50) return;
          if (Math.abs(r.y - headerY) <= 8) {
            heads.push({ x: Math.round(r.x + r.width / 2), w: Math.round(r.width), raw, name: raw.replace(/[↓↑↕\s]/g, '') });
          }
        });
        heads.sort((a, b) => a.x - a.w / 2 - (b.x - b.w / 2));
        const uniqueHeads = [];
        const seen = new Set();
        for (const h of heads) {
          const key = `${h.name}@${Math.round(h.x / 30)}`;
          if (seen.has(key)) continue;
          seen.add(key);
          uniqueHeads.push(h);
        }
        uniqueHeads.sort((a, b) => a.x - b.x);
        const cols = uniqueHeads.map(h => ({ x: h.x, name: h.name, raw: h.raw }));
        // 为可能无表头的最左列补"类型"
        const dcells = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length !== 0) return;
          const raw = (el.textContent || '').trim();
          if (!raw || raw.length > 60) return;
          if (/^(成交行情|成交明细|经纪商行情|今日最优|最优|全部经纪商|平安|国际|中诚|上田|国利|信唐|设置|批量导入)$/.test(raw)) return;
          if (raw === '权' || raw === '免') return;
          if (raw.startsWith('(支持导入最多') && raw.endsWith('个债券)')) return;
          const r = el.getBoundingClientRect();
          if (r.x < 100 || r.x > 4200) return;
          if (r.width < 8 || r.height < 5 || r.height > 120) return;
          const yMax = (() => {
            let h = null;
            document.querySelectorAll('*').forEach(el => { if (el.children.length === 0 && (el.textContent || '').trim().includes('平均成交')) { const r = el.getBoundingClientRect(); if (r.width > 0 && Math.abs(r.y - headerY) <= 6) h = el; } });
            if (!h) return headerY + 900;
            let p = h; while (p) { if (p.className && p.className.toString && /dmui-vt-background/.test(p.className.toString())) return p.getBoundingClientRect().bottom; p = p.parentElement; }
            return headerY + 900;
          })();
          if (r.y > headerY + 8 && r.y < yMax) {
            dcells.push({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y), w: Math.round(r.width), text: raw });
          }
        });
        if (cols.length && dcells.some(c => c.x < cols[0].x - 40)) {
          const minX = dcells.filter(c => c.x < cols[0].x - 40).reduce((m, c) => Math.min(m, c.x), Infinity);
          cols.unshift({ x: minX, name: '类型', raw: '类型' });
        }
        return { cols };
      }, headerY);
    }

    async function snapshotLower(headerY, cols) {
      return await targetFrame.evaluate(({ headerY, cols }) => {
        // 定位下框真实容器（和 rawcells dump 同源，避免 document 范围过大引入干扰）
        let hdr = null;
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0 && (el.textContent || '').trim().includes('平均成交')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && Math.abs(r.y - headerY) <= 6) hdr = el;
          }
        });
        let box = null, p = hdr;
        while (p) {
          if (p.className && p.className.toString && /dmui-vt-background/.test(p.className.toString())) { box = p; break; }
          p = p.parentElement;
        }
        if (!box) return { rows: [], names: [], diag: { error: 'lower box not found' } };
        const yMax = box.getBoundingClientRect().bottom;

        const dcells = [];
        function shouldSkipText(t) {
          if (!t || t.length > 60) return true;
          if (/^(成交行情|成交明细|经纪商行情|今日最优|最优|全部经纪商|平安|国际|中诚|上田|国利|信唐|设置|批量导入)$/.test(t)) return true;
          if (t === '权' || t === '免') return true;
          if (t.startsWith('(支持导入最多') && t.endsWith('个债券)')) return true;
          return false;
        }
        function addCell(text, r) {
          if (shouldSkipText(text)) return;
          if (r.x < 100 || r.x > 4200) return;
          if (r.width < 8 || r.height < 5 || r.height > 200) return;
          if (r.y > headerY + 8 && r.y < yMax) {
            dcells.push({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y), w: Math.round(r.width), text });
          }
        }
        box.querySelectorAll('*').forEach(el => {
          const r = el.getBoundingClientRect();
          if (r.width < 8 || r.height < 5 || r.height > 200) return;
          if (r.x < 100 || r.x > 4200) return;
          if (el.children.length === 0) {
            addCell((el.textContent || '').trim(), r);
            return;
          }
          // 非 leaf：提取直接文本节点（处理文本与子元素/图标并列导致文本不在 leaf 中的情况）
          let directText = '';
          for (const node of el.childNodes) {
            if (node.nodeType === Node.TEXT_NODE) directText += node.textContent;
          }
          directText = directText.trim();
          if (directText && !shouldSkipText(directText)) {
            let covered = false;
            for (const child of el.querySelectorAll('*')) {
              if (child.children.length === 0 && child.textContent.trim() === directText) { covered = true; break; }
            }
            if (!covered) addCell(directText, r);
          }
        });

        dcells.sort((a, b) => a.y - b.y || a.x - b.x);
        const rawRows = [];
        let cur = null, curMaxY = -Infinity;
        for (const c of dcells) {
          if (!cur || (c.y - curMaxY) > 12) { cur = []; rawRows.push(cur); }
          cur.push(c);
          if (c.y > curMaxY) curMaxY = c.y;
        }

        // 区间归属：表头中心为 seed，区间取前后两列 seed 的中点（兼容左对齐短文本列 + 右对齐数值列）
        cols.sort((a, b) => a.x - b.x);
        const colIntervals = cols.map((c, i) => {
          const prevSpacing = i > 0 ? (c.x - cols[i - 1].x) : (c.x - 100);
          const nextSpacing = (i + 1 < cols.length) ? (cols[i + 1].x - c.x) : 150;
          return { left: c.x - prevSpacing / 2, right: c.x + nextSpacing / 2, seed: c.x };
        });
        const nameColIdx = cols.findIndex(c => c.name.includes('债券简称'));
        const avgColIdx = cols.findIndex(c => c.name.includes('平均成交'));
        const codeColIdx = cols.findIndex(c => c.name.includes('债券代码'));

        const outRows = [];
        for (const row of rawRows) {
          const vals = new Array(cols.length).fill('');
          for (const c of row) {
            let best = -1, bestD = Infinity;
            for (let i = 0; i < cols.length; i++) {
              if (c.x >= colIntervals[i].left && c.x < colIntervals[i].right) {
                const d = Math.abs(c.x - cols[i].x);
                if (d < bestD) { bestD = d; best = i; }
              }
            }
            if (best >= 0) {
              const existing = vals[best];
              const t = c.text;
              if (best === nameColIdx) {
                const isBetter = !existing || t.length > existing.length;
                if (isBetter) vals[best] = t;
              } else if (best === codeColIdx) {
                if (!existing || t.length > existing.length) vals[best] = t;
              } else {
                if (!existing) vals[best] = t;
              }
            }
          }
          const bondName = nameColIdx >= 0 ? (vals[nameColIdx] || '').trim() : '';
          const bondCode = codeColIdx >= 0 ? (vals[codeColIdx] || '').trim() : '';
          let maxVal = 0;
          if (avgColIdx >= 0 && vals[avgColIdx]) {
            const av = parseFloat(vals[avgColIdx]);
            if (!isNaN(av)) maxVal = av;
          }
          // 兜底：平均成交列未解析到数值时，用同一行中“最新成交/开盘/最高/最低”的最大值
          if (maxVal === 0) {
            const candidateIdxs = [avgColIdx - 4, avgColIdx - 3, avgColIdx - 2, avgColIdx - 1].filter(i => i >= 0 && i < cols.length);
            for (const idx of candidateIdxs) {
              const v = parseFloat(vals[idx]);
              if (!isNaN(v) && v > maxVal) maxVal = v;
            }
          }
          if (bondName && !/^\d+$/.test(bondName) && bondName !== '休' && bondName.length >= 2) {
            outRows.push({ y: row[0].y, vals, maxVal, bondName, bondCode });
          }
        }
        const diag = { taifeng: [] };
        document.querySelectorAll('*').forEach(el => {
          const t = (el.textContent || '').trim();
          if (t.includes('泰丰')) {
            const r = el.getBoundingClientRect();
            diag.taifeng.push({ t: t.slice(0,40), x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height), leaf: el.children.length === 0 });
          }
        });
        return { rows: outRows, names: outRows.map(r => r.bondName), diag };
      }, { headerY, cols });
    }

    // 先检测列定义（只测一次）
    const colDetect = await detectLowerColumns(lowerMeta.headerY);
    const fixedCols = colDetect.cols;
    log(`  下框固定列定义(${fixedCols.length}列): ${fixedCols.map(c => c.name).join(' | ')}`);
    try {
      const domPath = path.join(HISTORY_DIR, `_dom_debug_${BJ_DATE}.json`);
      const domInfo = JSON.parse(fs.readFileSync(domPath, 'utf8'));
      domInfo.fixedCols = fixedCols.map(c => ({ name: c.name, x: c.x }));
      fs.writeFileSync(domPath, JSON.stringify(domInfo, null, 1), 'utf8');
    } catch (e) { log('  [DOM调试] 追加 fixedCols 失败: ' + e.message); }

    // 滚动 + 快照循环（自核对：累计每帧可见债券名；滚到底自动停，进度停滞才退出）
    const allRows = new Map();
    const domNameSeen = new Set();
    let iter = 0, sameCount = 0, lastTop = -1, bottomReached = false;
    while (iter < 150 && !bottomReached) {
      iter++;
      const snap = await snapshotLower(lowerMeta.headerY, fixedCols);
      let added = 0;
      for (const r of snap.rows) {
        const key = r.bondCode || r.bondName;
        if (!allRows.has(key)) { allRows.set(key, r); added++; }
        if (r.bondName) domNameSeen.add(r.bondName);
      }
      if (iter === 1) {
        log(`  [下框快照1 行明细] ${snap.rows.map(r => `${r.bondName}:${r.maxVal}`).join(' | ')}`);
        if (snap.diag && snap.diag.taifeng) log(`  [下框快照1 泰丰诊断] ${JSON.stringify(snap.diag.taifeng)}`);
      }
      log(`  [下框快照${iter}] 本帧${snap.rows.length}行，新增${added}，累计${allRows.size}，可见名${domNameSeen.size}`);
      await screenshot(page, 'broker_frame_' + String(iter).padStart(2, '0'));
      const si = await getLowerScrollInfo();
      if (si) {
        if (si.sh - si.ch - si.top <= 2) { bottomReached = true; log('  已到达下框底部'); }
        if (si.top === lastTop) sameCount++; else { sameCount = 0; lastTop = si.top; }
        if (sameCount >= 4) { log('  滚动不再前进，提前停止'); break; }
      }
      if (!bottomReached) { await scrollLowerDown(400); await sleep(500); }
    }

    // 滚动回顶部，方便后续双击进详情页（虚拟滚动后目标债券可能不在 DOM 中）
    await scrollLowerDown(-100000);
    await sleep(1200);
    await screenshot(page, '06-back-to-top');

    // ===== 自核对（模拟人眼）：每帧可见债券名 vs 捕获集合；不足则补抓 =====
    const capturedNames = new Set(Array.from(allRows.values()).map(r => r.bondName).filter(Boolean));
    const missingNames = [...domNameSeen].filter(n => !capturedNames.has(n));
    const estTotal = await targetFrame.evaluate(() => {
      let found = 0;
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length) return;
        const t = (el.textContent || '').trim();
        const m = t.match(/共\s*(\d+)\s*条/) || t.match(/(\d+)\s*条/);
        if (m) { const v = parseInt(m[1], 10); if (v > found && v < 5000) found = v; }
      });
      return found;
    });
    log(`  [自核对] 可见债券名${domNameSeen.size} / 捕获${capturedNames.size} / 页面估算总条数${estTotal || '未知'}`);
    if (missingNames.length) log(`  [自核对][警告] 可见但未捕获的债券: ${missingNames.join(' | ')}`);
    let recovered = 0;
    if (estTotal && capturedNames.size < estTotal) {
      log(`  [自核对] 捕获数(${capturedNames.size}) < 估算总数(${estTotal})，执行补抓扫描...`);
      await scrollLowerDown(-100000); await sleep(800);
      let it2 = 0, same2 = 0, last2 = -1;
      while (it2 < 150) {
        it2++;
        const snap = await snapshotLower(lowerMeta.headerY, fixedCols);
        for (const r of snap.rows) {
          const key = r.bondCode || r.bondName;
          if (!allRows.has(key)) { allRows.set(key, r); recovered++; }
          if (r.bondName) domNameSeen.add(r.bondName);
        }
        const si = await getLowerScrollInfo();
        if (!si) break;
        if (si.sh - si.ch - si.top <= 2) break;
        if (si.top === last2) { same2++; if (same2 >= 4) break; } else { same2 = 0; last2 = si.top; }
        await scrollLowerDown(400); await sleep(400);
      }
      log(`  [自核对] 补抓新增${recovered}条，累计${allRows.size}`);
      await scrollLowerDown(-100000); await sleep(1000);
    }
    // 保存核对报告（供邮件附件/人眼复核）
    try {
      const verifyReport = {
        bjDate: BJ_DATE, hourBJ, threshold: MIN_VOLUME,
        visibleNames: domNameSeen.size, capturedNames: capturedNames.size,
        estimatedTotal: estTotal || null, missingNames, recoveredExtra: recovered,
        captureComplete: estTotal ? (capturedNames.size >= estTotal) : null
      };
      fs.writeFileSync(path.join(SCREENSHOT_DIR, `broker_verify_${BJ_DATE}.json`), JSON.stringify(verifyReport, null, 2), 'utf8');
      log(`  [自核对] 报告已写 broker_verify_${BJ_DATE}.json`);
    } catch (e) { log('  [自核对] 写报告失败: ' + e.message); }

    const colDefs = fixedCols;

    const rows = Array.from(allRows.values()).map(r => ({
      cells: r.vals,
      maxVal: r.maxVal,
      bondName: r.bondName,
      bondCode: r.bondCode,
      y: r.y
    }));
    // 按平均成交降序
    rows.sort((a, b) => b.maxVal - a.maxVal);
    let filtered = rows.filter(r => r.maxVal >= MIN_VOLUME);
    log(`  下框总共${rows.length}行，>=${MIN_VOLUME} 筛选后 ${filtered.length}行`);
    // [DEBUG] 导出原始抓取行(阈值前)用于核对漏行（诊断用，体积小）
    try {
      const dbg = rows.map(r => ({ bondName: r.bondName, bondCode: r.bondCode, maxVal: r.maxVal, pass: r.maxVal >= MIN_VOLUME }));
      fs.writeFileSync(path.join(HISTORY_DIR, `_capture_debug_${BJ_DATE}.json`), JSON.stringify({ bjDate: BJ_DATE, hourBJ, threshold: MIN_VOLUME, visibleNames: domNameSeen.size, capturedNames: capturedNames.size, estimatedTotal: estTotal || null, missingNames, recoveredExtra: recovered, captureComplete: estTotal ? (capturedNames.size >= estTotal) : null, captured: dbg }, null, 1), 'utf8');
      log(`  [DEBUG] 原始抓取${rows.length}行已写入 _capture_debug_${BJ_DATE}.json`);
    } catch (e) { log('  [DEBUG] 写调试文件失败: ' + e.message); }
    for (const r of filtered.slice(0, 20)) {
      log(`    [${r.maxVal}] ${r.bondName} ${r.bondCode} | ${colDefs.map((c, i) => `${c.name}=${r.cells[i] || ''}`).slice(0, 12).join('  ')}`);
    }

    if (filtered.length === 0) {
      log('  ⚠️ 无满足阈值数据，仍继续生成空文件');
      await screenshot(page, '06-empty');
      try { await targetFrame.screenshot({ path: path.join(SCREENSHOT_DIR, '06-empty-frame.png') }); } catch(e){}
    }

    // ===== 4. 详情页提取区域 + 发行人全称 =====
    await sleep(1500);
    log('4. 详情页提取发行人全称与区域');
    const bondRegionMap = new Map();

    const lIdx = {};
    for (let i = 0; i < colDefs.length; i++) lIdx[colDefs[i].name] = i;
    function getLowerVal(r, colName) {
      const i = lIdx[colName];
      return (i != null && i >= 0 && r.cells[i]) ? r.cells[i] : '';
    }

    if (filtered.length > 0) {
      const PROV_LIST = ['北京','天津','上海','重庆','河北','山西','辽宁','吉林','黑龙江','江苏','浙江','安徽','福建','江西','山东','河南','湖北','湖南','广东','海南','四川','贵州','云南','陕西','甘肃','青海','内蒙古','广西','西藏','宁夏','新疆'];
      const iframeOffset = await page.evaluate(() => {
        const iframes = Array.from(document.querySelectorAll('iframe'));
        for (const f of iframes) if (f.src && f.src.includes('quote-web')) { const r = f.getBoundingClientRect(); return { x: r.x, y: r.y }; }
        return { x: 0, y: 0 };
      }).catch(() => ({ x: 0, y: 0 }));

      const getNavSig = () => [page.url(), ...page.frames().map(f => f.url())].join('|');
      const isDetailUrl = () => getNavSig().includes('bond/detail') || getNavSig().includes('/detail/');
      const isDetailContent = (frame) => frame.evaluate(() => /性质|主体评级|债项评级/.test(document.body?.textContent || '')).catch(() => false);
      const isTargetBond = (frame, bCode, bName) => {
        const code = String(bCode || '').split('.')[0];
        return frame.evaluate(({ code, name }) => {
          const body = document.body?.textContent || '';
          return !!(code && code.length >= 4 && body.includes(code)) || !!(name && name.length >= 3 && body.includes(name));
        }, { code, name: bName }).catch(() => false);
      };

      async function dblclickIntoDetail(frame, name, bCode, prevSig) {
        // 先通过 evaluate 找元素并滚到中间，返回 iframe 内坐标
        const targetInfo = await frame.evaluate((nm) => {
          const els = Array.from(document.querySelectorAll('*')).filter(el => {
            const t = (el.textContent || '').trim();
            return t === nm || (t.startsWith(nm) && t.length <= nm.length + 8);
          });
          if (!els.length) return null;
          els.sort((a, b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
          const target = els.find(el => { const r = el.getBoundingClientRect(); return r.width > 30 && r.width < 400 && r.height > 10 && r.height < 60; }) || els[0];
          target.scrollIntoView({ block: 'center' });
          const r = target.getBoundingClientRect();
          return { x: r.x + r.width / 2, y: r.y + r.height / 2, w: r.width, h: r.height };
        }, name).catch(() => null);
        if (!targetInfo) return false;
        // 用真实鼠标双击（Playwright 坐标为 page 坐标，需加上 iframe 偏移）
        const px = iframeOffset.x + targetInfo.x;
        const py = iframeOffset.y + targetInfo.y;
        log(`  双击进入详情: ${name} (${Math.round(px)},${Math.round(py)})`);
        await page.mouse.move(px, py);
        await sleep(120);
        await page.mouse.dblclick(px, py);
        for (let i = 0; i < 20; i++) {
          await sleep(500);
          if (getNavSig() !== prevSig && isDetailUrl()) {
            if (await isDetailContent(frame) && await isTargetBond(frame, bCode, name)) {
              await sleep(800); return true;
            }
          }
        }
        return false;
      }

      async function searchSwitch(frame, bCode, bName, prevSig) {
        const code = String(bCode || '').split('.')[0];
        if (!code) return false;
        const box = await frame.evaluate(() => {
          const input = document.querySelector('input.ant-select-search__field') || Array.from(document.querySelectorAll('input')).find(el => String(el.className || '').includes('ant-select-search'));
          if (!input) return null;
          const r = input.getBoundingClientRect();
          return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
        }).catch(() => null);
        if (!box) return false;
        try { await frame.click('input.ant-select-search__field', { timeout: 3000 }); } catch(e) { await page.mouse.click(iframeOffset.x + box.x, iframeOffset.y + box.y); }
        await sleep(250);
        await page.keyboard.press('Control+A'); await sleep(120);
        await page.keyboard.press('Backspace'); await sleep(120);
        await page.keyboard.type(code, { delay: 30 });
        await sleep(900);
        await page.keyboard.press('Enter');
        for (let i = 0; i < 24; i++) {
          await sleep(500);
          if (getNavSig() !== prevSig && isDetailUrl()) {
            if (await isDetailContent(frame) && await isTargetBond(frame, code, bName)) { await sleep(900); return true; }
          }
        }
        return false;
      }

      async function extractFrom(frame) {
        return await frame.evaluate((PROV_LIST) => {
          const result = { province: '', issuerFull: '', method: '' };
          document.querySelectorAll('*').forEach(el => {
            if (result.province) return;
            const t = (el.textContent || '').trim();
            const m = t.match(/^省(.{2,5})$/);
            if (m) {
              let p = m[1];
              if (p.endsWith('省') || p.endsWith('市')) p = p.slice(0, -1);
              if (PROV_LIST.includes(p)) { result.province = p; result.method = '省字段'; }
            }
          });
          if (!result.province) {
            const allEls = Array.from(document.querySelectorAll('*'));
            for (let i = 0; i < allEls.length; i++) {
              const el = allEls[i];
              const t = (el.textContent || '').trim();
              if (t !== '省') continue;
              for (let j = i + 1; j < Math.min(i + 6, allEls.length); j++) {
                const sib = allEls[j];
                const st = (sib.textContent || '').trim();
                if (st === '市' || st === '性质' || st.length > 10) break;
                let p = st.replace(/省$/, '');
                if (PROV_LIST.includes(p)) { result.province = p; result.method = '省邻元素'; break; }
                if (PROV_LIST.includes(st)) { result.province = st; result.method = '省邻元素'; break; }
              }
              if (result.province) break;
            }
          }
          if (!result.province) {
            document.querySelectorAll('*').forEach(el => {
              if (result.province) return;
              const t = (el.textContent || '').trim();
              if (t.includes('性质')) {
                const m = t.match(/(北京|天津|上海|重庆|河北|山西|辽宁|吉林|黑龙江|江苏|浙江|安徽|福建|江西|山东|河南|湖北|湖南|广东|海南|四川|贵州|云南|陕西|甘肃|青海|内蒙古|广西|西藏|宁夏|新疆)/);
                if (m) { result.province = m[1]; result.method = '性质字段'; }
              }
            });
          }
          const cands = Array.from(document.querySelectorAll('*')).filter(el => {
            const t = (el.textContent || '').trim();
            const r = el.getBoundingClientRect();
            return r.y >= 40 && r.y <= 260 && r.x >= 600 && t.length >= 4 && t.length <= 60 &&
              !/(成交|行情|最新|价格|Bid|Ofr|笔数|更新时间|流动性|中债|中证|主承销|发行日|剩余期限|收益率)/.test(t) &&
              /有限公司|公司|银行|政府|财政部|集团|投资|控股|开发|管理局|厅|委员会|国资|资产|有限/.test(t);
          }).sort((a, b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
          if (cands.length) {
            let name = cands[0].textContent.trim();
            const rc = name.search(/主体评级|债项评级|评级/);
            if (rc > 0) name = name.slice(0, rc).trim();
            const toks = name.split(/\s+/);
            let ci = -1;
            for (let ti = 0; ti < toks.length; ti++) {
              const tk = toks[ti];
              const m = tk.match(/^省(.{2,5})$/);
              if (m) { let p = m[1]; if (p.endsWith('省') || p.endsWith('市')) p = p.slice(0, -1); if (PROV_LIST.includes(p)) { ci = ti; break; } }
              if (tk.startsWith('性质')) { ci = ti; break; }
            }
            if (ci > 0) name = toks.slice(0, ci).join(' ').trim();
            name = name.replace(/市政市场化.*$|城投市场化.*$|市场化.*$/, '').replace(/(城投|国企|央企|民企|地方国企|中央国企|国有)$/, '').trim();
            result.issuerFull = name;
          }
          return result;
        }, PROV_LIST);
      }

      // 首只进详情页
      let onDetail = false;
      const firstName = getLowerVal(filtered[0], '债券简称');
      const firstCode = getLowerVal(filtered[0], '债券代码');
      for (let attempt = 0; attempt < 3 && !onDetail; attempt++) {
        const sigBefore = getNavSig();
        onDetail = await dblclickIntoDetail(targetFrame, firstName, firstCode, sigBefore);
        log(`  首只进入详情页(尝试${attempt+1}): ${onDetail ? '✅' : '❌'}`);
        if (attempt > 0) await sleep(800);
      }

      let prevIssuerFull = '';
      for (let ri = 0; ri < filtered.length; ri++) {
        const ur = filtered[ri];
        const bCode = getLowerVal(ur, '债券代码') || '';
        const bName = getLowerVal(ur, '债券简称') || '';
        if (!bCode) continue;

        if (ri > 0 || !onDetail) {
          const sigBefore = getNavSig();
          let ok = await searchSwitch(targetFrame, bCode, bName, sigBefore);
          if (!ok) { await sleep(1200); ok = await searchSwitch(targetFrame, bCode, bName, getNavSig()); }
          if (!ok) { log(`  [${ri + 1}] ${bName} 导航失败，跳过`); continue; }
          onDetail = true;
        }

        let ex = await extractFrom(targetFrame);
        if (ex.issuerFull && ex.issuerFull === prevIssuerFull) { await sleep(1500); ex = await extractFrom(targetFrame); }
        let retry = 0;
        while ((!ex.issuerFull || !ex.province) && retry < 2) { retry++; await sleep(1500); ex = await extractFrom(targetFrame); }
        const finalRegion = ex.province || '';
        bondRegionMap.set(bCode, {
          region: finalRegion,
          bondName: bName,
          method: finalRegion ? ('detail_' + (ex.method || 'none')) : 'detail_none',
          issuerFull: ex.issuerFull || ''
        });
        if (ex.issuerFull) prevIssuerFull = ex.issuerFull;
        log(`  [${ri + 1}/${filtered.length}] ${bName} → 省:${finalRegion || '(空)'} 全称:${ex.issuerFull || '(无)'}`);
        try { await page.screenshot({ path: path.join(SCREENSHOT_DIR, `detail-${bCode.split('.').join('_')}.png`) }); } catch(e){}
      }

      // 回填
      const missFull = [...bondRegionMap.entries()].filter(([, v]) => !v.issuerFull || !v.region);
      if (missFull.length) {
        log(`  回填: ${missFull.length}只债券`);
        for (const [bCode, v] of missFull) {
          const sigBefore = getNavSig();
          let ok = await searchSwitch(targetFrame, bCode, v.bondName, sigBefore);
          if (!ok) { await sleep(1200); ok = await searchSwitch(targetFrame, bCode, v.bondName, getNavSig()); }
          if (!ok) continue;
          await sleep(800);
          let ex = await extractFrom(targetFrame);
          let rt = 0;
          while ((!ex.issuerFull || !ex.province) && rt < 2) { rt++; await sleep(1500); ex = await extractFrom(targetFrame); }
          const merged = { ...v };
          if (ex.province) merged.region = ex.province;
          if (ex.issuerFull) merged.issuerFull = ex.issuerFull;
          bondRegionMap.set(bCode, merged);
          log(`  ↺ 回填 [${v.bondName}] 省:${ex.province || '(空)'} 全称:${ex.issuerFull || '(无)'}`);
        }
      }

      // 关闭详情标签页（回到列表）
      try {
        await targetFrame.evaluate(() => {
          window.location.hash = '#/bond/broker-market';
          window.dispatchEvent(new HashChangeEvent('hashchange'));
        });
        await sleep(3000);
      } catch(e){}
    }

    // ===== 5. 生成 Excel =====
    log('5. 生成 Excel');
    function cleanCell(text) {
      if (!text) return text;
      const s = String(text).trim();
      const tokens = s.split(/\s+/);
      const result = [tokens[0]];
      for (let i = 1; i < tokens.length; i++) if (tokens[i] !== result[result.length - 1]) result.push(tokens[i]);
      return result.join(' ');
    }

    const wb = new ExcelJS.Workbook();

    // Sheet1: 下框源数据
    const ws1 = wb.addWorksheet('成交行情_下框');
    if (filtered.length > 0) {
      ws1.addRow(colDefs.map(c => c.name));
      for (const r of filtered) ws1.addRow(r.cells.map(cleanCell));
    } else { ws1.addRow(['暂无数据']); }

    // Sheet2: 当天汇总（固定表头，发行人/区域来自详情页）
    const SHEET2_HEADERS = ['成交日期','债券简称','剩余期限','最新成交','中债/中证','中债偏离(BP)','发行人','区域','债券代码','中债隐含评级','主/债'];
    const ws2 = wb.addWorksheet('成交行情汇总');
    ws2.addRow(SHEET2_HEADERS);

    const [y, m, d] = BJ_DATE.split('-');
    const dateFormatted = `${y}-${parseInt(m)}-${parseInt(d)}`;
    const sheet2Rows = [];

    for (const r of filtered) {
      const code = getLowerVal(r, '债券代码');
      const name = getLowerVal(r, '债券简称');
      const info = bondRegionMap.get(code) || {};
      const gv = (cn) => getLowerVal(r, cn);
      const row = [
        dateFormatted,                  // 成交日期
        name,                           // 债券简称
        gv('剩余期限'),                  // 剩余期限
        gv('最新成交'),                  // 最新成交
        gv('中债/中证'),                 // 中债/中证
        gv('中债偏离(BP)'),              // 中债偏离(BP)
        info.issuerFull || '',          // 发行人（全称）
        info.region || '',              // 区域
        code,                           // 债券代码
        gv('中债隐含评级'),              // 中债隐含评级
        gv('主/债'),                     // 主/债
      ];
      ws2.addRow(row);
      sheet2Rows.push(row);
    }
    log(`  Sheet2 完成: ${sheet2Rows.length}行`);

    // Sheet3: 历史累积
    log('6. 构建 Sheet3: 历史累积');
    const ws3 = wb.addWorksheet('历史累积汇总');
    ws3.addRow(SHEET2_HEADERS);
    let sheet3Count = 0, sheet3Days = 0;
    const todayJsonPath = path.join(HISTORY_DIR, `${BJ_DATE}.json`);
    let skipEmptyWrite = false;

    if (sheet2Rows.length === 0 && fs.existsSync(todayJsonPath)) {
      try { const existing = JSON.parse(fs.readFileSync(todayJsonPath, 'utf8')); if (existing.rows?.length > 0) skipEmptyWrite = true; } catch(e){}
      if (skipEmptyWrite) log('  本次无数据但已有历史，跳过覆盖');
    }

    if (!skipEmptyWrite) {
      fs.writeFileSync(todayJsonPath, JSON.stringify({ date: BJ_DATE, dateFormatted, headers: SHEET2_HEADERS, rows: sheet2Rows }, null, 1), 'utf8');
      log(`  已写 daily json: ${todayJsonPath} (${sheet2Rows.length}行)`);
    }

    const dayFiles = fs.readdirSync(HISTORY_DIR).filter(f => /^\d{4}-\d{2}-\d{2}\.json$/.test(f)).sort();
    for (const f of dayFiles) {
      try {
        const j = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), 'utf8'));
        const rows = Array.isArray(j.rows) ? j.rows : [];
        for (const row of rows) { ws3.addRow(row); sheet3Count++; }
        if (rows.length) sheet3Days++;
      } catch(e) { log(`  跳过损坏历史文件 ${f}: ${e.message}`); }
    }
    log(`  Sheet3 完成: ${sheet3Count}行，覆盖${sheet3Days}个交易日`);

    // 列宽
    function autoWidth(ws, maxW) {
      for (let i = 1; i <= ws.columnCount; i++) {
        const col = ws.getColumn(i);
        let maxLen = 0;
        col.eachCell({ includeEmpty: true }, cell => {
          const v = cell.value == null ? '' : String(cell.value);
          let len = 0; for (const ch of v) len += (ch.charCodeAt(0) > 255 ? 2 : 1);
          if (len > maxLen) maxLen = len;
        });
        col.width = Math.min(Math.max(maxLen + 2, 10), maxW || 60);
      }
    }
    autoWidth(ws1, 60); autoWidth(ws2, 60); autoWidth(ws3, 60);

    const outName = `broker_quote_${BJ_DATE}.xlsx`;
    const outPath = path.join(BROKER_DIR, outName);
    const tmpPath = outPath + '.tmp';
    // 先写入临时文件，再 rename 替换，避免目标文件被占用时直接写失败
    await wb.xlsx.writeFile(tmpPath);
    let writeOk = false, lastErr = null;
    for (let wri = 0; wri < 5; wri++) {
      try {
        if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
        fs.renameSync(tmpPath, outPath);
        writeOk = true; break;
      } catch (e) {
        lastErr = e;
        log(`  Excel 写入/替换失败(${wri+1}/5): ${e.message.split('\n')[0]}`);
        await sleep(3000);
      }
    }
    if (!writeOk) {
      // 目标文件仍被占用（例如用户正在 Excel 中查看），改存到带时间戳的文件
      const altName = `broker_quote_${BJ_DATE}_${new Date().toISOString().slice(11,19).replace(/:/g,'')}.xlsx`;
      const altPath = path.join(BROKER_DIR, altName);
      fs.renameSync(tmpPath, altPath);
      log(`  ⚠️ 标准文件名被占用，已保存到: ${altPath}`);
      log(`   3 Sheet: 下框${filtered.length} + 当天汇总${sheet2Rows.length} + 历史累积${sheet3Count}(${sheet3Days}天)`);
    } else {
      log(`✅ 已保存: ${outPath}`);
      log(`   3 Sheet: 下框${filtered.length} + 当天汇总${sheet2Rows.length} + 历史累积${sheet3Count}(${sheet3Days}天)`);
    }

    await screenshot(page, '07-final');

  } catch (err) {
    log('❌ 错误:', err.message);
    log(err.stack);
    await screenshot(page, 'error');
  } finally {
    log('关闭浏览器');
    await context.close();
    try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch(e) {}
  }
}

main().catch(e => { console.error('Fatal:', e); process.exit(1); });
