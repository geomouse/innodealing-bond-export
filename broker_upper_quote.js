// 债立方 · 经纪商行情 · 上框(经纪商报价) 抽取
// 修复点：早期版本误落到 bond/detail 子路由导致读到"成交行情"下框(无 Ofr/Bid)。
//         本脚本直接带 hash 导航到主列表(#/bond/broker-market) 并校验 URL 不含 /detail/。
// 筛选：类型=全部信用(UI)，债项=AAA(UI)，剩余期限 0~1.9年(客户端)，Ofr>=1.90(客户端)
// 输出：Excel 一行一只债券(按债券去重，取最优=最低 Ofr)，含 发行人全称、担保人
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const USERNAME = 'wangle6';
const PASSWORD = '123456';
const LOGIN_URL = 'https://web.innotealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innotealing.com/quote-web/#/bond/broker-market';

const OFR_THRESHOLD = 1.90;
const MAX_MATURITY = 1.9;

const HEADLESS = true;
const WORKSPACE = __dirname;
const OUT_DIR = path.join(WORKSPACE, 'data', 'broker_upper');
fs.mkdirSync(OUT_DIR, { recursive: true });

function ts() { return new Date().toLocaleTimeString('zh-CN', { hour12: false }); }
function log(...a) { console.log(`[${ts()}]`, ...a); }
async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
async function shot(page, n) { try { await page.screenshot({ path: path.join(OUT_DIR, n + '.png'), timeout: 20000 }); log('  截图', n); } catch (e) {} }

// 带重试的导航：每次重试重建 page（避免 page.goto 失败后 page 进入坏状态），仅重试 DNS 解析类错误
async function gotoRetry(ctx, url, tries = 8) {
  let lastErr;
  for (let i = 0; i < tries; i++) {
    let page = null;
    try {
      page = await ctx.newPage();
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
      return page;
    } catch (e) {
      lastErr = e;
      if (page) { try { await page.close(); } catch (_) {} }
      if (/ERR_NAME_NOT_RESOLVED|ERR_DNS|ERR_NETWORK|ERR_CONNECTION/.test(e.message)) {
        log(`  导航 DNS/网络失败，${i + 1}/${tries} 重试...`);
        await sleep(3000);
        continue;
      }
      throw e;
    }
  }
  const marked = new Error('CONNECTION_FAILED: ' + (lastErr ? lastErr.message : 'unknown'));
  marked.isConn = true;
  throw marked;
}

function parseNum(s) { if (s == null) return NaN; const v = parseFloat(String(s).replace(/,/g, '').replace(/[\s%]/g, '')); return v; }
function parseMaturity(s) { const m = String(s).trim().match(/([0-9]+(?:\.[0-9]+)?)\s*[Yy年]/); return m ? parseFloat(m[1]) : NaN; }
function splitRating(s) { const t = String(s || '').trim(); if (!t) return { zhu: '', zhai: '' }; const p = t.split(/[/／]/); return { zhu: p[0] || '', zhai: p[1] || '' }; }

// 列名匹配（容错）
function colIndex(cols, ...candidates) {
  for (const c of candidates) {
    const i = cols.findIndex(col => col.name && (col.name === c || col.name.includes(c)));
    if (i >= 0) return i;
  }
  return -1;
}

(async () => {
  const launchOpts = { headless: HEADLESS, args: ['--no-sandbox', '--disable-setuid-sandbox'], viewport: { width: 3840, height: 1200 }, acceptDownloads: true };
  if (process.env.PW_CHROME) launchOpts.executablePath = process.env.PW_CHROME;
  let lastErr;
  for (let attempt = 1; attempt <= 40; attempt++) {
    const userDataDir = path.join(WORKSPACE, '.buq-' + Date.now());
    fs.mkdirSync(userDataDir, { recursive: true });
    const ctx = await chromium.launchPersistentContext(userDataDir, launchOpts);
    let page;
    let targetFrame = null;
    const issuerMap = new Map();
    try {
      // ===== 1. 登录（复用已验证方式：launchPersistentContext + networkidle，与 innodealing_broker.js 一致）=====
      log(`[尝试 ${attempt}/40] 1. 登录`);
      page = await ctx.newPage();
      await page.goto(LOGIN_URL, { waitUntil: 'networkidle', timeout: 30000 });
      await sleep(2000);
      if (page.url().includes('signin')) {
        await page.fill('input[placeholder*="手机号"]', USERNAME);
        await page.fill('input[type="password"]', PASSWORD);
        const terms = page.locator('label:has-text("我已阅读并同意相关服务条款和政策")');
        if (await terms.count() > 0) await terms.click();
        await sleep(400);
        await page.click('button:has-text("登录")');
        let ok = false; for (let i = 0; i < 45; i++) { await sleep(2000); if (!page.url().includes('signin')) { ok = true; break; } }
        if (!ok) throw new Error('登录失败');
        log('  登录成功'); await sleep(2500);
      } else log('  已有会话');

      // ===== 2. 导航到主列表（复用已验证方式：先到 quote-web 根再切 hash）=====
      log('2. 导航到经纪商行情主列表');
      await page.goto('https://web.innotealing.com/quote-web/', { waitUntil: 'domcontentloaded', timeout: 30000 });
      await sleep(5000);
    const frames = page.frames();
    targetFrame = frames.find(f => f.url().includes('quote-web') && !f.url().includes('auth-service'));
    if (!targetFrame) throw new Error('未找到 quote-web iframe');
    log('  iframe:', targetFrame.url());
    if (targetFrame.url().includes('/detail/')) {
      // 兜底：强制整页导航到主列表
      log('  ⚠️ 落在 detail 子路由，强制整页导航');
      await targetFrame.evaluate(() => { window.location.href = 'https://web.innotealing.com/quote-web/#/bond/broker-market'; });
      await sleep(6000);
      log('  重导航后 iframe:', targetFrame.url());
    }
    await shot(page, '01-initial');

    // 关弹窗
    await targetFrame.evaluate(() => {
      const mods = document.querySelectorAll('.ant-modal, [class*="Modal"], [class*="modal"], [class*="dialog"], [class*="Dialog"], [class*="adModal"], [class*="AdModal"]');
      mods.forEach(m => { const r = m.getBoundingClientRect(); if (r.width > 200 && r.height > 100) { const b = m.querySelector('button, [role="button"], [aria-label="Close"], .ant-modal-close'); if (b) b.click(); } });
    });
    await sleep(1000);
    try { await targetFrame.keyboard.press('Escape'); } catch (e) {}
    await sleep(1000);
    await shot(page, '02-after-modal');

    // ===== 3. 左栏筛选：全部信用 + 债项AAA =====
    async function clickTag(text) {
      return await targetFrame.evaluate((t) => {
        let el = null;
        document.querySelectorAll('*').forEach(e => { if (e.children.length === 0 && (e.textContent || '').trim() === t) { const r = e.getBoundingClientRect(); if (r.width > 0 && r.x < 360) el = e; } });
        if (!el) return false; el.click(); return true;
      }, text);
    }
    for (const t of ['全部信用', '债项AAA']) {
      const ok = await clickTag(t);
      log(`  点击 ${t}: ${ok ? '✅' : '⚠️未找到'}`);
      await sleep(2500);
    }
    await shot(page, '03-after-filters');

    // ===== 4. 定位上框(经纪商报价)并检测列 =====
    log('4. 定位上框(经纪商报价)');
    const upperMeta = await targetFrame.evaluate(() => {
      let upperY = 0;
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0 && (el.textContent || '').trim() === '经纪商') { const r = el.getBoundingClientRect(); if (r.y > 40 && r.y < 200) upperY = Math.max(upperY, r.y); }
      });
      if (!upperY) return { error: 'not found 经纪商 header' };
      const heads = [];
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length !== 0) return;
        const raw = (el.textContent || '').trim();
        if (!raw || raw.length > 35) return;
        const r = el.getBoundingClientRect();
        if (r.x < 300 || r.x > 4200) return;
        if (r.width < 4 || r.height < 4 || r.height > 60) return;
        if (Math.abs(r.y - upperY) <= 10) heads.push({ x: Math.round(r.x + r.width / 2), w: Math.round(r.width), name: raw.replace(/[↓↑↕\s]/g, '') });
      });
      heads.sort((a, b) => a.x - a.w / 2 - (b.x - b.w / 2));
      const uniq = []; const seen = new Set();
      for (const h of heads) { const key = `${h.name}@${Math.round(h.x / 40)}`; if (seen.has(key)) continue; seen.add(key); uniq.push(h); }
      uniq.sort((a, b) => a.x - b.x);
      return { headerY: upperY, cols: uniq };
    });
    if (upperMeta.error) throw new Error(upperMeta.error);
    const cols = upperMeta.cols;
    log('  上框列(' + cols.length + '):', cols.map(c => c.name).join(' | '));
    const iBroker = colIndex(cols, '经纪商');
    const iOfr = colIndex(cols, 'Ofr', '卖出', '最优Ofr');
    const iBid = colIndex(cols, 'Bid', '买入', '最优Bid');
    const iName = colIndex(cols, '债券简称');
    const iCode = colIndex(cols, '债券代码');
    const iMat = colIndex(cols, '剩余期限');
    const iRating = colIndex(cols, '主/债', '债项', '评级');
    const iIssuer = colIndex(cols, '发行人');
    const iGuar = colIndex(cols, '担保人');
    log(`  关键列: 经纪商=${iBroker} Ofr=${iOfr} Bid=${iBid} 简称=${iName} 代码=${iCode} 期限=${iMat} 评级=${iRating} 发行人=${iIssuer} 担保人=${iGuar}`);
    if (iBroker < 0 || iOfr < 0) throw new Error('上框缺少 经纪商/Ofr 列，可能定位错框');
    const haveIssuerCol = iIssuer >= 0, haveGuarCol = iGuar >= 0;
    log(`  上框自带发行人列=${haveIssuerCol} 担保人列=${haveGuarCol}`);

    // 上框滚动容器
    const scrollInfo = await targetFrame.evaluate((upperY) => {
      let hdr = null;
      document.querySelectorAll('*').forEach(el => { if (el.children.length === 0 && (el.textContent || '').trim() === '经纪商') { const r = el.getBoundingClientRect(); if (r.y > 40 && r.y < 200 && Math.abs(r.y - upperY) <= 8) hdr = el; } });
      if (!hdr) return { found: false };
      let el = hdr.parentElement;
      while (el && el.parentElement) { if (el.scrollHeight > el.clientHeight + 50 && el.clientHeight > 80) { el.__upperBox = true; const r = el.getBoundingClientRect(); return { found: true, top: Math.round(r.top), sh: el.scrollHeight, ch: el.clientHeight }; } el = el.parentElement; }
      return { found: false };
    }, upperMeta.headerY);
    log('  上框滚动容器:', JSON.stringify(scrollInfo));

    async function scrollUpper(px) { return await targetFrame.evaluate((px) => { let c = null; document.querySelectorAll('*').forEach(el => { if (el.__upperBox) c = el; }); if (!c) return false; const old = c.scrollTop; c.scrollTop += px; return c.scrollTop !== old; }, px); }
    async function getUpperScroll() { return await targetFrame.evaluate(() => { let c = null; document.querySelectorAll('*').forEach(el => { if (el.__upperBox) c = el; }); if (!c) return null; return { top: c.scrollTop, sh: c.scrollHeight, ch: c.clientHeight }; }); }

    // 快照
    async function snapshotUpper() {
      return await targetFrame.evaluate(({ headerY, cols }) => {
        let hdr = null;
        document.querySelectorAll('*').forEach(el => { if (el.children.length === 0 && (el.textContent || '').trim() === '经纪商') { const r = el.getBoundingClientRect(); if (r.y > 40 && r.y < 200 && Math.abs(r.y - headerY) <= 8) hdr = el; } });
        let box = null, p = hdr; while (p) { if (p.className && p.className.toString && /dmui-vt-background/.test(p.className.toString())) { box = p; break; } p = p.parentElement; }
        if (!box) return { rows: [], diag: 'no upper box' };
        const br = box.getBoundingClientRect(); const yMax = br.bottom;
        const skip = new Set(['有买', '有卖', '筛选', '清空筛选', '全部信用', '全部利率', '国债', '央票', '地方债', '政策银行债', '双边']);
        const dcells = [];
        box.querySelectorAll('*').forEach(el => {
          const r = el.getBoundingClientRect();
          if (r.width < 5 || r.height < 5 || r.height > 200) return;
          if (r.x < 300 || r.x > 4200) return;
          if (el.children.length === 0) { const t = (el.textContent || '').trim(); if (!t || t.length > 50 || skip.has(t)) return; if (r.y > headerY + 6 && r.y < yMax) dcells.push({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y), w: Math.round(r.width), text: t }); return; }
          let dt = ''; for (const n of el.childNodes) if (n.nodeType === Node.TEXT_NODE) dt += n.textContent; dt = dt.trim();
          if (dt && !skip.has(dt) && dt.length <= 50) { let cov = false; for (const c of el.querySelectorAll('*')) if (c.children.length === 0 && c.textContent.trim() === dt) { cov = true; break; } if (!cov && r.y > headerY + 6 && r.y < yMax) dcells.push({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y), w: Math.round(r.width), text: dt }); }
        });
        dcells.sort((a, b) => a.y - b.y || a.x - b.x);
        const rawRows = []; let cur = null, curMaxY = -Infinity;
        for (const c of dcells) { if (!cur || (c.y - curMaxY) > 13) { cur = []; rawRows.push(cur); } cur.push(c); if (c.y > curMaxY) curMaxY = c.y; }
        cols.sort((a, b) => a.x - b.x);
        const intervals = cols.map((c, i) => { const prev = i > 0 ? cols[i - 1].x : c.x - 80; const next = i + 1 < cols.length ? cols[i + 1].x : c.x + 80; return { left: (c.x + prev) / 2, right: (c.x + next) / 2, seed: c.x }; });
        const outRows = [];
        for (const row of rawRows) {
          const vals = new Array(cols.length).fill('');
          for (const c of row) { let best = -1, bestD = Infinity; for (let i = 0; i < cols.length; i++) { if (c.x >= intervals[i].left && c.x < intervals[i].right) { const d = Math.abs(c.x - cols[i].x); if (d < bestD) { bestD = d; best = i; } } } if (best >= 0) { const ex = vals[best]; if ([colIndex(cols, '债券简称'), colIndex(cols, '债券代码')].includes(best)) { if (!ex || c.text.length > ex.length) vals[best] = c.text; } else if (!ex) vals[best] = c.text; } }
          const name = vals[iName] || ''; const code = vals[iCode] || ''; const broker = vals[iBroker] || '';
          if (!name && !code) continue;
          outRows.push({ y: row[0].y, vals, name, code, broker });
        }
        return { rows: outRows, diag: `cells=${dcells.length}` };
      }, { headerY: upperMeta.headerY, cols });
    }

    // 滚动循环抓取
    const allRows = new Map();
    let iter = 0, sameCount = 0, lastTop = -1, bottom = false;
    while (iter < 120 && !bottom) {
      iter++;
      const snap = await snapshotUpper();
      let added = 0;
      for (const r of snap.rows) { const key = `${r.code || r.name}_${r.broker}_${r.vals.slice(0, 6).join('|')}`; if (!allRows.has(key)) { allRows.set(key, r); added++; } }
      log(`  [快照${iter}] 本帧${snap.rows.length}行 新增${added} 累计${allRows.size} (${snap.diag})`);
      const si = await getUpperScroll();
      if (si) { if (si.sh - si.ch - si.top <= 2) { bottom = true; log('  到达底部'); } if (si.top === lastTop) sameCount++; else { sameCount = 0; lastTop = si.top; } if (sameCount >= 4) { log('  滚动不再前进'); break; } }
      else if (iter > 1) { log('  未找到滚动容器'); break; }
      if (!bottom) { const moved = await scrollUpper(350); if (!moved && iter > 1) { log('  滚动未移动'); break; } await sleep(500); }
    }
    await scrollUpper(-100000); await sleep(1000);

    // 客户端过滤
    log('5. 过滤：债项AAA + 剩余期限0~1.9年 + Ofr>=1.90');
    const gv = (r, i) => (i >= 0 ? (r.vals[i] || '') : '');
    const quoteRows = [];
    for (const r of allRows.values()) {
      const mat = parseMaturity(gv(r, iMat));
      const ofr = parseNum(gv(r, iOfr));
      const { zhai } = splitRating(gv(r, iRating));
      if (zhai !== 'AAA') continue;
      if (isNaN(mat) || mat <= 0 || mat > MAX_MATURITY) continue;
      if (isNaN(ofr) || ofr < OFR_THRESHOLD) continue;
      quoteRows.push({ name: r.name, code: r.code, broker: r.broker, maturityStr: gv(r, iMat), ratingStr: gv(r, iRating), ofr: gv(r, iOfr), bid: gv(r, iBid), issuerCol: haveIssuerCol ? gv(r, iIssuer) : '', guarCol: haveGuarCol ? gv(r, iGuar) : '' });
    }
    log(`  过滤后 ${quoteRows.length} 条报价`);

    // 按债券去重，取最低 Ofr（最优买价）
    const best = new Map();
    for (const q of quoteRows) { const ex = best.get(q.code); if (!ex || parseNum(q.ofr) < parseNum(ex.ofr)) best.set(q.code, q); }
    const uniqueBonds = Array.from(best.values()).sort((a, b) => parseNum(a.ofr) - parseNum(b.ofr));
    log(`  去重后 ${uniqueBonds.length} 只债券`);

    // 保存中间结果（防止详情页阶段中断丢数据）
    fs.writeFileSync(path.join(OUT_DIR, '_extracted.json'), JSON.stringify({ quoteRows, uniqueBonds, haveIssuerCol, haveGuarCol }, null, 1), 'utf8');

    if (uniqueBonds.length === 0) log('  无满足条件数据');
    else for (const b of uniqueBonds.slice(0, 20)) log(`    ${b.name} ${b.code} ${b.maturityStr} Ofr=${b.ofr} (${b.broker})`);

    // ===== 6. 发行人全称 + 担保人 =====
    const needDetail = !(haveIssuerCol && haveGuarCol) || uniqueBonds.some(b => !b.issuerCol || !b.guarCol);
    if (uniqueBonds.length > 0 && needDetail) {
      log('6. 详情页提取发行人全称、担保人');
      const iframeOffset = await page.evaluate(() => { const f = Array.from(document.querySelectorAll('iframe')).find(x => x.src && x.src.includes('quote-web')); if (!f) return { x: 0, y: 0 }; const r = f.getBoundingClientRect(); return { x: r.x, y: r.y }; }).catch(() => ({ x: 0, y: 0 }));
      const getNavSig = () => [page.url(), ...page.frames().map(f => f.url())].join('|');
      const activeDetailFrame = async () => {
        for (const f of page.frames()) { const u = f.url(); if (u.includes('bond/detail') || u.includes('/detail/')) return f; }
        for (const f of page.frames()) { const body = await f.evaluate(() => document.body?.textContent || '').catch(() => ''); if (/性质|主体评级|债项评级|债券代码/.test(body) && body.length > 200) return f; }
        return null;
      };
      const isDetailContent = (f) => f.evaluate(() => /性质|主体评级|债项评级/.test(document.body?.textContent || '')).catch(() => false);

      async function dblclickFirst(name, code) {
        const info = await targetFrame.evaluate((nm) => { const els = Array.from(document.querySelectorAll('*')).filter(el => { const t = (el.textContent || '').trim(); return t === nm || (t.startsWith(nm) && t.length <= nm.length + 8); }); if (!els.length) return null; els.sort((a, b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y); const t = els.find(el => { const r = el.getBoundingClientRect(); return r.width > 30 && r.width < 400 && r.height > 10 && r.height < 60; }) || els[0]; t.scrollIntoView({ block: 'center' }); const r = t.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; }, name).catch(() => null);
        if (!info) return false;
        const px = iframeOffset.x + info.x, py = iframeOffset.y + info.y;
        log(`  双击进入详情: ${name} (${Math.round(px)},${Math.round(py)})`);
        await page.mouse.move(px, py); await sleep(120); await page.mouse.dblclick(px, py);
        const prev = getNavSig();
        for (let i = 0; i < 24; i++) { await sleep(500); const df = await activeDetailFrame(); if (df && getNavSig() !== prev && await isDetailContent(df)) { await sleep(800); return true; } }
        return false;
      }
      async function searchSwitch(code, name) {
        const df = await activeDetailFrame(); if (!df) return false;
        const clean = String(code || '').split('.')[0];
        const box = await df.evaluate(() => { const i = document.querySelector('input.ant-select-search__field') || Array.from(document.querySelectorAll('input')).find(el => String(el.className || '').includes('ant-select-search')); if (!i) return null; const r = i.getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; }).catch(() => null);
        if (!box) return false;
        try { await df.click('input.ant-select-search__field', { timeout: 3000 }); } catch (e) { await page.mouse.click(iframeOffset.x + box.x, iframeOffset.y + box.y); }
        await sleep(250); await page.keyboard.press('Control+A'); await sleep(120); await page.keyboard.press('Backspace'); await sleep(120);
        await page.keyboard.type(clean, { delay: 30 }); await sleep(900); await page.keyboard.press('Enter');
        const prev = getNavSig();
        for (let i = 0; i < 24; i++) { await sleep(500); if (getNavSig() !== prev || await df.evaluate(({ c, n }) => { const b = document.body?.textContent || ''; return (c && b.includes(c)) || (n && b.includes(n)); }, { c: clean, n: name })) { await sleep(800); return true; } }
        return false;
      }
      async function extractDetail() {
        const df = await activeDetailFrame(); if (!df) return { issuerFull: '', guarantor: '', method: 'no_frame' };
        return await df.evaluate(() => {
          const res = { issuerFull: '', guarantor: '', guaranteeMethod: '', method: '' };
          const all = Array.from(document.querySelectorAll('*'));
          // 发行人全称（顶部信息栏）
          const cands = all.filter(el => { const t = (el.textContent || '').trim(); const r = el.getBoundingClientRect(); return r.y >= 30 && r.y <= 260 && r.x >= 500 && t.length >= 4 && t.length <= 60 && !/(成交|行情|最新|价格|Bid|Ofr|笔数|更新时间|流动性|中债|中证|主承销|发行日|剩余期限|收益率|债券简称|债券代码)/.test(t) && /有限公司|公司|银行|政府|财政部|集团|投资|控股|开发|管理局|厅|委员会|国资|资产|有限/.test(t); }).sort((a, b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
          if (cands.length) { let name = cands[0].textContent.trim(); const rc = name.search(/主体评级|债项评级|评级/); if (rc > 0) name = name.slice(0, rc).trim(); res.issuerFull = name; res.method = 'top'; }
          // 担保人 / 担保方式
          for (let i = 0; i < all.length; i++) { const el = all[i]; const t = (el.textContent || '').trim(); if (t === '担保人' || t === '担保方' || t === '担保' || t === '担保方式') { for (let j = i + 1; j < Math.min(i + 10, all.length); j++) { const sib = all[j]; const st = (sib.textContent || '').trim(); const sr = sib.getBoundingClientRect(); if (!st || st === t || st.length > 60) continue; if (st === '——' || st === '--' || st === '-' || st === '无') continue; if (/^[0-9\.\-]+$/.test(st) && st.length < 10) continue; if (t === '担保方式') res.guaranteeMethod = st; else res.guarantor = st; break; } } }
          if (!res.guarantor && /无担保/.test(res.guaranteeMethod || '')) res.guarantor = '';
          return res;
        });
      }
      let onDetail = false;
      for (let a = 0; a < 3 && !onDetail; a++) { onDetail = await dblclickFirst(uniqueBonds[0].name, uniqueBonds[0].code); log(`  首只进详情(尝试${a + 1}): ${onDetail ? '✅' : '❌'}`); if (!onDetail) await sleep(1500); }
      if (!onDetail) log('  ⚠️ 无法进入详情页，发行人/担保人将留空');
      for (let i = 0; i < uniqueBonds.length; i++) {
        const b = uniqueBonds[i];
        if (i > 0 || !onDetail) { if (!onDetail) continue; const ok = await searchSwitch(b.code, b.name); if (!ok) { log(`  [${i + 1}] ${b.name} 导航失败`); continue; } }
        let ex = await extractDetail(); let rt = 0;
        while ((!ex.issuerFull) && rt < 2) { rt++; await sleep(1200); ex = await extractDetail(); }
        issuerMap.set(b.code, ex);
        log(`  [${i + 1}/${uniqueBonds.length}] ${b.name} → 发行人:${ex.issuerFull || '(空)'} 担保人:${ex.guarantor || '(空)'}`);
      }
      try { await targetFrame.evaluate(() => { window.location.hash = '#/bond/broker-market'; window.dispatchEvent(new HashChangeEvent('hashchange')); }); await sleep(3000); } catch (e) {}
    }

    // 上框自带发行人/担保人列时直接用
    if (haveIssuerCol || haveGuarCol) {
      for (const b of uniqueBonds) { const ex = issuerMap.get(b.code) || {}; if (haveIssuerCol && !ex.issuerFull) ex.issuerFull = b.issuerCol || ''; if (haveGuarCol && !ex.guarantor) ex.guarantor = b.guarCol || ''; issuerMap.set(b.code, ex); }
    }
    fs.writeFileSync(path.join(OUT_DIR, '_issuers.json'), JSON.stringify(Array.from(issuerMap.entries()), null, 1), 'utf8');

    // ===== 7. 生成 Excel =====
    log('7. 生成 Excel');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('经纪商报价_上框');
    const headers = ['债券简称', '债券代码', '剩余期限(年)', '债项评级', 'Ofr报价(%)', '经纪商', '发行人全称', '担保人'];
    ws.addRow(headers);
    for (const b of uniqueBonds) {
      const ex = issuerMap.get(b.code) || {};
      ws.addRow([b.name, b.code, (parseMaturity(b.maturityStr) || b.maturityStr), splitRating(b.ratingStr).zhai, (parseNum(b.ofr) || b.ofr), b.broker, ex.issuerFull || '', ex.guarantor || '']);
    }
    for (let i = 1; i <= headers.length; i++) { const col = ws.getColumn(i); let max = 0; col.eachCell({ includeEmpty: true }, cell => { const v = cell.value == null ? '' : String(cell.value); let len = 0; for (const ch of v) len += (ch.charCodeAt(0) > 255 ? 2 : 1); if (len > max) max = len; }); col.width = Math.min(Math.max(max + 2, 10), 70); }
    const fileName = `经纪商报价_上框_${new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10)}.xlsx`;
    const outPath = path.join(OUT_DIR, fileName);
    await wb.xlsx.writeFile(outPath);
    log(`✅ 已保存: ${outPath} (${uniqueBonds.length}行)`);

    const issuerFilled = uniqueBonds.filter(b => (issuerMap.get(b.code)?.issuerFull || '').length > 0).length;
    const guarFilled = uniqueBonds.filter(b => (issuerMap.get(b.code)?.guarantor || '').length > 0).length;
    log(`  自检: 债券=${uniqueBonds.length} 发行人非空=${issuerFilled}/${uniqueBonds.length} 担保人非空=${guarFilled}/${uniqueBonds.length}`);
    await shot(page, '08-final');
    await ctx.close();
    try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch (e) {}
    return; // 成功，直接退出

  } catch (err) {
    const isConn = (err && err.isConn) || /ERR_NAME_NOT_RESOLVED|ERR_DNS|ERR_NETWORK|ERR_CONNECTION|net::/.test(String((err && err.message) || ''));
    if (isConn) {
      log(`  ⚠️ 连接失败(第${attempt}次)，3s 后重建浏览器重试...`);
      try { await ctx.close(); } catch (_) {}
      try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch (e) {}
      await sleep(3000);
      continue;
    }
    log('❌ 错误:', err.message); log(err.stack); try { await shot(page, 'error'); } catch (_) {}
    try { await ctx.close(); } catch (_) {}
    lastErr = err; break;
  }
  }
  if (lastErr) { log('❌ 最终失败:', lastErr.message); process.exit(1); }
})();
