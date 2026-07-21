// 债立方二级行情 - 成交数据抓取（v3 - 重构版）
// 路径：首页 → 二级行情 → 我的关注 → 主体勾选a → 最新成交倒序 → 复制价格
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const USERNAME = process.env.INNODEALING_USERNAME || 'wangle6';
const PASSWORD = process.env.INNODEALING_PASSWORD || '123456';
const LOGIN_URL = 'https://web.innodealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innodealing.com/quote-web/#/bond/my-focus';

const ENTITY_FILTER = 'all-A';
const MIN_VOLUME = 2.0;

// 云端/本地通用开关：HEADLESS=true 或 CI 环境走无头 + playwright 自带 chromium；本地默认有头 + msedge
const HEADLESS = process.env.HEADLESS === 'true' || process.env.CI === 'true';
const BROWSER_CHANNEL = process.env.BROWSER_CHANNEL != null
  ? process.env.BROWSER_CHANNEL
  : (HEADLESS ? '' : 'msedge');

const WORKSPACE = __dirname;
const DOWNLOAD_DIR = path.join(WORKSPACE, 'downloads');
const SCREENSHOT_DIR = path.join(WORKSPACE, 'screenshots');
// Sheet4 历史累积的单一数据源（本地/云端共享）：每天一份 daily/{date}.json
const SECONDARY_DIR = process.env.SECONDARY_DIR || path.join(WORKSPACE, 'data', 'secondary');
const HISTORY_DIR = path.join(SECONDARY_DIR, 'daily');
[ DOWNLOAD_DIR, SCREENSHOT_DIR, SECONDARY_DIR, HISTORY_DIR ].forEach(d => fs.mkdirSync(d, { recursive: true }));

// 统一用北京时间(UTC+8)的日期作为文件名 key 与 Sheet4 去重键
const BJ_DATE = new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10);

function ts() { return new Date().toLocaleTimeString('zh-CN', { hour12: false }); }
function log(...args) { console.log(`[${ts()}]`, ...args); }
async function screenshot(page, name) {
  const p = path.join(SCREENSHOT_DIR, `${name}.png`);
  try {
    await page.screenshot({ path: p, fullPage: false, timeout: 15000, animations: 'disabled' });
    log(`截图: ${name}`);
  } catch (e) {
    log(`截图跳过(${name}): ${e.message.split('\n')[0]}`);
  }
}
async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function main() {
  log('=== 启动浏览器 ===');
  // 用临时 user-data-dir 避免被污染
  const userDataDir = path.join(WORKSPACE, '.chrome-data-' + Date.now());
  fs.mkdirSync(userDataDir, { recursive: true });
  const launchOpts = {
    headless: HEADLESS,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    viewport: { width: 1920, height: 1080 },
    acceptDownloads: true
  };
  if (BROWSER_CHANNEL) launchOpts.channel = BROWSER_CHANNEL; // 本地 msedge；云端留空用自带 chromium
  log(`浏览器模式: headless=${HEADLESS} channel=${BROWSER_CHANNEL || 'chromium(自带)'}`);
  const browser = await chromium.launchPersistentContext(userDataDir, launchOpts);
  const context = browser;  // launchPersistentContext 直接返回 context
  const page = await context.newPage();

  // 关闭弹窗的辅助函数 - 在 main 内部定义，但用 const + 闭包传入 targetFrame
  let targetFrame = null;
  async function closeAllModals(maxRounds = 5) {
    if (!targetFrame) return { clicked: 0, details: [] };
    for (let i = 0; i < maxRounds; i++) {
      const closeInfo = await targetFrame.evaluate(() => {
        const result = { clicked: 0, details: [] };
        // 只找可见的模态容器（必须 > 50x50 避免点到普通 UI）
        const modalContainers = document.querySelectorAll(
          '.ant-modal, .dmuiv4-modal, [class*="ant-modal"]:not([class*="tab"]), [class*="dmuiv4-modal"], [class*="Modal"], [class*="adModal"], [class*="AdModal"], [class*="banner-modal"], [class*="popup"], [class*="dialog"]'
        );
        const visibleModals = [];
        for (const m of modalContainers) {
          const rect = m.getBoundingClientRect();
          // 模态通常有大尺寸+居中位置
          if (rect.width > 200 && rect.height > 100 && rect.x >= 0 && rect.y >= 0) {
            visibleModals.push(m);
          }
        }
        result.details.push(`可见模态: ${visibleModals.length}`);

        for (const m of visibleModals) {
          const closeSels = [
            '.ant-modal-close', '.dmuiv4-modal-close',
            '[class*="modal-close"]', '[class*="ModalClose"]',
            '[class*="closeIcon"]', '[class*="CloseIcon"]',
            '[class*="icon-close"]', '[class*="iconClose"]',
            '[class*="close-btn"]', '[class*="closeBtn"]',
            'button[aria-label*="close" i]', 'span[aria-label*="close" i]',
            'button[aria-label*="关闭"]', 'span[aria-label*="关闭"]',
            'button'
          ];
          // 找模态内的所有按钮，优先点看起来像关闭的（右上角、小尺寸）
          const allBtns = m.querySelectorAll('button, [role="button"], span[aria-label]');
          for (const btn of allBtns) {
            const r = btn.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) continue;
            const mRect = m.getBoundingClientRect();
            // 右上角区域：x > mRect.x + mRect.width*0.7 && y < mRect.y + 80
            const isTopRight = r.x > mRect.x + mRect.width * 0.6 && r.y < mRect.y + 80;
            // 或者包含 close/关闭/X 关键字
            const text = (btn.textContent || btn.getAttribute('aria-label') || '').toLowerCase();
            const ariaLabel = (btn.getAttribute('aria-label') || '').toLowerCase();
            const className = (btn.className || '').toLowerCase();
            const isCloseBtn =
              text.includes('close') || text.includes('关闭') || text === '×' || text === '✕' ||
              ariaLabel.includes('close') || ariaLabel.includes('关闭') ||
              className.includes('close');
            if (isTopRight && r.width < 60 && r.height < 60) {
              btn.click();
              result.clicked++;
              result.details.push(`topRight ${Math.round(r.x)},${Math.round(r.y)}`);
            } else if (isCloseBtn) {
              btn.click();
              result.clicked++;
              result.details.push(`text "${text.substring(0, 10)}"`);
            }
          }
        }
        return result;
      });

      if (closeInfo.clicked > 0) {
        log(`  关闭弹窗第 ${i + 1} 轮: ${closeInfo.clicked} 个, ${closeInfo.details.slice(0, 2).join(' | ')}`);
      }
      if (closeInfo.clicked === 0) break;
      await sleep(500);
    }
    // ESC 保险
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
      if (await termsLabel.count() > 0) {
        await termsLabel.click();
        log('已勾选服务条款');
      }
      await sleep(500);
      await page.click('button:has-text("登录")');
      // 健壮登录：网站慢时跳转可能 >20s，用轮询兜底（最长 90s），只要离开 signin 即成功
      let logged = false;
      for (let i = 0; i < 45; i++) {
        await sleep(2000);
        if (!page.url().includes('signin')) { logged = true; break; }
      }
      if (!logged) throw new Error('登录失败：90s 内未离开 signin 页面（可能账号/密码错误或验证码）');
      log('登录成功:', page.url());
      await sleep(2000); // 等 dashboard 稳定
    }
    await sleep(2000);
    await screenshot(page, '02-after-login');

    // ===== 2. 导航到二级行情 - 我的关注 =====
    log('2. 导航到我的关注');
    await sleep(2000);

    // 关闭 dashboard 上的弹窗
    try {
      await page.evaluate(() => {
        document.querySelectorAll('.ant-modal-close, .dmuiv4-modal-close, [class*="modal-close"]').forEach(el => {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && r.height > 0) el.click();
        });
      });
    } catch (e) {}
    await sleep(500);

    // 先到 quote-web 根路径
    log('  步骤a: 跳到 quote-web 根路径');
    await page.goto('https://web.innodealing.com/quote-web/', { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(5000);

    // 拿到 quote-web frame
    let allFrames = page.frames();
    log(`  共 ${allFrames.length} 个 frame:`);
    for (const f of allFrames) {
      log(`    - ${f.url().substring(0, 120)}`);
    }
    targetFrame = allFrames.find(f => f.url().includes('quote-web') && !f.url().includes('auth-service'));
    if (!targetFrame) throw new Error('未找到 quote-web iframe');
    log(`  使用 iframe: ${targetFrame.url()}`);

    // 在 iframe 内点击"我的关注" tab
    log('  步骤b: 在 iframe 内查找"我的关注" tab');
    // 先调试：找所有可能含"我的关注"或"my-focus"的元素
    const debugTabs = await targetFrame.evaluate(() => {
      const result = [];
      document.querySelectorAll('*').forEach(el => {
        const text = (el.textContent || '').trim();
        // 找顶层 tab 栏（位置在 y < 100 或 x < 200）
        const rect = el.getBoundingClientRect();
        if (text === '我的关注' || text === '我的关注' || text.startsWith('我的关注')) {
          result.push({
            tag: el.tagName,
            cls: (el.className || '').substring(0, 100),
            text: text.substring(0, 30),
            x: Math.round(rect.x), y: Math.round(rect.y),
            w: Math.round(rect.width), h: Math.round(rect.height),
            role: el.getAttribute('role') || '',
            href: el.getAttribute('href') || ''
          });
        }
      });
      return result;
    });
    log(`  找到 ${debugTabs.length} 个"我的关注"相关元素:`);
    for (const t of debugTabs) {
      log(`    ${t.tag}.${t.cls} "${t.text}" at (${t.x},${t.y}) ${t.w}x${t.h} role=${t.role} href=${t.href}`);
    }

    // 找所有顶部 tab 元素
    const allTopTabs = await targetFrame.evaluate(() => {
      const result = [];
      document.querySelectorAll('a, [class*="tab"], [class*="Tab"], [role="tab"], [class*="menu"]').forEach(el => {
        const text = (el.textContent || '').trim().substring(0, 20);
        const rect = el.getBoundingClientRect();
        // 顶部 tab 通常 y < 100
        if (rect.y < 100 && rect.width > 20 && rect.height > 10 && text) {
          result.push({
            tag: el.tagName,
            cls: (el.className || '').substring(0, 80),
            text,
            x: Math.round(rect.x), y: Math.round(rect.y),
            w: Math.round(rect.width), h: Math.round(rect.height),
            href: el.getAttribute('href') || ''
          });
        }
      });
      return result;
    });
    log(`  顶部 tab 元素 ${allTopTabs.length} 个:`);
    for (const t of allTopTabs.slice(0, 30)) {
      log(`    ${t.tag} "${t.text}" at (${t.x},${t.y}) ${t.w}x${t.h} href=${t.href.substring(0, 50)}`);
    }

    // 点击"我的关注" tab
    log('  步骤c: 直接修改 hash 触发 SPA 路由');
    const hashChangeResult = await targetFrame.evaluate(() => {
      const before = window.location.hash;
      // 设置新 hash
      window.location.hash = '#/bond/my-focus';
      // 触发 hashchange 事件
      window.dispatchEvent(new HashChangeEvent('hashchange'));
      return { before, after: window.location.hash };
    });
    log(`  hash: ${hashChangeResult.before} -> ${hashChangeResult.after}`);
    await sleep(4000);

    // 重新获取 frame
    allFrames = page.frames();
    log(`  重新查询，共 ${allFrames.length} 个 frame:`);
    for (const f of allFrames) {
      log(`    - ${f.url().substring(0, 120)}`);
    }
    targetFrame = allFrames.find(f => f.url().includes('my-focus'))
                || allFrames.find(f => f.url().includes('quote-web'));
    if (!targetFrame) throw new Error('未找到 quote-web iframe');
    log(`  当前 iframe: ${targetFrame.url()}`);

    // ===== 3. 关闭所有弹窗 =====
    log('3. 关闭弹窗');
    await closeAllModals();
    await sleep(1500);
    await screenshot(page, '03-modals-closed');

    // ===== 4. 点击左侧"主体" tab =====
    log('4. 点击左侧"主体" tab');
    const clickedSubject = await targetFrame.evaluate(() => {
      const all = document.querySelectorAll('*');
      for (const el of all) {
        if (el.children.length === 0 && el.textContent?.trim() === '主体') {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0) {
            el.click();
            return { x: rect.x, y: rect.y, w: rect.width, h: rect.height };
          }
        }
      }
      return null;
    });
    if (clickedSubject) {
      log(`已点击"主体" tab at (${clickedSubject.x},${clickedSubject.y})`);
    } else {
      log('未找到"主体" tab！');
    }
    await sleep(1500);
    await closeAllModals();
    await screenshot(page, '04-subject-clicked');

    // ===== 5. 勾选主体组"a" =====
    log('5. 勾选主体组 a');
    const aClickResult = await targetFrame.evaluate((filterName) => {
      const candidates = [];
      document.querySelectorAll('label, div, span').forEach(el => {
        const text = el.textContent?.trim() || '';
        if (text === filterName) {
          const rect = el.getBoundingClientRect();
          if (rect.width > 0 && rect.height > 0 && rect.x < 250) {
            candidates.push({
              tag: el.tagName,
              cls: el.className?.substring(0, 80) || '',
              text,
              x: Math.round(rect.x), y: Math.round(rect.y),
              w: Math.round(rect.width), h: Math.round(rect.height)
            });
          }
        }
      });

      const result = { candidates, clicked: null };
      for (const c of candidates) {
        const textNode = Array.from(document.querySelectorAll('*')).find(
          el => el.textContent?.trim() === filterName && el.getBoundingClientRect().x < 250
        );
        if (textNode) {
          let parent = textNode.parentElement;
          let foundCheckbox = null;
          for (let i = 0; i < 6 && parent; i++) {
            const cb = parent.querySelector('input[type="checkbox"]');
            if (cb) { foundCheckbox = cb; break; }
            parent = parent.parentElement;
          }
          if (foundCheckbox) {
            foundCheckbox.click();
            result.clicked = { via: 'parent-checkbox', pos: c };
            return result;
          }
          const prev = textNode.previousElementSibling;
          if (prev && prev.matches('input[type="checkbox"]')) {
            prev.click();
            result.clicked = { via: 'sibling-checkbox', pos: c };
            return result;
          }
          textNode.click();
          result.clicked = { via: 'self', pos: c };
          return result;
        }
      }
      return result;
    }, ENTITY_FILTER);

    log(`  找到 ${aClickResult.candidates?.length || 0} 个 "a" 候选`);
    if (aClickResult.candidates) {
      for (const c of aClickResult.candidates) {
        log(`    ${c.tag} "${c.text}" at (${c.x},${c.y}) ${c.w}x${c.h}`);
      }
    }
    if (aClickResult.clicked) {
      log(`  已点击 a via: ${aClickResult.clicked.via}`);
    } else {
      log('  ❌ 未找到 "a" 元素');
    }
    await sleep(3000);
    await closeAllModals();
    await screenshot(page, '05-a-checked');

    // ===== 6. 点击"最新成交"列两次倒序 =====
    log('6. 点击"最新成交"列两次倒序');
    const sortResult = await targetFrame.evaluate(() => {
      const headers = document.querySelectorAll('th, [class*="header"], [class*="Header"], [class*="col-head"], [class*="ColHead"]');
      const candidates = [];
      for (const h of headers) {
        const text = h.textContent?.trim() || '';
        if (text.includes('最新成交')) {
          const rect = h.getBoundingClientRect();
          candidates.push({
            tag: h.tagName, cls: h.className?.substring(0, 80) || '',
            text: text.substring(0, 30),
            x: Math.round(rect.x), y: Math.round(rect.y),
            w: Math.round(rect.width)
          });
        }
      }
      // 也找叶子节点
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0) {
          const text = el.textContent?.trim() || '';
          if (text.startsWith('最新成交')) {
            const rect = el.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {
              candidates.push({
                tag: el.tagName + '(leaf)', cls: el.className?.substring(0, 80) || '',
                text: text.substring(0, 30),
                x: Math.round(rect.x), y: Math.round(rect.y),
                w: Math.round(rect.width)
              });
            }
          }
        }
      });
      return candidates;
    });
    log(`  找到 ${sortResult.length} 个"最新成交"表头候选`);
    for (const s of sortResult.slice(0, 10)) {
      log(`    ${s.tag}.${s.cls} "${s.text}" at (${s.x},${s.y}) ${s.w}w`);
    }

    // 选 y 最大的"最新成交"表头 = 下面的框
    let headerY = 80;
    if (sortResult.length > 0) {
      const sortedHeaders = [...sortResult].sort((a, b) => b.y - a.y);
      const headerTarget = sortedHeaders[0];
      headerY = headerTarget.y;
      log(`  所有"最新成交"表头位置 (按 y 降序):`);
      for (const h of sortedHeaders) {
        log(`    ${h.tag} at y=${h.y} x=${h.x} "${h.text}"`);
      }
      log(`  [下面框] 表头目标: ${headerTarget.tag} at (${headerTarget.x},${headerTarget.y}) ${headerTarget.w}w "${headerTarget.text}"`);

      // 点击两次倒序（选 y 最大的 = 下面的框）
      for (let attempt = 0; attempt < 2; attempt++) {
        log(`  尝试点击第 ${attempt + 1} 次`);
        const clickResult = await targetFrame.evaluate((targetText) => {
          const candidates = [];
          document.querySelectorAll('*').forEach(el => {
            if (el.children.length === 0) {
              const text = (el.textContent || '').trim();
              if (text === targetText || text.startsWith(targetText + '↓') || text.startsWith(targetText + '↑') || text.startsWith(targetText + '↕')) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) {
                  candidates.push({ el, x: rect.x, y: rect.y, text });
                }
              }
            }
          });
          if (candidates.length === 0) return { ok: false, reason: 'no-candidate' };
          candidates.sort((a, b) => b.y - a.y);
          const target = candidates[0];
          target.el.click();
          return { ok: true, x: target.x, y: target.y, text: target.text, total: candidates.length };
        }, '最新成交');

        log(`  点击结果: ${JSON.stringify(clickResult)}`);
        // 不 break！需要点两次
        await sleep(1000);
      }
      await sleep(2000);
    }

    // ===== 5b. 对上面框（关注列表）也按"最新成交"排序 =====
    // 用户确认：双击上框的"最新成交"后，上框也会按最新成交降序排列，显示同样的 10 只 ≥2.0 的债
    // 之前只排了下框、没排上框 → 上框停在默认顺序 → findUpper 匹配不到 3 只
    if (sortResult.length >= 2) {
      const sortedByY = [...sortResult].sort((a, b) => a.y - b.y);
      const upperHeader = sortedByY[0]; // y 最小 = 上面框
      log(`  [上面框] 排序目标: ${upperHeader.tag} at (${upperHeader.x},${upperHeader.y}) ${upperHeader.w}w "${upperHeader.text}"`);
      // 用与下框相同的方式（两次 .click()）+ Playwright 坐标点击兜底
      // 先尝试 evaluate 内 click（与下框一致的方式）
      const upperClickResult = await targetFrame.evaluate((targetText) => {
        const candidates = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0) {
            const text = (el.textContent || '').trim();
            if (text === targetText || text.startsWith(targetText + '↓') || text.startsWith(targetText + '↑') || text.startsWith(targetText + '↕')) {
              const rect = el.getBoundingClientRect();
              if (rect.width > 0 && rect.height > 0) candidates.push({ el, x: rect.x, y: rect.y, text });
            }
          }
        });
        if (candidates.length === 0) return { ok: false, reason: 'no-candidate' };
        candidates.sort((a, b) => a.y - b.y); // y 最小优先 = 上面框
        const target = candidates[0];
        target.el.click();
        return { ok: true, x: target.x, y: target.y, text: target.text };
      }, '最新成交');
      log(`  上框第1次点击: ${JSON.stringify(upperClickResult)}`);
      await sleep(1200);
      // 第2次点击 → 降序
      await targetFrame.evaluate((targetText) => {
        const candidates = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0) {
            const text = (el.textContent || '').trim();
            if (text === targetText || text.startsWith(targetText + '↓') || text.startsWith(targetText + '↑') || text.startsWith(targetText + '↕')) {
              const rect = el.getBoundingClientRect();
              if (rect.width > 0 && rect.height > 0) candidates.push({ el, x: rect.x, y: rect.y, text });
            }
          }
        });
        if (candidates.length === 0) return false;
        candidates.sort((a, b) => a.y - b.y);
        candidates[0].el.click();
        return true;
      }, '最新成交');
      log(`  已对上框"最新成交"列头点击2次（降序）`);
      await sleep(2500); // 等待排序完成
    } else {
      log(`  ⚠️ 只找到 ${sortResult.length} 个"最新成交"表头，无法单独排上框`);
    }

    await screenshot(page, '06-sorted');

    // ===== 7a. dump 下面框的表头列名 =====
    log('7a. dump 下面框表头列名 (y=' + headerY + ')');
    const lowerHeaders = await targetFrame.evaluate((hy) => {
      const result = [];
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0) {
          const text = (el.textContent || '').trim();
          const rect = el.getBoundingClientRect();
          if (text && text.length > 0 && text.length < 20 && rect.width > 20 && rect.height > 0 &&
              rect.y >= hy - 5 && rect.y < hy + 35 && rect.x > 150) {
            result.push({ tag: el.tagName, text: text.substring(0, 20), x: Math.round(rect.x), y: Math.round(rect.y), w: Math.round(rect.width) });
          }
        }
      });
      return result.sort((a, b) => a.x - b.x);
    }, headerY);
    log(`  下面框表头候选 ${lowerHeaders.length} 个:`);
    for (const h of lowerHeaders) {
      log(`    ${h.tag} "${h.text}" at x=${h.x} y=${h.y} w=${h.w}`);
    }

    // ===== 7b. 横向滚动收集所有列 + 提取行数据 =====
    log('7b. 横向滚动收集所有列 + 提取行数据');

    // 快照函数：返回当前视口内 下面框 的 表头(headers) 和 数据cell(cells)，以及横向滚动状态
    async function snapshotLower(hy) {
      return await targetFrame.evaluate((hy) => {
        // 动态定位下面框表头 y（取含"最新成交"的最大 y）
        let curHeaderY = 0;
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.y > curHeaderY) curHeaderY = r.y;
          }
        });
        // 表头叶子节点（y ≈ curHeaderY）
        const headers = [];
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length !== 0) return;
          const t = (el.textContent || '').replace(/[↓↑↕\s]/g, '').trim();
          const r = el.getBoundingClientRect();
          if (t && t.length > 0 && t.length < 20 && r.width > 15 &&
              Math.abs(r.y - curHeaderY) <= 6 && r.x > 140 && r.x < 1920) {
            headers.push({ name: t, x: Math.round(r.x), w: Math.round(r.width) });
          }
        });
        headers.sort((a, b) => a.x - b.x);
        // 数据叶子节点（y > 表头下方）
        const dcells = [];
        document.querySelectorAll('div').forEach(el => {
          const t = (el.textContent || '').trim();
          if (!t || t.length > 60) return;
          // 过滤页面 UI 标签文字（Tab 名、导航文字等，不是数据）
          if (/^(成交行情|我的关注|关注列表|利率债|一级发行|信用债|二级|市场观点|行情|关注|发行|市场)$/.test(t)) return;
          if (t === '权' || t === '免') return;
          const r = el.getBoundingClientRect();
          if (r.x < 140 || r.x > 1920) return;
          if (r.y <= curHeaderY + 5) return;
          if (r.width < 8 || r.height < 5 || r.height > 50) return;
          dcells.push({ x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), text: t });
        });
        // 按 y 分行
        const yKeys = [];
        for (const c of dcells) if (!yKeys.find(k => Math.abs(k - c.y) <= 8)) yKeys.push(c.y);
        const rows = [];
        for (const by of yKeys) {
          const rc = dcells.filter(c => Math.abs(c.y - by) <= 8).sort((a, b) => a.x - b.x);
          if (rc.length < 2) continue;
          // 本快照内按当前表头 x 对齐（阈值 50px）
          const vals = {};
          for (const c of rc) {
            let md = Infinity, mi = -1;
            for (let i = 0; i < headers.length; i++) {
              const d = Math.abs(c.x - headers[i].x);
              if (d < md) { md = d; mi = i; }
            }
            if (mi >= 0 && md < 50 && !vals[headers[mi].name]) vals[headers[mi].name] = c.text;
          }
          rows.push({ y: Math.round(by), vals });
        }
        // 横向滚动容器状态
        let headerEl = null;
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && Math.abs(r.y - curHeaderY) <= 15) headerEl = el;
          }
        });
        let scrollLeft = 0, maxScroll = 0;
        if (headerEl) {
          let el = headerEl.parentElement;
          while (el && el.parentElement) {
            if (el.scrollWidth > el.clientWidth + 30) {
              scrollLeft = el.scrollLeft; maxScroll = el.scrollWidth - el.clientWidth; break;
            }
            el = el.parentElement;
          }
        }
        return { headers, rows, scrollLeft: Math.round(scrollLeft), maxScroll: Math.round(maxScroll) };
      }, hy);
    }

    // 横向滚动函数：把下面框容器向右滚 px 像素
    async function scrollLowerRight(hy, px) {
      return await targetFrame.evaluate(({ hy, px }) => {
        let headerEl = null;
        document.querySelectorAll('*').forEach(el => {
          if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && Math.abs(r.y - hy) <= 15) headerEl = el;
          }
        });
        if (!headerEl) return { ok: false };
        let el = headerEl.parentElement;
        while (el && el.parentElement) {
          if (el.scrollWidth > el.clientWidth + 30) {
            const old = el.scrollLeft; el.scrollLeft += px;
            return { ok: el.scrollLeft !== old, sl: Math.round(el.scrollLeft) };
          }
          el = el.parentElement;
        }
        return { ok: false };
      }, { hy, px });
    }

    // 加宽视口：让下面框所有列一次性渲染（避免自定义横向滚动条问题）
    try {
      await page.setViewportSize({ width: 3400, height: 1080 });
      await sleep(1500);
      log('  视口已加宽到 3400，等待重排');
    } catch (e) { log(`  加宽视口失败: ${e.message}`); }

    // 单次快照：收集全宽叶子节点，用数据 cell 的 x 聚类定义列（可捕获无表头右侧列）
    const extract = await targetFrame.evaluate(() => {
      // 定位下面框表头 y（含"最新成交"的最大 y）
      let curHeaderY = 0;
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && r.y > curHeaderY) curHeaderY = r.y;
        }
      });
      const heads = [];   // 表头叶子节点
      const dcells = [];  // 数据叶子节点
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length !== 0) return;
        const raw = (el.textContent || '').trim();
        if (!raw || raw.length > 60) return;
        const r = el.getBoundingClientRect();
        if (r.x < 140 || r.x > 3380) return;
        if (r.width < 8 || r.height < 5 || r.height > 50) return;
        if (Math.abs(r.y - curHeaderY) <= 6) {
          const name = raw.replace(/[↓↑↕\s]/g, '');
          if (name && name.length < 20) heads.push({ x: Math.round(r.x), w: Math.round(r.width), name });
        } else if (r.y > curHeaderY + 6 && r.y < curHeaderY + 900) {
          // 过滤页面 UI 标签文字（Tab 名、导航文字等）
          if (/^(成交行情|我的关注|关注列表|利率债|一级发行|信用债|二级|市场观点|行情|关注|发行|市场)$/.test(raw)) return;
          if (raw === '权' || raw === '免') return;
          dcells.push({ x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), text: raw });
        }
      });
      heads.sort((a, b) => a.x - b.x);

      // 按 y 分行
      const yKeys = [];
      for (const c of dcells) if (!yKeys.find(k => Math.abs(k - c.y) <= 8)) yKeys.push(c.y);
      yKeys.sort((a, b) => a - b);
      const rawRows = yKeys.map(by => dcells.filter(c => Math.abs(c.y - by) <= 8).sort((a, b) => a.x - b.x));

      // 用所有数据 cell 的 x 聚类定义列（不依赖表头文字）
      const allX = [];
      for (const row of rawRows) for (const c of row) allX.push(c.x);
      allX.sort((a, b) => a - b);
      const centers = [];
      for (const x of allX) {
        const last = centers[centers.length - 1];
        if (last && x - last.max <= 30) { last.sum += x; last.n++; last.max = x; }
        else centers.push({ sum: x, n: 1, max: x });
      }
      // 只保留出现足够多的列（过滤展开子行/噪声）
      const minCount = Math.max(3, Math.floor(rawRows.length * 0.3));
      const cols = centers.filter(c => c.n >= minCount).map(c => ({ x: Math.round(c.sum / c.n) }));
      // 给每列配表头名（最近表头 x，阈值 45px），否则 列N
      cols.forEach((col, i) => {
        let md = Infinity, nm = '';
        for (const h of heads) { const d = Math.abs(h.x - col.x); if (d < md) { md = d; nm = h.name; } }
        col.name = (md < 45 && nm) ? nm : `列${i + 1}`;
      });

      // 每行 cell 分配到最近列（阈值 30px）；债券简称列多选优，避开 休/纯数字/纯代码
      const nameColIdx = cols.findIndex(c => c.name.includes('债券简称'));
      const outRows = [];
      for (let ri = 0; ri < rawRows.length; ri++) {
        const colCells = {};
        for (const c of rawRows[ri]) {
          let md = Infinity, mi = -1;
          for (let i = 0; i < cols.length; i++) { const d = Math.abs(c.x - cols[i].x); if (d < md) { md = d; mi = i; } }
          if (mi >= 0 && md < 30) { (colCells[mi] || (colCells[mi] = [])).push(c.text); }
        }
        const vals = new Array(cols.length).fill('');
        for (const mi in colCells) {
          const cands = colCells[mi];
          if (Number(mi) === nameColIdx) {
            const good = cands.filter(t => { t = (t || '').trim(); return t && t !== '休' && !/^\d+$/.test(t) && !/^\d+[YD]/.test(t) && t.length >= 2 && !/^[A-Za-z0-9]+$/.test(t); });
            const pick = (good.length ? good : cands).sort((a, b) => (b || '').length - (a || '').length)[0];
            if (pick) vals[mi] = pick;
          } else {
            vals[mi] = cands[0]; // 非债券简称列：取最左（first-wins）
          }
        }
        outRows.push({ y: yKeys[ri], vals });
      }
      return { headerY: curHeaderY, cols: cols.map(c => c.name), colX: cols.map(c => c.x), rows: outRows };
    });

    log(`  下面框表头 y=${extract.headerY}，识别 ${extract.cols.length} 列: ${extract.cols.join(' | ')}`);
    log(`  列 x 坐标: ${extract.colX.join(', ')}`);

    const colDefs = extract.cols.map(name => ({ name }));
    const nameIdx = colDefs.findIndex(c => c.name.includes('债券简称'));
    const latestColIdx = colDefs.findIndex(c => c.name.includes('最新成交'));
    const avgColIdx = colDefs.findIndex(c => c.name.includes('平均成交'));

    // 组装 allRowsMap（cells 对齐 colDefs 顺序）
    const allRowsMap = new Map();
    for (const row of extract.rows) {
      const bondName = nameIdx >= 0 ? (row.vals[nameIdx] || '').trim() : '';
      if (!bondName) continue; // 无债券简称 = 表头/空行
      const cells = row.vals.slice();
      let maxVal = 0;
      if (latestColIdx >= 0 && cells[latestColIdx]) {
        const rv = cells[latestColIdx].trim();
        if (/^[\d.]+$/.test(rv)) {
          maxVal = parseFloat(rv);
        } else if (/^休/.test(rv) && avgColIdx >= 0 && cells[avgColIdx]) {
          const av = parseFloat(cells[avgColIdx]);
          if (!isNaN(av) && av > 0) { maxVal = av; cells[latestColIdx] = cells[avgColIdx]; }
        }
      }
      allRowsMap.set(bondName, { cells, maxVal, bondName, y: row.y });
    }

    // 转为数组
    const rows = Array.from(allRowsMap.values());
    rows.sort((a, b) => b.maxVal - a.maxVal); // 按最新成交降序
    log(`  总共提取到 ${rows.length} 行（去重后）`);
    // 调试：逐行输出完整 13 列（标注列名），重点看 最新成交 及后列
    for (const r of rows) {
      const parts = colDefs.map((c, i) => `${c.name}=${r.cells[i] || ''}`);
      log(`  [${r.maxVal}] ${r.bondName}`);
      log(`      ${parts.join('  ')}`);
    }

    // ===== 7c. 提取上面框（y 值较小，表头 y≈71）全部数据 → Sheet2 =====
    log('7c. 提取上面框（全部关注）数据');
    let upperColDefs = [];
    let upperRows = [];
    // 定位上面框表头 y（含"最新成交"的最小 y）+ 标记其滚动容器
    const upperMeta = await targetFrame.evaluate(() => {
      let ys = [];
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
          const r = el.getBoundingClientRect();
          if (r.width > 0) ys.push(r.y);
        }
      });
      if (!ys.length) return null;
      ys.sort((a, b) => a - b);
      const hy = ys[0]; // 上面框表头（最小 y）
      // 找到上面框滚动容器并标记
      let headerEl = null;
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length === 0 && (el.textContent || '').trim().includes('最新成交')) {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && Math.abs(r.y - hy) <= 2) headerEl = el;
        }
      });
      let container = null;
      if (headerEl) {
        let el = headerEl.parentElement;
        while (el && el.parentElement) {
          // 真正的滚动体：可滚动且自身高度合理（排除 20px 表头行容器和页面级根节点）
          if (el.scrollHeight > el.clientHeight + 60 && el.clientHeight > 60) { container = el; break; }
          el = el.parentElement;
        }
        // 找不到合适容器时保持 null —— 不要兜底到表头父节点（会误标 20px 小节点导致 0 行）
      }
      if (container) container.__upperBox = true;
      // 收集表头（全宽，x<3380，按 y≈hy）
      const heads = [];
      document.querySelectorAll('*').forEach(el => {
        if (el.children.length !== 0) return;
        const raw = (el.textContent || '').trim();
        if (!raw || raw.length > 20) return;
        const r = el.getBoundingClientRect();
        if (r.width < 8 || r.height < 5 || r.height > 50) return;
        if (Math.abs(r.y - hy) <= 6 && r.x > 140 && r.x < 3380) {
          const name = raw.replace(/[↓↑↕\s]/g, '');
          if (name) heads.push({ x: Math.round(r.x), w: Math.round(r.width), name });
        }
      });
      heads.sort((a, b) => a.x - b.x);
      const rect = container ? container.getBoundingClientRect() : null;
      return {
        headerY: hy,
        headers: heads,
        contTop: rect ? Math.round(rect.top) : hy,
        contBottom: rect ? Math.round(rect.bottom) : hy + 800,
        scrollTop: container ? container.scrollTop : 0,
        scrollHeight: container ? container.scrollHeight : 0,
        clientHeight: container ? container.clientHeight : 0
      };
    });
    if (!upperMeta) {
      log('  ⚠️ 未找到上面框表头，跳过');
    } else {
      upperMeta.lowerHeaderY = extract.headerY; // 下框表头 y 作为上框数据区下边界（用 y 区间分隔两框）
      upperMeta.contTop = upperMeta.headerY;     // 容器可见区用于滚动兜底，数据行判断改用 y 区间
      log(`  上面框表头 y=${upperMeta.headerY}，原始表头 ${upperMeta.headers.length} 个: ${upperMeta.headers.map(h => h.name).join(' | ')}`);
      log(`  上面框表头 x: ${upperMeta.headers.map(h => `${h.name}(${h.x})`).join(' | ')}`);
      log(`  上面框滚动容器: scrollTop=${upperMeta.scrollTop} scrollHeight=${upperMeta.scrollHeight} clientHeight=${upperMeta.clientHeight}`);

      // 快照函数：收集当前视口内上面框数据行（用数据 cell x 聚类定义列，配表头名）
      async function snapshotUpper() {
        return await targetFrame.evaluate((meta) => {
          const heads = meta.headers;
          const hy = meta.headerY, lowerHeaderY = meta.lowerHeaderY;
          const dcells = [];
          document.querySelectorAll('*').forEach(el => {
            if (el.children.length !== 0) return;
            const raw = (el.textContent || '').trim();
            if (!raw || raw.length > 60) return;
            // 过滤页面 UI 标签文字（Tab 名、导航文字等）
            if (/^(成交行情|我的关注|关注列表|利率债|一级发行|信用债|二级|市场观点|行情|关注|发行|市场)$/.test(raw)) return;
            if (raw === '权' || raw === '免') return;
            const r = el.getBoundingClientRect();
            if (r.x < 140 || r.x > 3380) return;
            if (r.width < 8 || r.height < 5 || r.height > 50) return;
            const y = Math.round(r.y);
            if (y <= hy + 6) return;                       // 排除上框表头行
            if (y >= lowerHeaderY - 8) return;            // 排除下框及下方区域（用表头 y 区间分隔两框）
            dcells.push({ x: Math.round(r.x), y, w: Math.round(r.width), text: raw });
          });
          // 按 y 分行
          const yKeys = [];
          for (const c of dcells) if (!yKeys.find(k => Math.abs(k - c.y) <= 8)) yKeys.push(c.y);
          yKeys.sort((a, b) => a - b);
          const rawRows = yKeys.map(by => dcells.filter(c => Math.abs(c.y - by) <= 8).sort((a, b) => a.x - b.x));
          // 用表头 x 强制初始化列 centers（确保每个表头列都有锚点，不会因数据稀疏丢失/串列）
          const seedCenters = heads.map(h => ({ sum: h.x, n: 1, max: h.x, _seed: true }));
          // 再加入数据 cell 的 x 做增量合并（gap<=30 合并到最近 center）
          const allX = dcells.map(c => c.x).sort((a, b) => a - b);
          const centers = [...seedCenters];
          for (const x of allX) {
            const last = centers[centers.length - 1];
            if (last && x - last.max <= 30) { last.sum += x; last.n++; last.max = x; }
            else centers.push({ sum: x, n: 1, max: x, _seed: false });
          }
          // 保留所有表头种子列 + 数据足够密集的非种子列
          const dataMinCount = Math.max(2, Math.floor(rawRows.length * 0.25));
          const cols = centers.filter(c => c._seed || c.n >= dataMinCount).map(c => ({ x: Math.round(c.sum / c.n) }));
          cols.forEach((col, i) => {
            let md = Infinity, nm = '';
            for (const h of heads) { const d = Math.abs(h.x - col.x); if (d < md) { md = d; nm = h.name; } }
            col.name = (md < 45 && nm) ? nm : `列${i + 1}`;
          });
          // 每行 cell 分配到最近列（阈值 30px）；债券简称列多选优，避开 休/纯数字/纯代码
          const out = [];
          for (const row of rawRows) {
            const colCells = {};
            for (const c of row) {
              let md = Infinity, mi = -1;
              for (let i = 0; i < cols.length; i++) { const d = Math.abs(c.x - cols[i].x); if (d < md) { md = d; mi = i; } }
              if (mi >= 0 && md < 30) {
                const cn = cols[mi].name;
                (colCells[cn] || (colCells[cn] = [])).push({ text: c.text, x: c.x });
              }
            }
            const vals = {};
            for (const cn in colCells) {
              const cands = colCells[cn];
              if (cn === '债券简称') {
                // 优选像真实债券简称的：非 休、非纯数字、非"数字+Y/D"、非纯代码（如 22YN01EB），长度≥2
                const good = cands.filter(c => {
                  const t = (c.text || '').trim();
                  return t && t !== '休' && !/^\d+$/.test(t) && !/^\d+[YD]/.test(t)
                         && t.length >= 2 && !/^[A-Za-z0-9]+$/.test(t);
                });
                const pick = (good.length ? good : cands).sort((a, b) => (b.text || '').length - (a.text || '').length)[0];
                if (pick) vals[cn] = pick.text;
              } else {
                // 其他列：取最左（first-wins）
                vals[cn] = cands.sort((a, b) => a.x - b.x)[0].text;
              }
            }
            if (Object.keys(vals).length >= 2) out.push(vals);
          }
          return out;
        }, upperMeta);
      }

      // 滚动函数：优先原生 scrollTop，失败用鼠标滚轮
      async function scrollUpperDown(px) {
        const moved = await targetFrame.evaluate((px) => {
          let cont = null;
          document.querySelectorAll('*').forEach(el => { if (el.__upperBox) cont = el; });
          if (!cont) return false;
          const old = cont.scrollTop; cont.scrollTop += px;
          return cont.scrollTop !== old;
        }, px);
        if (!moved) {
          const my = Math.round((upperMeta.headerY + upperMeta.lowerHeaderY) / 2);
          await page.mouse.move(600, my);
          await page.mouse.wheel(0, px);
        }
        return moved;
      }

      // ===== 辅助：判断异常简称 / 生成归一化对比 key =====
      const isAbnormalName = (s) => {
        s = (s || '').trim();
        return !s || s === '休' || /^\d+$/.test(s) || /^\d+[YD]/.test(s) || s.length < 2;
      };
      // 归一化 key：优先债券代码（唯一ID），其次债券简称
      const normBondKey = (vals) => {
        const c = (vals['债券代码'] || '').trim();
        const n = (vals['债券简称'] || '').trim();
        if (c && c !== '休' && c.length >= 2) return 'CODE:' + c;
        if (n && n !== '休' && n.length >= 2 && !/^\d+[YD]/.test(n)) return 'NAME:' + n;
        return null;
      };

      // 主循环：纵向滚动收集所有行（按 债券代码 优先、债券简称 兜底的稳定 key 去重，确保不遗漏）
      const upperMap = new Map();
      let stall = 0, iter = 0;
      while (stall < 3 && iter < 80) {
        iter++;
        const snap = await snapshotUpper();
        let added = 0;
        for (const vals of snap) {
          const c0 = (vals['债券代码'] || '').trim();
          const n0 = (vals['债券简称'] || '').trim();
          // 决定稳定 key：优先债券代码（唯一），代码异常/空时退回债券简称
          let key = null;
          if (c0 && c0 !== '休' && c0.length >= 2) key = 'CODE:' + c0;
          else if (n0 && n0 !== '休' && !/^\d+[YD]/.test(n0) && n0.length >= 2) key = 'NAME:' + n0;
          if (!key) continue; // 代码和简称都不可用，跳过（避免把"休"/数字当 key 导致串列）
          if (!upperMap.has(key)) {
            upperMap.set(key, vals); added++;
          } else {
            // 合并：若已存条目的债券简称异常、当前更完整，则替换（保留真实名称）
            const prevName = (upperMap.get(key)['债券简称'] || '').trim();
            const curName = n0;
            if (isAbnormalName(prevName) && !isAbnormalName(curName)) upperMap.set(key, vals);
          }
        }
        log(`  [上面框快照${iter}] 本帧 ${snap.length} 行, 新增 ${added}, 累计 ${upperMap.size}`);
        if (added === 0) stall++; else stall = 0;
        await scrollUpperDown(700);
        await sleep(500);
      }

      // 后处理：把"仅以代码为 key"的条目，若能与某"名称条目"通过债券代码匹配，则合并（去重，保留名称条目的完整简称）
      {
        const codeToNameKey = new Map();
        for (const [k, v] of upperMap) if (k.startsWith('CODE:')) { const c = (v['债券代码'] || '').trim(); if (c) codeToNameKey.set(c, k); }
        for (const [k, v] of upperMap) {
          if (k.startsWith('NAME:')) {
            const c = (v['债券代码'] || '').trim();
            if (c && codeToNameKey.has(c)) { upperMap.delete(codeToNameKey.get(c)); codeToNameKey.delete(c); }
          }
        }
      }

      // 诊断：检查上框是否有异常简称 / 仅代码条目（名称缺失）
      let abnormalCount = 0, codeOnlyCount = 0;
      const abnormalList = [], codeOnlyList = [];
      for (const [key, vals] of upperMap) {
        const name = (vals['债券简称'] || '').trim();
        const code = (vals['债券代码'] || '').trim();
        const latest = (vals['最新成交'] || '').trim();
        if (isAbnormalName(name)) {
          abnormalCount++;
          abnormalList.push(`key="${key}" 简称="${name}" 代码="${code}" 剩余期限="${(vals['剩余期限'] || '').trim()}" 最新成交="${latest}"`);
        } else if (key.startsWith('CODE:') && !name) {
          codeOnlyCount++;
          codeOnlyList.push(`代码="${code}" 最新成交="${latest}"`);
        }
      }
      if (abnormalCount > 0) { log(`  ⚠️ 上框发现 ${abnormalCount} 个异常简称:`); for (const a of abnormalList) log(`    ${a}`); }
      if (codeOnlyCount > 0) { log(`  ℹ️ 上框有 ${codeOnlyCount} 个仅含代码的条目（债券简称未能捕获，已按代码保留，建议人工核对）:`); for (const a of codeOnlyList) log(`    ${a}`); }
      if (abnormalCount === 0 && codeOnlyCount === 0) log(`  ✅ 上框债券简称全部正常捕获，无异常/遗漏`);

      // 对比上下框一致性（归一化 key：债券代码优先，否则债券简称）
      const lowerNorm = new Map();
      for (const r of rows) {
        const lv = {};
        colDefs.forEach((c, i) => { lv[c.name] = r.cells[i] || ''; });
        const k = normBondKey(lv);
        if (k) lowerNorm.set(k, r.bondName);
      }
      const upperNorm = new Map();
      for (const vals of upperMap.values()) {
        const k = normBondKey(vals);
        if (k && (vals['最新成交'] || '').trim()) upperNorm.set(k, (vals['债券简称'] || '').trim() || (vals['债券代码'] || '').trim());
      }
      const onlyUpper = [...upperNorm.keys()].filter(k => !lowerNorm.has(k));
      const onlyLower = [...lowerNorm.keys()].filter(k => !upperNorm.has(k));
      if (onlyUpper.length === 0 && onlyLower.length === 0) {
        log(`  ✅ 上下框债券集合一致（均有最新成交）: 上框 ${upperNorm.size} = 下框 ${lowerNorm.size}`);
      } else {
        log(`  📊 上下框债券差异: 仅上框有 ${onlyUpper.length} 个, 仅下框有 ${onlyLower.length} 个`);
        const showU = onlyUpper.slice(0, 30).map(k => upperNorm.get(k) || k);
        const showL = onlyLower.slice(0, 30).map(k => lowerNorm.get(k) || k);
        if (showU.length) log(`    仅上框: ${showU.join(', ')}`);
        if (showL.length) log(`    仅下框: ${showL.join(', ')}`);
      }

      // 审计：导出上框全量 + 下框全量（代码/简称/最新成交）到 frame_audit.txt 供核对"是否遗漏"
      try {
        const audit = { upper: [], lower: [] };
        for (const [key, vals] of upperMap) {
          audit.upper.push({ key, code: (vals['债券代码'] || '').trim(), name: (vals['债券简称'] || '').trim(), latest: (vals['最新成交'] || '').trim(), rest: (vals['剩余期限'] || '').trim() });
        }
        for (const r of rows) {
          const lv = {}; colDefs.forEach((c, i) => { lv[c.name] = r.cells[i] || ''; });
          audit.lower.push({ code: (lv['债券代码'] || '').trim(), name: r.bondName, latest: (lv['最新成交'] || '').trim() });
        }
        fs.writeFileSync(path.join(WORKSPACE, 'frame_audit.txt'), JSON.stringify(audit, null, 1), 'utf8');
        log(`  审计文件已写出: frame_audit.txt (上框 ${audit.upper.length} / 下框 ${audit.lower.length})`);
      } catch (e) { log(`  审计写出失败: ${e.message}`); }

      // 构建 upperColDefs + upperRows（按表头顺序，补齐缺失列）
      const upperColNames = upperMeta.headers.map(h => h.name);
      const seen = new Set(upperColNames);
      const extraNames = [];
      for (const vals of upperMap.values()) {
        for (const k of Object.keys(vals)) if (!seen.has(k)) { seen.add(k); extraNames.push(k); }
      }
      upperColDefs = upperColNames.concat(extraNames).map(n => ({ name: n }));
      upperRows = [];
      for (const vals of upperMap.values()) {
        const cells = upperColDefs.map(c => vals[c.name] || '');
        let maxVal = 0;
        const lv = (vals['最新成交'] || '').trim();
        if (/^[\d.]+$/.test(lv)) maxVal = parseFloat(lv);
        else if (/^休/.test(lv)) {
          const av = parseFloat(vals['平均成交']);
          if (!isNaN(av) && av > 0) maxVal = av;
        }
        upperRows.push({ cells, maxVal, bondName: (vals['债券简称'] || '').trim() || (vals['债券代码'] || '').trim() });
      }
      upperRows.sort((a, b) => b.maxVal - a.maxVal);
      log(`  上面框总共提取 ${upperRows.length} 行`);
    }

    // 筛选 ≥ MIN_VOLUME
    const filtered = rows.filter(r => r.maxVal >= MIN_VOLUME);
    log(`  筛选 ≥${MIN_VOLUME}：${filtered.length} 行`);

    // 上框"休市"/解析失败行 → 回退用下框同债券的最新成交值（保证上下框一致，均为≥2）
    // 下框(成交行情)含 开盘/最高/最低/平均成交，最新成交解析可靠；上框无平均成交列时以此补全
    const lowerLatestByKey = new Map();
    for (const r of filtered) {
      const c = (r.cells[1] || '').trim();           // 债券代码列(index 1)
      const n = (r.bondName || '').trim();
      if (c) lowerLatestByKey.set('CODE:' + c, r.maxVal);
      if (n) lowerLatestByKey.set('NAME:' + n, r.maxVal);
    }
    let boosted = 0;
    for (const ur of upperRows) {
      if (ur.maxVal < MIN_VOLUME) {
        const c = (ur.cells[1] || '').trim();
        const n = (ur.bondName || '').trim();
        const fb = lowerLatestByKey.get('CODE:' + c) || lowerLatestByKey.get('NAME:' + n);
        if (fb && fb >= MIN_VOLUME) { ur.maxVal = fb; boosted++; }
      }
    }
    if (boosted > 0) log(`  ↳ 上框 ${boosted} 行回退用下框值补足（最新成交解析失败）`);

    // ===== 8. 保存为 Excel =====
    log('8. 保存为 Excel');
    // 清洗每个 cell：去除连续重复 token
    function cleanCell(text) {
      if (!text) return text;
      const s = String(text).trim();
      const tokens = s.split(/\s+/);
      if (tokens.length === 0) return s;
      const result = [tokens[0]];
      for (let i = 1; i < tokens.length; i++) {
        if (tokens[i] !== result[result.length - 1]) {
          result.push(tokens[i]);
        }
      }
      return result.join(' ');
    }

    const wb = new ExcelJS.Workbook();
    // 上框按 代码/简称 查找（下框债券回填上框专属列；上框"我的关注"实时性不稳定，以下框为准）
    const upperByKey = new Map();
    for (const r of upperRows) {
      const c = (r.cells[1] || '').trim();
      const n = (r.bondName || '').trim();
      if (c) upperByKey.set('CODE:' + c, r);
      if (n) upperByKey.set('NAME:' + n, r);
    }
    const findUpper = (code, name) => upperByKey.get('CODE:' + code) || upperByKey.get('NAME:' + name) || null;
    // 修正下框表头：数据列按 x 聚类命名时可能把"中债偏离(BP)"误标为"更新时间"
    // 用内容判定：标"更新时间"但其值多为数值（非 HH:MM:SS）→ 实为"中债偏离(BP)"
    {
      const isTimeStr = (v) => typeof v === 'string' && /^(\d{1,2}):\d{2}(:\d{2})?$/.test(v.trim());
      for (let i = 0; i < colDefs.length; i++) {
        if (colDefs[i].name !== '更新时间') continue;
        let times = 0, decs = 0;
        for (const rr of rows) {
          const v = (rr.cells[i] || '').trim();
          if (!v) continue;
          if (isTimeStr(v)) times++;
          else if (/^-?\d+(\.\d+)?$/.test(v)) decs++;
        }
        if (decs > times) {
          colDefs[i].name = '中债偏离(BP)';
          log(`  ⚠️ 下框列${i + 1} 原标"更新时间"但内容为数值，修正为"中债偏离(BP)"`);
        }
      }
    }
    // 上框/下框列名→index 映射（供 getUpperVal/getLowerVal 使用，须在 Sheet 写入前初始化）
    const uIdx = {};
    for (let i = 0; i < upperColDefs.length; i++) uIdx[upperColDefs[i].name] = i;
    const lIdx = {};
    for (let i = 0; i < colDefs.length; i++) lIdx[colDefs[i].name] = i;

    // ---- 逐只点击债券，从详情面板抓取"区域" ----
    // 主方案：发行人简称 → 区域映射表（稳定可靠）
    // 辅助方案：双击债券行从页面提取（可能因DOM变化失败）
    log('  填充区域信息...');
    const ISSUER_REGION_MAP = {
      // 重庆
      '涪陵国投': '重庆', '万州经开': '重庆', '重庆市': '重庆', '涪陵': '重庆',
      // 云南
      '云南能投': '云南', '云能投': '云南', '云南能投集团': '云南',
      // 湖北
      '联投集团': '湖北', '湖北联投': '湖北', '武汉城建': '湖北',
      // 湖南
      '湖南银行': '湖南', '华融湘江': '湖南',
      // 浙江
      '浙商银行': '浙江',
      // 山东
      '恒丰银行': '山东',
      // 四川
      '鸿飞集团': '四川', '达州投资': '四川', '达州': '四川',
      // 全国性银行（无特定省份，留空）
      '农业银行': '', '农行': '', '中国银行': '', '建设银行': '',
      '工商银行': '', '交通银行': '', '邮储银行': '',
      '招商银行': '', '中信银行': '', '浦发银行': '',
      '民生银行': '', '兴业银行': '', '光大银行': '',
      '平安银行': '', '华夏银行': '', '渤海银行': '',
    };

    function inferRegion(issuerName) {
      if (!issuerName) return '';
      // 精确匹配
      if (ISSUER_REGION_MAP.hasOwnProperty(issuerName)) return ISSUER_REGION_MAP[issuerName];
      // 模糊匹配：发行人名包含映射key
      for (const [key, region] of Object.entries(ISSUER_REGION_MAP)) {
        if (issuerName.includes(key) || key.includes(issuerName)) return region;
      }
      return '';
    }

    const bondRegionMap = new Map();

    // 银行判定：映射表中值为 '' 的发行人视为全国性银行（区域留空，不截图）
    const BANK_NAMES = Object.keys(ISSUER_REGION_MAP).filter(k => ISSUER_REGION_MAP[k] === '');
    function isBank(issuer) {
      if (!issuer) return false;
      return BANK_NAMES.some(b => issuer === b || issuer.includes(b) || b.includes(issuer));
    }

    // 先用发行人映射预填区域（已知省份 / 已知银行留空），未知发行人留待详情页提取
    for (const r of filtered) {
      const bCode = getLowerVal(r, '债券代码');
      const issuer = getLowerVal(r, '发行人简称');
      const bondName = getLowerVal(r, '债券简称');
      const region = inferRegion(issuer);
      if (region) {
        bondRegionMap.set(bCode, { region, bondName, method: '映射' });
      } else if (isBank(issuer)) {
        bondRegionMap.set(bCode, { region: '', bondName, method: '映射-银行' });
      }
      // 其余（region 为空且非银行）不写入 → 进入详情页提取补充
    }

    // 计算 quote-web iframe 在页面中的偏移（用于裁剪截图坐标换算）
    const iframeOffset = await page.evaluate(() => {
      const iframes = Array.from(document.querySelectorAll('iframe'));
      for (const f of iframes) {
        if (f.src && f.src.includes('quote-web')) {
          const r = f.getBoundingClientRect();
          return { x: r.x, y: r.y };
        }
      }
      return { x: 0, y: 0 };
    }).catch(() => ({ x: 0, y: 0 }));
    log(`  iframe 页面偏移: (${Math.round(iframeOffset.x)}, ${Math.round(iframeOffset.y)})`);

    // ===== 区域自动提取：我的关注双击债券名称进详情页 → DOM提取省份/全称 =====
    // 首只: 列表页dispatch双击债券名称(绕开模态框拦截); 其余: 详情页搜索框(ant-select-search)填代码+Enter切换
    // 省份: 优先"省XX"字段, 兜底"性质 XX市/省"; 全称: 信息栏截断到评级/省/性质前
    const FORCE_REGION = process.env.FORCE_SHOT === '1'; // 复用环境变量名
    const needRegion = [];
    for (const r of filtered) {
      const bCode = getLowerVal(r, '债券代码');
      if (FORCE_REGION || !bondRegionMap.has(bCode)) needRegion.push(r);
    }
    log(`  需从详情页提取区域的债券: ${needRegion.length}/${filtered.length}`);

    if (needRegion.length > 0) {
      log('  进入详情页提取省份与发行人全称...');
      const PROV_LIST = ['北京','天津','上海','重庆','河北','山西','辽宁','吉林','黑龙江','江苏','浙江','安徽','福建','江西','山东','河南','湖北','湖南','广东','海南','四川','贵州','云南','陕西','甘肃','内蒙古','广西','西藏','宁夏','新疆'];

      // 首只：列表页dispatch双击债券名称 → 进详情页（带重试）；其余：详情页搜索框填代码+Enter切换
      // 验证结论(probe系列):
      //  1) 坐标双击/受信dblclick被列表页模态框(dmuiv4-modal-wrap)拦截
      //     须用 element.dispatchEvent(new MouseEvent('dblclick',{bubbles:true})) 触发React合成事件
      //  2) 详情页路由用债券内部ID，hash直跳无效；须由App自身导航(双击/搜索框)触发
      //  3) 省份: 优先"省XX"字段, 兜底"性质 XX市/省"；全称: 信息栏截断到评级/省/性质前
      const isDetailUrl = () => {
        const urls = [page.url(), ...page.frames().map(f=>f.url())].join('|');
        return urls.includes('bond/detail') || urls.includes('/detail/');
      };
      const isDetailContent = (frame) => frame.evaluate(() => /性质|主体评级|债项评级/.test(document.body?.textContent||'')).catch(()=>false);

      // 首只：在列表页dispatch双击债券名称 → 进详情页（带重试）
      async function dblclickIntoDetail(frame, name) {
        const fired = await frame.evaluate((nm) => {
          const els = Array.from(document.querySelectorAll('*')).filter(el => {
            const t = (el.textContent||'').trim();
            return t === nm || (t.startsWith(nm) && t.length <= nm.length + 8);
          });
          if (!els.length) return false;
          els.sort((a,b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
          const target = els.find(el => {
            const r = el.getBoundingClientRect();
            return r.width > 30 && r.width < 400 && r.height > 10 && r.height < 60;
          }) || els[0];
          target.scrollIntoView({ block: 'center' });
          target.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
          return true;
        }, name).catch(() => false);
        if (!fired) return false;
        for (let i=0; i<12; i++) {
          await sleep(500);
          if (isDetailUrl()) break;
        }
        for (let vi=0; vi<8; vi++) {
          await sleep(500);
          if (await isDetailContent(frame)) return true;
        }
        return false;
      }

      // 详情页搜索框：填代码 + Enter 切换债券（停留详情页，无需回列表）
      async function searchSwitch(frame, bCode) {
        const code = String(bCode||'').split('.')[0];
        if (!code) return false;
        const box = await frame.evaluate(() => {
          const input = document.querySelector('input.ant-select-search__field') ||
            Array.from(document.querySelectorAll('input')).find(el => String(el.className||'').includes('ant-select-search'));
          if (!input) return null;
          const r = input.getBoundingClientRect();
          return { x: r.x + r.width/2, y: r.y + r.height/2 };
        }).catch(()=>null);
        if (!box) return false;
        try {
          await frame.click('input.ant-select-search__field', { timeout: 3000 });
        } catch(e) {
          await page.mouse.click(iframeOffset.x + box.x, iframeOffset.y + box.y);
        }
        await sleep(250);
        await page.keyboard.press('Control+A');
        await sleep(120);
        await page.keyboard.press('Backspace');
        await sleep(120);
        await page.keyboard.type(code, { delay: 30 });
        await sleep(900);
        await page.keyboard.press('Enter');
        let switched = false;
        for (let i=0; i<14; i++) {
          await sleep(500);
          if (isDetailUrl()) {
            for (let vi=0; vi<6; vi++) {
              await sleep(400);
              if (await isDetailContent(frame)) { switched = true; break; }
            }
            if (switched) break;
          }
        }
        if (switched) {
          const curVal = await frame.evaluate(() => {
            const input = document.querySelector('input.ant-select-search__field');
            return input ? (input.value||'') : '';
          }).catch(()=> '');
          if (curVal && !curVal.includes(code)) {
            log(`  ⚠️ 搜索框切换后值[${curVal}]未含目标代码[${code}]`);
          }
        }
        return switched;
      }

      // 提取省份 + 发行人全称
      async function extractFrom(frame) {
        return await frame.evaluate((PROV_LIST) => {
          const result = { province:'', city:'', issuerFull:'', method:'' };
          document.querySelectorAll('*').forEach(el => {
            if (result.province) return;
            const t = (el.textContent||'').trim();
            const m = t.match(/^省(.{2,5})$/);
            if (m) {
              let p = m[1];
              if (p.endsWith('省') || p.endsWith('市')) p = p.slice(0, -1);
              if (PROV_LIST.includes(p)) { result.province = p; result.method='省字段'; }
            }
          });
          if (!result.province) {
            document.querySelectorAll('*').forEach(el => {
              if (result.province) return;
              const t = (el.textContent||'').trim();
              if (t.includes('性质')) {
                const m = t.match(/(北京|天津|上海|重庆|河北|山西|辽宁|吉林|黑龙江|江苏|浙江|安徽|福建|江西|山东|河南|湖北|湖南|广东|海南|四川|贵州|云南|陕西|甘肃|内蒙古|广西|西藏|宁夏|新疆)/);
                if (m) { result.province = m[1]; result.method='性质字段'; }
              }
            });
          }
          const cands = Array.from(document.querySelectorAll('*')).filter(el => {
            const t = (el.textContent||'').trim();
            const r = el.getBoundingClientRect();
            return r.y>=40 && r.y<=260 && r.x>=600 && t.length>=4 && t.length<=60 &&
                   !/(成交|行情|最新|价格|Bid|Ofr|笔数|更新时间|流动性|中债|中证|主承销|发行日|剩余期限|收益率)/.test(t) &&
                   /有限公司|公司|银行|政府|财政部|集团|投资|控股|开发|管理局|厅|委员会|国资|资产|有限/.test(t);
          }).sort((a,b) => a.getBoundingClientRect().y - b.getBoundingClientRect().y);
          if (cands.length) {
            let name = cands[0].textContent.trim();
            // 去掉评级及之后内容（公司全称在最前）
            const rc = name.search(/主体评级|债项评级|评级/);
            if (rc > 0) name = name.slice(0, rc).trim();
            // 按空白分词，精准去除 "省XX"(省+已知省份) 与 "性质XX" 字段（避免误切"湖北省…"公司名）
            const toks = name.split(/\s+/);
            let ci = -1;
            for (let ti=0; ti<toks.length; ti++) {
              const tk = toks[ti];
              const m = tk.match(/^省(.{2,5})$/);
              if (m) {
                let p = m[1];
                if (p.endsWith('省') || p.endsWith('市')) p = p.slice(0, -1);
                if (PROV_LIST.includes(p)) { ci = ti; break; }
              }
              if (tk.startsWith('性质')) { ci = ti; break; }
            }
            if (ci > 0) name = toks.slice(0, ci).join(' ').trim();
            // 兜底: 去掉末尾直接拼接的 性质值(市政市场化/城投市场化/城投/国企/央企等)
            name = name.replace(/市政市场化.*$|城投市场化.*$|市场化.*$/, '')
                       .replace(/(城投|国企|央企|民企|地方国企|中央国企|国有)$/, '')
                       .trim();
            result.issuerFull = name;
          }
          // 注意：票面利率/债券余额/到期日/YY评分/久期/中债净价 等数值列一律来自「我的关注」(上框)，
          // 不在此处从详情页抓取（避免与用户要求"Sheet2/数值列源自我的关注"冲突）。故 extra 留空。
          result.extra = {};
          return result;
        }, PROV_LIST);
      }

      // 首只进入详情页
      let onDetail = false;
      const firstName = getLowerVal(needRegion[0], '债券简称');
      for (let attempt=0; attempt<3 && !onDetail; attempt++) {
        if (attempt>0) await sleep(800);
        onDetail = await dblclickIntoDetail(targetFrame, firstName);
        log(`  首只双击进入详情页(尝试${attempt+1}): ${onDetail ? '✅' : '❌'}`);
      }

      // 逐只处理
      for (let ri = 0; ri < needRegion.length; ri++) {
        const ur = needRegion[ri];
        const bCode = getLowerVal(ur, '债券代码') || '';
        const bName = getLowerVal(ur, '债券简称') || '';
        if (!bCode) continue;

        if (ri > 0 || !onDetail) {
          const ok = await searchSwitch(targetFrame, bCode);
          if (!ok) { log(`  [${ri+1}] ${bName} 导航失败，跳过`); continue; }
          onDetail = true;
        }

        const ex = await extractFrom(targetFrame);
        bondRegionMap.set(bCode, {
          region: ex.province,
          bondName: bName,
          method: 'detail_' + (ex.method || 'none'),
          city: ex.city,
          issuerFull: ex.issuerFull,
          extra: ex.extra || {}
        });
        log(`  [${ri+1}/${needRegion.length}] ${bName} → 省:${ex.province||'(空)'} 全称:${ex.issuerFull||'(无)'} (${ex.method||'none'}) extra=${JSON.stringify(ex.extra||{})} `);
        try { await page.screenshot({ path: path.join(SCREENSHOT_DIR, `detail-${bCode.split('.').join('_')}.png`) }); } catch(e) {}
      }

      // 关闭所有打开的详情标签页：回到我的关注 → 点"关闭其他标签页"
      try {
        // 先回到我的关注
        await targetFrame.evaluate(() => {
          window.location.hash = '#/bond/my-focus';
          window.dispatchEvent(new HashChangeEvent('hashchange'));
        });
        await sleep(1500);
        // 点顶部导航栏的"关闭其他标签页"
        const closed = await page.evaluate(() => {
          const els = Array.from(document.querySelectorAll('*'));
          const btn = els.find(el => {
            const t = (el.textContent||'').trim();
            return t.includes('关闭其他标签页') && el.getBoundingClientRect().width > 0;
          });
          if (btn) { btn.click(); return true; }
          return false;
        });
        log(`  关闭其他标签页: ${closed ? '✅' : '❌ 未找到按钮'}`);
      } catch(e) { log(`  关闭标签页异常: ${e.message}`); }

      log(`  详情页提取完成，覆盖: ${[...bondRegionMap.values()].filter(v=>v.region).length}/${bondRegionMap.size}`);
    }

    // 最终区域覆盖统计
    const regionStats = [...bondRegionMap.values()];
    log(`  区域填充: ${regionStats.filter(v=>v.region).length}/${bondRegionMap.size} (映射+详情页)`);

    // Sheet1 = 下面框（成交行情，筛选 ≥MIN_VOLUME）
    const ws = wb.addWorksheet('成交行情_下框');
    if (filtered.length > 0) {
      const colCount = colDefs.length;
      const header = colDefs.map(c => c.name);
      ws.addRow(header);
      for (const r of filtered) {
        const row = r.cells.slice(0, colCount).map(cleanCell);
        while (row.length < colCount) row.push('');
        ws.addRow(row);
      }
    } else {
      ws.addRow(['暂无数据']);
    }
    // Sheet2 = 上面框（关注列表_我的关注）：数据严格来自「我的关注」(upperRows) 提取。
    // 以下框(filtered)为行基准保证 10 行一致；仅两框共有列(剩余期限/最新成交等)在下框回退；
    // 绝不从债券详情页回填（详情页只用于 Sheet3 的区域/发行人全称这类只有详情页才有的字段）。
    const ws2 = wb.addWorksheet('关注列表_上框');
    if (filtered.length > 0) {
      const upHeader = upperColDefs.map(c => c.name);
      ws2.addRow(upHeader);
      for (const r of filtered) {
        const code = (r.cells[1] || '').trim();
        const name = (r.bondName || '').trim();
        const u = findUpper(code, name);
        const row = upperColDefs.map(c => {
          const ui = uIdx[c.name];
          if (u && ui != null && ui >= 0 && u.cells[ui] != null && u.cells[ui] !== '') return cleanCell(u.cells[ui]);
          const li = lIdx[c.name];
          if (li != null && li >= 0 && r.cells[li] != null && r.cells[li] !== '') return cleanCell(r.cells[li]);
          return '';
        });
        ws2.addRow(row);
      }
      log(`  上面框(我的关注)筛选 ≥${MIN_VOLUME}：${filtered.length} 行（数据严格来自我的关注提取，缺券的上框专属列留空）`);
    } else {
      ws2.addRow(['暂无数据']);
    }

    // ===== Sheet3 = 成交行情汇总（按用户模板 17 列） =====
    log('9. 构建 Sheet3: 成交行情汇总');
    const SHEET3_HEADERS = ['成交日期', '债券简称', '剩余期限', '最新成交', '中债(行/到)', '中债偏离(BP)',
      '发行人简称', '区域', '债券代码', '中债隐含评级', '中债净价', '主/债',
      '票面利率(%)', '债券余额(亿)', '到期日', 'YY评分', '久期'];

    // 上框列名→index 映射（用于从 upperRows 取值）
    function getUpperVal(rowCells, colName) {
      const i = uIdx[colName];
      return (i != null && i >= 0 && rowCells[i]) ? cleanCell(rowCells[i]) : '';
    }
    function getLowerVal(r, colName) {
      const i = lIdx[colName];
      return (i != null && i >= 0 && r.cells[i]) ? cleanCell(r.cells[i]) : '';
    }


    // ---- 构建 Sheet3 数据 ----
    const ws3 = wb.addWorksheet('成交行情汇总');
    ws3.addRow(SHEET3_HEADERS);

    // 用上框筛选数据为主（字段更全），补入区域和成交日期（统一用北京时间日期）
    const [y, m, d] = BJ_DATE.split('-');
    const dateFormatted = `${y}-${parseInt(m)}-${parseInt(d)}`; // 2026-7-20

    let sheet3Count = 0;
    const sheet3Rows = []; // 收集当天 Sheet3 行，用于写入历史 json 与 Sheet4
    for (const r of filtered) {
      const code = getLowerVal(r, '债券代码');
      const name = getLowerVal(r, '债券简称');
      const u = findUpper(code, name);
      // 优先上框值；上框未匹配（实时不稳定缺券）时回退下框同名列，保证中债等列不空
      const gv = (cn) => {
        const up = u ? getUpperVal(u.cells, cn) : '';
        if (up) return up;
        return getLowerVal(r, cn);
      };

      // 从点击结果获取区域/发行人全称（这两列只有债券详情页才有，保留详情页来源）
      const rInfo = bondRegionMap.get(code) || bondRegionMap.get(name) || {};

      // 注意：除 G(发行人全称) / H(区域) 来自详情页外，其余数值列严格来自
      // 「我的关注」(gv: 上框优先 → 下框回退)，不从债券详情页回填。
      const row3 = [
        dateFormatted,                                              // A 成交日期
        name,                                                       // B 债券简称
        gv('剩余期限') || getLowerVal(r, '剩余期限'),                // C 剩余期限
        gv('最新成交') || getLowerVal(r, '最新成交'),                // D 最新成交
        gv('中债(行/到)'),                                          // E 中债(行/到)（上框缺则下框）
        gv('中债偏离(BP)'),                                         // F 中债偏离(BP)（上框缺则下框）
        rInfo.issuerFull || gv('发行人简称') || name,               // G 发行人简称(优先详情页全称)
        rInfo.region || '',                                         // H 区域（详情页获取）
        code,                                                       // I 债券代码
        gv('中债隐含评级') || getLowerVal(r, '中债隐含评级'),         // J 中债隐含评级
        gv('中债净价'),                                             // K 中债净价（仅我的关注）
        gv('主/债'),                                                // L 主/债
        gv('票面利率(%)'),                                          // M 票面利率(%)
        gv('债券余额(亿)'),                                         // N 债券余额(亿)
        gv('到期日'),                                               // O 到期日
        gv('YY评分'),                                               // P YY评分
        gv('久期'),                                                 // Q 久期
      ];
      ws3.addRow(row3);
      sheet3Rows.push(row3);
      sheet3Count++;
    }
    log(`  Sheet3 完成: ${sheet3Count} 行 (成交日期=${dateFormatted})`);

    // ===== 历史存储：把今天的 Sheet3 存为 daily/{BJ_DATE}.json（Sheet4 的单一数据源）=====
    const todayJsonPath = path.join(HISTORY_DIR, `${BJ_DATE}.json`);
    fs.writeFileSync(todayJsonPath, JSON.stringify({
      date: BJ_DATE,
      dateFormatted,
      headers: SHEET3_HEADERS,
      rows: sheet3Rows,
    }, null, 1), 'utf8');
    log(`  历史已写: ${todayJsonPath} (${sheet3Rows.length} 行)`);

    // ===== Sheet4 = 历史累积（读取所有 daily/*.json，按日期升序合并）=====
    log('10. 构建 Sheet4: 历史累积汇总');
    const ws4 = wb.addWorksheet('历史累积');
    ws4.addRow(SHEET3_HEADERS);
    let sheet4Count = 0, sheet4Days = 0;
    const dayFiles = fs.readdirSync(HISTORY_DIR)
      .filter(f => /^\d{4}-\d{2}-\d{2}\.json$/.test(f))
      .sort(); // 文件名即日期，字典序=时间序
    for (const f of dayFiles) {
      try {
        const j = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), 'utf8'));
        const rows = Array.isArray(j.rows) ? j.rows : [];
        for (const row of rows) { ws4.addRow(row); sheet4Count++; }
        if (rows.length) sheet4Days++;
      } catch (e) { log(`  ⚠️ 跳过损坏历史文件 ${f}: ${e.message}`); }
    }
    log(`  Sheet4 完成: ${sheet4Count} 行，覆盖 ${sheet4Days} 个交易日`);

    // 自动列宽：按单元格内容（含中文按 2 宽度计）调整，确保内容显示完整
    function autoWidth(worksheet, maxW) {
      const ncol = worksheet.columnCount;
      for (let i = 1; i <= ncol; i++) {
        const col = worksheet.getColumn(i);
        let maxLen = 0;
        col.eachCell({ includeEmpty: true }, cell => {
          const v = cell.value == null ? '' : String(cell.value);
          let len = 0;
          for (const ch of v) len += (ch.charCodeAt(0) > 255 ? 2 : 1);
          if (len > maxLen) maxLen = len;
        });
        col.width = Math.min(Math.max(maxLen + 2, 10), maxW || 60);
      }
    }
    autoWidth(ws, 60);
    autoWidth(ws2, 60);
    autoWidth(ws3, 60);
    autoWidth(ws4, 60); // Sheet4 历史累积也要自动列宽

    // 固定输出路径到 SECONDARY_DIR，文件名按北京时间日期，便于本地/云端同步与历史累积
    const outName = `secondary_quote_${BJ_DATE}.xlsx`;
    const outPath = path.join(SECONDARY_DIR, outName);
    await wb.xlsx.writeFile(outPath);
    log(`✅ 已保存: ${outPath}`);
    log(`   共 4 Sheet: 下框${filtered.length} + 上框${filtered.length} + 当天汇总${sheet3Count} + 历史累积${sheet4Count}(${sheet4Days}个交易日)（均≥${MIN_VOLUME}）`);

    // 截全页（含上面框+下面框），便于核对
    const finalShot = path.join(SCREENSHOT_DIR, '07-final.png');
    await page.screenshot({ path: finalShot, fullPage: true });
    log(`截图(全页): 07-final`);

  } catch (err) {
    log('❌ 错误:', err.message);
    log(err.stack);
    await screenshot(page, 'error');
  } finally {
    log('关闭浏览器');
    await context.close();
    // 清理临时目录
    try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch (e) {}
  }
}

main().catch(e => {
  console.error('Fatal:', e);
  process.exit(1);
});
