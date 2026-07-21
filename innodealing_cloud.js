const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const DATA_DIR = path.join(__dirname, 'data');
const LOGIN_URL = 'https://web.innodealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innodealing.com/quote-web/#/bond/primary-issue/new-issue-bond/credit-bond-issue';

// 导出最近 N 个交易日（含今天），确保票面利率补全
const EXPORT_BUSINESS_DAYS = 5;

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// 计算最近 N 个交易日（北京时间 UTC+8，跳过周末）
function getBusinessDays(count) {
  const days = [];
  const now = new Date();
  const beijing = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  let d = new Date(beijing.getUTCFullYear(), beijing.getUTCMonth(), beijing.getUTCDate());
  while (days.length < count) {
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) {
      days.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`);
    }
    d.setDate(d.getDate() - 1);
  }
  return days;
}

// 导出单个日期的数据
async function exportForDate(page, targetFrame, dateStr) {
  console.log(`  === 导出 ${dateStr} ===`);

  // 1. 设置发行起始日
  const dateSet = await targetFrame.evaluate((ds) => {
    const inputs = document.querySelectorAll('input');
    for (const inp of inputs) {
      const ph = inp.getAttribute('placeholder') || '';
      if (ph.includes('起始日') || ph.includes('开始日期')) {
        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
          window.HTMLInputElement.prototype, 'value'
        ).set;
        nativeInputValueSetter.call(inp, '');
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        nativeInputValueSetter.call(inp, ds);
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      }
    }
    return false;
  }, dateStr);

  if (!dateSet) {
    const dateInput = targetFrame.locator('input[placeholder*="起始"]').first();
    if (await dateInput.count() > 0) {
      await dateInput.click();
      await sleep(300);
      await dateInput.fill('');
      await dateInput.fill(dateStr);
      await dateInput.press('Enter');
    } else {
      console.log(`    [WARN] ${dateStr} 未找到日期输入框`);
      return { date: dateStr, success: false };
    }
  }
  console.log(`    [OK] 已设置发行起始日: ${dateStr}`);
  await page.keyboard.press('Escape');
  await sleep(3000);

  // 2. 确认主体组 all-A 仍然选中
  const allAStillSelected = await targetFrame.evaluate(() => {
    const selectors = document.querySelectorAll('.dmuiv4-select');
    for (const sel of selectors) {
      if (sel.textContent.includes('all-A')) return true;
    }
    return false;
  });
  if (!allAStillSelected) {
    console.log(`    [INFO] ${dateStr} 主体组未选中，重新选择...`);
    const box = await targetFrame.evaluate(() => {
      const sel = Array.from(document.querySelectorAll('.dmuiv4-select'))
        .find(s => s.textContent.includes('请选择主体组'));
      if (!sel) return null;
      const rect = sel.getBoundingClientRect();
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    });
    if (box) {
      await page.mouse.click(box.x + box.width - 12, box.y + box.height / 2);
      await sleep(1000);
      await targetFrame.evaluate(() => {
        const items = document.querySelectorAll('.dmuiv4-select-item-option-content');
        for (const item of items) {
          if (item.textContent.trim() === 'all-A') { item.click(); return true; }
        }
        return false;
      });
      await sleep(1000);
    }
  }

  // 3. 导出数据
  let downloadPath = null;
  let downloadError = null;
  const downloadPromise = new Promise((resolve) => {
    const timer = setTimeout(() => {
      downloadError = 'timeout';
      resolve(null);
    }, 35000);

    page.once('download', async (download) => {
      clearTimeout(timer);
      try {
        const savePath = path.join(DATA_DIR, `credit_bond_${dateStr}.xlsx`);
        await download.saveAs(savePath);
        const size = fs.statSync(savePath).size;
        console.log(`    [OK] 下载成功: credit_bond_${dateStr}.xlsx (${(size/1024).toFixed(1)} KB)`);
        resolve(savePath);
      } catch (e) {
        downloadError = `save: ${e.message}`;
        console.error(`    [ERR] 保存失败: ${e.message}`);
        resolve(null);
      }
    });
  });

  try {
    const exportBtn = targetFrame.locator('button:has-text("导出数据"), button:has-text("导出")').first();
    await exportBtn.click({ force: true });
    console.log(`    [OK] 已点击导出按钮`);
  } catch (e) {
    console.log(`    [WARN] 点击导出按钮失败: ${e.message}`);
  }

  downloadPath = await downloadPromise;
  return { date: dateStr, success: !!downloadPath, path: downloadPath };
}

async function main() {
  const username = process.env.INNODEALING_USERNAME;
  const password = process.env.INNODEALING_PASSWORD;
  if (!username || !password) {
    console.error('缺少环境变量: INNODEALING_USERNAME 或 INNODEALING_PASSWORD');
    process.exit(1);
  }

  const businessDays = getBusinessDays(EXPORT_BUSINESS_DAYS);
  console.log(`=== 债立方信用债多日导出 | 北京时间 ===`);
  console.log(`导出日期: ${businessDays.join(', ')}\n`);

  const browser = await chromium.launch({
    headless: true,
    args: [
      '--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage',
      '--disable-gpu', '--disable-extensions', '--disable-background-networking',
    ],
  });

  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1920, height: 1080 },
    locale: 'zh-CN',
  });

  const page = await context.newPage();

  try {
    // ===== STEP 1: 登录 =====
    console.log('[1/4] 登录...');
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(2000);

    if (!page.url().includes('signin')) {
      console.log('  [OK] 已有有效登录会话，跳过登录');
    } else {
      await page.locator('input[placeholder*="手机"]').first().fill(username);
      await page.locator('input[type="password"]').first().fill(password);

      try {
        const checkboxText = page.getByText(/服务条款/).first();
        if (await checkboxText.isVisible({ timeout: 2000 }).catch(() => false)) {
          await checkboxText.click();
          console.log('  [OK] 已勾选服务条款');
        }
      } catch (e) {
        console.log('  [WARN] 服务条款复选框未找到，跳过');
      }

      await page.locator('button:has-text("登录")').first().click();
      await sleep(6000);

      let loggedIn = false;
      for (let i = 0; i < 10; i++) {
        if (!page.url().includes('signin')) { loggedIn = true; break; }
        await sleep(1000);
      }
      if (!loggedIn) throw new Error('登录超时，仍在登录页');
      console.log('  [OK] 登录成功');
    }

    // ===== STEP 2: 导航到信用债发行页面 =====
    console.log('[2/4] 导航到信用债发行页面...');
    try {
      await page.goto('https://web.innodealing.com/quote-web/#/bond/primary-issue', {
        waitUntil: 'domcontentloaded', timeout: 30000,
      });
    } catch (e) {
      console.log('  主页导航:', e.message.substring(0, 80));
    }
    await sleep(3000);

    let targetFrame = null;
    for (let i = 0; i < 60; i++) {
      targetFrame = page.frames().find(f => f.url().includes('/quote-web/'));
      if (targetFrame) break;
      await sleep(500);
    }
    if (!targetFrame) throw new Error('未找到 quote-web frame');
    console.log(`  [OK] 找到 frame: ${targetFrame.url().substring(0, 80)}`);

    await targetFrame.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(5000);
    await page.keyboard.press('Escape');
    await sleep(1000);
    console.log('  [OK] 已加载信用债发行页面');

    // ===== STEP 3: 选择主体组 all-A =====
    console.log('[3/4] 选择主体组 all-A...');
    const foundAllA = await targetFrame.evaluate(() => {
      const selectors = document.querySelectorAll('.dmuiv4-select');
      for (const sel of selectors) {
        if (sel.textContent.includes('all-A')) return 'already_selected';
      }
      return 'need_select';
    });

    if (foundAllA === 'already_selected') {
      console.log('  [OK] 已选择主体组 all-A (已存在)');
    } else {
      const box = await targetFrame.evaluate(() => {
        const sel = Array.from(document.querySelectorAll('.dmuiv4-select'))
          .find(s => s.textContent.includes('请选择主体组'));
        if (!sel) return null;
        const rect = sel.getBoundingClientRect();
        return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
      });
      if (box) {
        await page.mouse.click(box.x + box.width - 12, box.y + box.height / 2);
        await sleep(1000);
        await targetFrame.evaluate(() => {
          const items = document.querySelectorAll('.dmuiv4-select-item-option-content');
          for (const item of items) {
            if (item.textContent.trim() === 'all-A') { item.click(); return true; }
          }
          return false;
        });
        await sleep(500);
        console.log('  [OK] 已选择主体组 all-A');
      } else {
        console.log('  [WARN] 未找到主体组下拉框');
      }
    }
    await page.keyboard.press('Escape');
    await sleep(3000);

    // ===== STEP 4: 逐日导出 =====
    console.log(`[4/4] 逐日导出 ${businessDays.length} 个交易日...`);
    const results = [];
    for (const dateStr of businessDays) {
      const result = await exportForDate(page, targetFrame, dateStr);
      results.push(result);
      await sleep(1000);
    }

    // 汇总结果
    console.log('\n=== 导出结果汇总 ===');
    let successCount = 0;
    for (const r of results) {
      const status = r.success ? '✅' : '❌';
      console.log(`  ${r.date}: ${status}`);
      if (r.success) successCount++;
    }
    console.log(`成功: ${successCount}/${results.length}`);

    // 列出 data 目录
    console.log('\n=== data 目录 ===');
    fs.readdirSync(DATA_DIR).filter(f => f.startsWith('credit_bond_')).sort().forEach(f => {
      const stat = fs.statSync(path.join(DATA_DIR, f));
      console.log(`  ${f}  (${(stat.size/1024).toFixed(1)} KB)`);
    });

    return { success: successCount > 0, results };

  } catch (err) {
    console.error('\n[FATAL] 发生错误:', err.message);
    console.error(err.stack);
    try {
      await page.screenshot({ path: path.join(DATA_DIR, 'error_screenshot.png'), fullPage: true });
    } catch (e) {}
    return { success: false, error: err.message };
  } finally {
    await browser.close();
    console.log('\n浏览器已关闭');
  }
}

main().then(result => {
  if (result.success) {
    console.log('\n✅ 导出成功');
    process.exit(0);
  } else {
    console.log('\n❌ 导出失败: ' + (result.error || '下载未完成'));
    process.exit(1);
  }
});
