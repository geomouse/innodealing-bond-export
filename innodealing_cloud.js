const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const DATA_DIR = path.join(__dirname, 'data');
const LOGIN_URL = 'https://web.innodealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innodealing.com/quote-web/#/bond/primary-issue/new-issue-bond/credit-bond-issue';

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function getDateStr() {
  // 使用北京时间 (UTC+8)
  const now = new Date();
  const beijing = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  return beijing.toISOString().split('T')[0];
}

async function main() {
  const username = process.env.INNODEALING_USERNAME;
  const password = process.env.INNODEALING_PASSWORD;

  if (!username || !password) {
    console.error('缺少环境变量: INNODEALING_USERNAME 或 INNODEALING_PASSWORD');
    process.exit(1);
  }

  const dateStr = getDateStr();
  console.log(`=== 债立方信用债导出 | ${dateStr} (北京时间) ===`);

  const browser = await chromium.launch({
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--disable-extensions',
      '--disable-background-networking',
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
    console.log('[1/5] 登录...');
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(2000);

    // 检查是否已登录（如果已登录则跳过登录步骤）
    if (!page.url().includes('signin')) {
      console.log('  [OK] 已有有效登录会话，跳过登录');
    } else {
      await page.locator('input[placeholder*="手机"]').first().fill(username);
      await page.locator('input[type="password"]').first().fill(password);

      // 勾选服务条款
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

      // 验证登录
      let loggedIn = false;
      for (let i = 0; i < 10; i++) {
        if (!page.url().includes('signin')) {
          loggedIn = true;
          break;
        }
        await sleep(1000);
      }
      if (!loggedIn) {
        throw new Error('登录超时，仍在登录页');
      }
      console.log('  [OK] 登录成功');
    }

    // ===== STEP 2: 导航到信用债发行页面 =====
    console.log('[2/5] 导航到信用债发行页面...');

    // 先导航到主页面以触发 iframe 加载
    try {
      await page.goto('https://web.innodealing.com/quote-web/#/bond/primary-issue', {
        waitUntil: 'domcontentloaded', timeout: 30000,
      });
    } catch (e) {
      console.log('  主页导航:', e.message.substring(0, 80));
    }
    await sleep(3000);

    // 等待并定位 targetFrame
    let targetFrame = null;
    for (let i = 0; i < 60; i++) {
      targetFrame = page.frames().find(f => f.url().includes('/quote-web/'));
      if (targetFrame) break;
      await sleep(500);
    }

    if (!targetFrame) {
      console.log('  可用 frames:', page.frames().map(f => f.url().substring(0, 80)));
      throw new Error('未找到 quote-web frame');
    }

    console.log(`  [OK] 找到 frame: ${targetFrame.url().substring(0, 80)}`);

    // 在 targetFrame 中导航到信用债发行页面
    await targetFrame.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(5000);

    // 关闭可能的弹窗
    await page.keyboard.press('Escape');
    await sleep(1000);
    console.log('  [OK] 已加载信用债发行页面');

    // ===== STEP 3: 设置发行起始日 =====
    console.log(`[3/5] 设置发行起始日为: ${dateStr}...`);

    // 在 targetFrame 中找日期输入框
    const dateSet = await targetFrame.evaluate((ds) => {
      const inputs = document.querySelectorAll('input');
      for (const inp of inputs) {
        const ph = inp.getAttribute('placeholder') || '';
        if (ph.includes('起始日') || ph.includes('开始日期')) {
          // 清空并设置日期
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

    if (dateSet) {
      console.log(`  [OK] 已设置发行起始日: ${dateStr}`);
    } else {
      // 尝试直接在 targetFrame 的 input[placeholder] 上操作
      const dateInput = targetFrame.locator('input[placeholder*="起始"]').first();
      if (await dateInput.count() > 0) {
        await dateInput.click();
        await sleep(300);
        await dateInput.fill('');
        await dateInput.fill(dateStr);
        await dateInput.press('Enter');
        console.log(`  [OK] 已设置发行起始日(备选): ${dateStr}`);
      } else {
        console.log('  [WARN] 未找到日期输入框');
      }
    }

    await page.keyboard.press('Escape');
    await sleep(3000);

    // ===== STEP 4: 选择主体组 all-A =====
    console.log('[4/5] 选择自选管理-主体组 all-A...');

    const foundAllA = await targetFrame.evaluate(() => {
      // 先检查是否已经选择了 all-A
      const selectors = document.querySelectorAll('.dmuiv4-select');
      for (const sel of selectors) {
        const text = sel.textContent || '';
        if (text.includes('all-A')) return 'already_selected';
      }
      return 'need_select';
    });

    if (foundAllA === 'already_selected') {
      console.log('  [OK] 已选择主体组 all-A (已存在)');
    } else {
      // 需要选择
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
            if (item.textContent.trim() === 'all-A') {
              item.click();
              return true;
            }
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

    // ===== STEP 5: 导出数据 =====
    console.log('[5/5] 导出数据...');

    // 先检查数据是否已加载
    const rowCount = await targetFrame.evaluate(() => {
      const table = document.querySelector('table');
      if (!table) return -1;
      const rows = table.querySelectorAll('tbody tr, tr[class*="row"]');
      return rows.length;
    });
    console.log(`  检测到 ${rowCount} 条数据行`);

    if (rowCount === -1) {
      console.log('  [WARN] 未找到数据表格，尝试继续导出');
    }

    // 设置下载监听（在点击前）
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
          console.log(`\n[SUCCESS] 下载成功: ${savePath}`);
          console.log(`  文件大小: ${(fs.statSync(savePath).size / 1024).toFixed(1)} KB`);
          resolve(savePath);
        } catch (e) {
          downloadError = `save: ${e.message}`;
          console.error(`  保存失败: ${e.message}`);
          resolve(null);
        }
      });
    });

    // 点击导出按钮
    try {
      const exportBtn = targetFrame.locator('button:has-text("导出数据"), button:has-text("导出")').first();
      await exportBtn.click({ force: true });
      console.log('  [OK] 已点击导出按钮，等待下载...');
    } catch (e) {
      console.log(`  [WARN] 点击导出按钮失败: ${e.message}`);
    }

    const downloadPath = await downloadPromise;

    if (downloadPath) {
      console.log(`\n✅ 导出完成: ${downloadPath}`);
    } else {
      console.log(`\n❌ 下载失败: ${downloadError || '未知原因'}`);

      // 截图用于调试
      try {
        await page.screenshot({ path: path.join(DATA_DIR, 'debug_screenshot.png') });
        const html = await targetFrame.evaluate(() => document.body.innerHTML);
        fs.writeFileSync(path.join(DATA_DIR, 'debug_page.html'), html.substring(0, 50000));
        console.log('  调试截图和HTML已保存');
      } catch (e) {
        console.log('  调试文件保存失败:', e.message);
      }
    }

    // 列出 data 目录
    console.log('\n=== data 目录内容 ===');
    fs.readdirSync(DATA_DIR).forEach(f => {
      const stat = fs.statSync(path.join(DATA_DIR, f));
      const kb = (stat.size / 1024).toFixed(1);
      console.log(`  ${f}  (${kb} KB)`);
    });

    return { success: !!downloadPath, date: dateStr, path: downloadPath };

  } catch (err) {
    console.error('\n[FATAL] 发生错误:', err.message);
    console.error(err.stack);

    // 保存错误截图
    try {
      await page.screenshot({ path: path.join(DATA_DIR, 'error_screenshot.png'), fullPage: true });
      console.log('  错误截图已保存');
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
