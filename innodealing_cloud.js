const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const DATA_DIR = path.join(__dirname, 'data');
const LOGIN_URL = 'https://web.innodealing.com/auth-service/signin';
const TARGET_URL = 'https://web.innodealing.com/quote-web/#/bond/primary-issue/new-issue-bond/credit-bond-issue';

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function main() {
  const username = process.env.INNODEALING_USERNAME;
  const password = process.env.INNODEALING_PASSWORD;

  if (!username || !password) {
    console.error('缺少环境变量: INNODEALING_USERNAME 或 INNODEALING_PASSWORD');
    process.exit(1);
  }

  const today = new Date();
  const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  console.log(`=== 债立方信用债导出 | ${dateStr} ===`);

  const browser = await chromium.launch({
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
    ],
  });

  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1920, height: 1080 },
    locale: 'zh-CN',
  });

  const page = await context.newPage();
  let targetFrame = null;

  try {
    // ===== STEP 1: 登录 =====
    console.log('[1/5] 登录...');
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.locator('input[placeholder*="手机"]').first().fill(username);
    await page.locator('input[type="password"]').first().fill(password);
    await page.getByText('我已阅读并同意相关服务条款和政策').first().click();
    await page.locator('button:has-text("登录")').first().click();
    await sleep(5000);

    if (page.url().includes('signin')) {
      throw new Error('登录失败，仍停留在登录页');
    }
    console.log('  [OK] 登录成功');

    // ===== STEP 2: 导航到信用债发行页面 =====
    console.log('[2/5] 导航到信用债发行页面...');
    for (let i = 0; i < 60; i++) {
      targetFrame = page.frames().find(f => f.url().includes('/quote-web/'));
      if (targetFrame) break;
      await sleep(500);
    }
    if (!targetFrame) throw new Error('未找到 quote-web frame');

    await targetFrame.goto(TARGET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(4000);
    await page.keyboard.press('Escape');
    await sleep(500);
    console.log('  [OK] 已加载信用债发行页面');

    // ===== STEP 3: 设置发行起始日为今天 =====
    console.log('[3/5] 设置发行起始日为今天...');
    const dateInputs = targetFrame.locator('input');
    const count = await dateInputs.count();
    for (let i = 0; i < count; i++) {
      const input = dateInputs.nth(i);
      const ph = await input.getAttribute('placeholder').catch(() => '');
      if (ph && (ph.includes('起始日') || ph.includes('开始日期'))) {
        await input.click();
        await input.fill('');
        await input.fill(dateStr);
        await input.press('Enter');
        console.log(`  [OK] 已设置发行起始日: ${dateStr}`);
        break;
      }
    }
    await page.keyboard.press('Escape');
    await sleep(2000);

    // ===== STEP 4: 选择主体组 all-A =====
    console.log('[4/5] 选择自选管理-主体组 all-A...');
    const box = await targetFrame.evaluate(() => {
      const sel = Array.from(document.querySelectorAll('.dmuiv4-select')).find(s => s.textContent.includes('请选择主体组'));
      if (!sel) return null;
      const rect = sel.getBoundingClientRect();
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    });

    if (box) {
      await page.mouse.click(box.x + box.width - 12, box.y + box.height / 2);
      await sleep(500);

      const optionClicked = await targetFrame.evaluate(() => {
        const items = document.querySelectorAll('.dmuiv4-select-item-option-content');
        for (const item of items) {
          if (item.textContent.trim() === 'all-A') {
            item.click();
            return true;
          }
        }
        return false;
      });

      if (optionClicked) {
        console.log('  [OK] 已选择主体组 all-A');
      } else {
        console.log('  [WARN] 未找到主体组 all-A 选项');
      }
    } else {
      console.log('  [WARN] 未找到主体组下拉框');
    }

    await page.keyboard.press('Escape');
    await sleep(2000);

    // ===== STEP 5: 导出数据 =====
    console.log('[5/5] 导出数据...');

    const downloadPromise = page.waitForEvent('download', { timeout: 30000 }).catch(() => null);
    try {
      const exportBtn = targetFrame.locator('button').filter({ hasText: /导出数据/ }).first();
      await exportBtn.click({ force: true });
      console.log('  [OK] 已点击导出数据按钮');
    } catch (e) {
      console.log('  [WARN] 点击导出按钮失败:', e.message);
    }
    await sleep(3000);

    const download = await downloadPromise;
    if (download) {
      const savePath = path.join(DATA_DIR, `credit_bond_${dateStr}.xlsx`);
      await download.saveAs(savePath);
      console.log(`\n[SUCCESS] 下载成功: ${savePath}`);
    } else {
      console.log('\n[WARN] 未检测到下载事件');
    }

    // 列出 data 目录内容
    console.log('\n=== data 目录内容 ===');
    fs.readdirSync(DATA_DIR).forEach(f => console.log(`  ${f}`));

    return { success: true, date: dateStr };

  } catch (err) {
    console.error('\n[ERROR] 发生错误:', err.message);
    console.error(err.stack);
    return { success: false, error: err.message };
  } finally {
    await browser.close();
    console.log('浏览器已关闭');
  }
}

main().then(result => {
  if (result.success) {
    console.log('\n✅ 导出完成');
    process.exit(0);
  } else {
    console.log('\n❌ 导出失败');
    process.exit(1);
  }
});
