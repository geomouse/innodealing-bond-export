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

// 可靠地设置一级发行页面的「日期区间」(RangePicker) 为单日 dateStr，并触发查询刷新。
// 真实结构：页面有一个下拉(dmuiv4-select)选择按哪个日期维度（默认=发行起始日）+ 一个 RangePicker，
// 其两个 input 的 placeholder 分别为『开始日期』与『结束日期』（绝非『发行起始日』，那是维度下拉的文本）。
// 把 开始日期 与 结束日期 都设为同一天 -> 区间即“当天”，筛出当日发行。
// 关键点：React 通过 _valueTracker 追踪输入值，必须 tracker.setValue('') 清空后再 set 新值并派发 input/change，
// 否则 React 误判“未变化”而忽略，日期框不刷新、查询不发起。
async function setRangeDate(frame, dateStr) {
  return await frame.evaluate((ds) => {
    const inputs = Array.from(document.querySelectorAll('input'));
    const start = inputs.find(i => (i.getAttribute('placeholder') || '') === '开始日期');
    const end = inputs.find(i => (i.getAttribute('placeholder') || '') === '结束日期');
    if (!start || !end) return 'NO_INPUT:' + (!start ? 'no-start' : '') + (!end ? 'no-end' : '');
    function setVal(inp, v) {
      try { inp.focus(); } catch (e) {}
      const proto = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
      const tracker = inp._valueTracker;
      if (tracker) { try { tracker.setValue(''); } catch (e) {} }
      proto.set.call(inp, v);
      inp.dispatchEvent(new Event('input', { bubbles: true }));
      inp.dispatchEvent(new Event('change', { bubbles: true }));
    }
    setVal(start, ds);
    setVal(end, ds);
    return 'SET';
  }, dateStr);
}

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
// applyAllA=true  → 筛选关注组 all-A（写入 credit_bond_{date}.xlsx，用于主库/Sheet1/Sheet2）
// applyAllA=false → 不筛选关注组（页面默认全市场，写入 credit_bond_all_{date}.xlsx，用于新 sheet）
async function exportForDate(page, targetFrame, dateStr, applyAllA = true) {
  console.log(`  === 导出 ${dateStr} ===`);

  // 诊断：打印所有 input（placeholder/type/class/readonly）与日期相关元素，定位真实日期选择器
  try {
    const diag = await targetFrame.evaluate(() => {
      const inputs = Array.from(document.querySelectorAll('input')).map(i => ({
        ph: i.getAttribute('placeholder') || '',
        type: i.getAttribute('type') || '',
        cls: (i.className || '').slice(0, 80),
        readonly: i.readOnly,
      }));
      const dateEls = Array.from(document.querySelectorAll('*')).filter(e => {
        const t = (e.textContent || '').trim();
        const c = (e.className || '');
        return (t.includes('起始') || t.includes('截标') || t.includes('发行日') || t.includes('日期') || t.includes('申购'))
          && (c.includes('picker') || c.includes('Picker') || c.includes('date') || c.includes('Date') || c.includes('select') || c.includes('Select') || c.includes('input') || e.tagName === 'INPUT' || e.tagName === 'LABEL');
      }).slice(0, 40).map(e => ({ tag: e.tagName, cls: (e.className||'').slice(0,80), text: (e.textContent||'').trim().slice(0,40) }));
      return { inputs: inputs.slice(0, 50), dateEls };
    });
    console.log('  [DIAG-INPUTS]', JSON.stringify(diag.inputs));
    console.log('  [DIAG-DATEELS]', JSON.stringify(diag.dateEls));
  } catch (e) { console.log('  [DIAG] err', e.message); }

  // 两种模式都显式把日期区间设为单日 dateStr（今天），避免页面默认停在前一天。
  // RangePicker 的 input placeholder 为『开始日期』『结束日期』；把两者都设为同一天即“当日”。
  console.log(`    [INFO] 设置日期区间为单日 ${dateStr} ...`);
  const setResult = await setRangeDate(targetFrame, dateStr);
  console.log(`    [DATE] setRangeDate -> ${setResult}`);
  if (setResult !== 'SET') {
    try {
      const ds = targetFrame.locator('input[placeholder="开始日期"]').first();
      const de = targetFrame.locator('input[placeholder="结束日期"]').first();
      await ds.fill(''); await ds.fill(dateStr); await ds.press('Enter');
      await de.fill(''); await de.fill(dateStr); await de.press('Enter');
      console.log(`    [DATE] Playwright fill 重设日期区间并回车 OK`);
    } catch (e) {
      console.log(`    [WARN] 重设日期区间兜底失败: ${e.message}`);
    }
  }
  await sleep(3000);

  // 兜底：点击“查询/搜索/刷新”按钮，强制刷新列表
  try {
    const clicked = await targetFrame.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      for (const b of btns) {
        const t = (b.textContent || '').trim();
        if (t === '查询' || t === '搜索' || t === '刷新' || t.includes('查询')) { b.click(); return t; }
      }
      return '';
    });
    if (clicked) console.log(`    [OK] 已点击「${clicked}」按钮刷新列表`);
  } catch (e) {
    console.log(`    [WARN] 点击刷新/查询按钮异常: ${e.message}`);
  }
  await sleep(3000);

  // 2. 确认主体组 all-A 仍然选中（仅筛选模式；未筛选模式跳过，保留页面默认全市场视图）
  if (applyAllA) {
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
  } else {
    console.log(`    [INFO] ${dateStr} 未筛选模式（全市场），不做关注组选择`);
  }

  // 3. 导出数据
  let downloadPath = null;
  let downloadError = null;
  const prefix = applyAllA ? '' : 'all_';
  const downloadPromise = new Promise((resolve) => {
    const timer = setTimeout(() => {
      downloadError = 'timeout';
      resolve(null);
    }, 35000);

    page.once('download', async (download) => {
      clearTimeout(timer);
      try {
        const savePath = path.join(DATA_DIR, `credit_bond_${prefix}${dateStr}.xlsx`);
        await download.saveAs(savePath);
        const size = fs.statSync(savePath).size;
        console.log(`    [OK] 下载成功: credit_bond_${prefix}${dateStr}.xlsx (${(size/1024).toFixed(1)} KB)`);
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

    // ===== STEP 3.5: 导出未筛选(全市场)当日数据 → credit_bond_all_{today}.xlsx =====
    // 用于新 sheet（不筛选关注组、新债预测≥2.0 的全市场新债）。
    // 此时尚未选择任何主体组，页面默认即全市场视图；该文件独立保存，后续 STEP3 选 all-A 不影响它。
    console.log('[3.5] 导出未筛选(全市场)当日数据...');
    const unfilteredDate = businessDays[0]; // 今天（北京时间，businessDays[0] 为最近交易日）
    const unResult = await exportForDate(page, targetFrame, unfilteredDate, false);
    if (unResult.success) {
      console.log(`  [OK] 未筛选当日数据已导出: credit_bond_all_${unfilteredDate}.xlsx`);
    } else {
      console.log(`  [WARN] 未筛选当日数据导出失败（不影响关注组主流程）: ${unfilteredDate}`);
    }
    await sleep(1000);

    // ===== STEP 3: 选择主体组 all-A（带校验，失败即终止，避免混入非关注组）=====
    console.log('[3/4] 选择主体组 all-A...');
    const hasAllASelected = () => targetFrame.evaluate(() => {
      const selectors = document.querySelectorAll('.dmuiv4-select');
      for (const sel of selectors) {
        if (sel.textContent.includes('all-A')) return true;
      }
      return false;
    });

    let ok = await hasAllASelected();
    if (!ok) {
      for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
        const box = await targetFrame.evaluate(() => {
          const sel = Array.from(document.querySelectorAll('.dmuiv4-select'))
            .find(s => s.textContent.includes('请选择主体组') || s.textContent.includes('all-A'));
          if (!sel) return null;
          const rect = sel.getBoundingClientRect();
          return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
        });
        if (box) {
          await page.mouse.click(box.x + box.width - 12, box.y + box.height / 2);
          await sleep(1200);
          const clicked = await targetFrame.evaluate(() => {
            const items = document.querySelectorAll('.dmuiv4-select-item-option-content');
            for (const item of items) {
              if (item.textContent.trim() === 'all-A') { item.click(); return true; }
            }
            return false;
          });
          await sleep(800);
          ok = await hasAllASelected();
          console.log(`  [attempt ${attempt}] 点击all-A=${clicked} 校验已选中=${ok}`);
          if (!ok) { await page.keyboard.press('Escape'); await sleep(500); }
        } else {
          console.log(`  [WARN] 未找到主体组下拉框 (attempt ${attempt})`);
          await sleep(500);
        }
      }
    }
    if (!ok) {
      console.log('  [FATAL] 主体组 all-A 选择校验失败，终止导出以避免混入非关注组债券');
      process.exit(1);
    }
    console.log('  [OK] 已确认主体组 all-A 选中');
    await page.keyboard.press('Escape');
    await sleep(2000);

    // ===== STEP 4: 逐日导出 =====
    console.log(`[4/4] 逐日导出 ${businessDays.length} 个交易日...`);
    const results = [];
    for (const dateStr of businessDays) {
      const result = await exportForDate(page, targetFrame, dateStr, true);
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
