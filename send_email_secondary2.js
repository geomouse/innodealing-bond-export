#!/usr/bin/env node
// 发送二级行情（我的关注 · 成交行情下框）汇总邮件（带 3-Sheet Excel 附件）
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

// 统一北京时间日期
const BJ_DATE = new Date(Date.now() + 8 * 3600 * 1000).toISOString().slice(0, 10);

async function sendEmail(excelPath, dateStr) {
  const emailUser = process.env.EMAIL_USER;
  const emailPass = process.env.EMAIL_PASS;
  const emailTo = process.env.EMAIL_TO;

  if (!emailUser || !emailPass || !emailTo) {
    console.log('[邮件] 未配置，跳过发送');
    return false;
  }

  console.log(`[邮件] 发送到 ${emailTo}...`);
  const transporter = nodemailer.createTransport({
    host: 'smtp.qq.com',
    port: 465,
    secure: true,
    auth: { user: emailUser, pass: emailPass },
  });

  const subject = `[二级行情成交汇总] ${dateStr} | 3-Sheet(下框源数据/当天汇总/历史累积)`;

  // 附带：排序后截图 + 自核对报告，便于人眼复核
  const extraAttach = [];
  const shotDir = path.join(secondary2Dir, 'screenshots');
  const candFiles = [];
  if (fs.existsSync(shotDir)) {
    for (const f of fs.readdirSync(shotDir)) {
      if (/^secondary2_verify_.*\.json$/.test(f) || /^05-after-sort\.png$/.test(f)) {
        candFiles.push(path.join(shotDir, f));
      }
    }
  }
  for (const f of candFiles) {
    try { extraAttach.push({ filename: path.basename(f), path: f }); } catch (e) {}
  }

  await transporter.sendMail({
    from: `"二级行情汇总" <${emailUser}>`,
    to: emailTo,
    subject: subject,
    html: `
      <h2>债立方 · 二级行情「成交行情」数据汇总（我的关注下框）</h2>
      <p><strong>成交日期：</strong>${dateStr}</p>
      <p><strong>工作簿含 3 个 Sheet：</strong></p>
      <ul>
        <li>Sheet1 成交行情_下框：下框原始全字段</li>
        <li>Sheet2 成交行情汇总：当天固定表头（经纪商模板），发行人/区域来自债券详情页</li>
        <li>Sheet3 历史累积汇总：每日 Sheet2 逐日累积</li>
      </ul>
      <p><strong>数据来源：</strong>债立方 web.innodealing.com 二级行情（我的关注 · 成交行情下框），抓全量不按阈值筛选</p>
      <p><strong>自核对：</strong>脚本排序后已对全表截图并与提取结果逐行比对，核对报告见附件 <code>secondary2_verify_${dateStr}.json</code>，排序后表格截图见 <code>05-after-sort.png</code>，可直接人眼复核是否遗漏。</p>
      <hr>
      <p><small>由 GitHub Actions 自动生成 | ${new Date().toISOString()}</small></p>
    `,
    attachments: [{
      filename: `二级行情成交汇总_${dateStr}.xlsx`,
      path: excelPath,
    }, ...extraAttach],
  });

  console.log(`[邮件] 发送成功: ${emailTo}`);
  return true;
}

// 定位要发送的 Excel：优先当天文件，否则取最新的 secondary_quote_*.xlsx
const secondary2Dir = process.env.SECONDARY2_DIR || path.join(__dirname, 'data', 'secondary2');
let excelPath = process.env.EXCEL_PATH || path.join(secondary2Dir, `secondary_quote_${BJ_DATE}.xlsx`);
if (!fs.existsSync(excelPath)) {
  const cands = fs.readdirSync(secondary2Dir)
    .filter(f => /^secondary_quote_.*\.xlsx$/.test(f))
    .sort().reverse();
  if (cands.length) excelPath = path.join(secondary2Dir, cands[0]);
}

const dateStr = process.env.DATE_STR || BJ_DATE;

if (!fs.existsSync(excelPath)) {
  console.log(`[邮件] Excel 不存在: ${excelPath}，跳过`);
  process.exit(0);
}

sendEmail(excelPath, dateStr).catch(err => {
  console.error('[邮件] 发送失败:', err.message);
  process.exit(1);
});
