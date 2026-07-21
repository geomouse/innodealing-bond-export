#!/usr/bin/env node
// 发送二级行情成交数据汇总邮件（带 4-Sheet Excel 附件）
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

  const subject = `[二级行情成交汇总] ${dateStr} | 4-Sheet(下框/上框/当天/历史累积)`;

  await transporter.sendMail({
    from: `"二级行情汇总" <${emailUser}>`,
    to: emailTo,
    subject: subject,
    html: `
      <h2>债立方 · 二级行情「我的关注」成交数据汇总</h2>
      <p><strong>成交日期：</strong>${dateStr}</p>
      <p><strong>工作簿含 4 个 Sheet：</strong></p>
      <ul>
        <li>Sheet1 成交行情(下框)：最新成交等 13 列</li>
        <li>Sheet2 我的关注(上框)：39 列全字段</li>
        <li>Sheet3 当天汇总：上框为主 + 详情页区域/发行人</li>
        <li>Sheet4 历史累积：每日 Sheet3 逐日追加</li>
      </ul>
      <p><strong>数据来源：</strong>债立方 web.innodealing.com（仅含 ≥2.0 成交量债券）</p>
      <hr>
      <p><small>由 GitHub Actions 自动生成 | ${new Date().toISOString()}</small></p>
    `,
    attachments: [{
      filename: `二级行情成交汇总_${dateStr}.xlsx`,
      path: excelPath,
    }],
  });

  console.log(`[邮件] 发送成功: ${emailTo}`);
  return true;
}

// 定位要发送的 Excel：优先当天文件，否则取最新的 secondary_quote_*.xlsx
const secondaryDir = process.env.SECONDARY_DIR || path.join(__dirname, 'data', 'secondary');
let excelPath = process.env.EXCEL_PATH || path.join(secondaryDir, `secondary_quote_${BJ_DATE}.xlsx`);
if (!fs.existsSync(excelPath)) {
  const cands = fs.readdirSync(secondaryDir)
    .filter(f => /^secondary_quote_.*\.xlsx$/.test(f))
    .sort().reverse();
  if (cands.length) excelPath = path.join(secondaryDir, cands[0]);
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
