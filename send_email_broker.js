#!/usr/bin/env node
// 发送经纪商行情（成交行情下框）汇总邮件（带 3-Sheet Excel 附件）
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

  const subject = `[经纪商行情成交汇总] ${dateStr} | 3-Sheet(下框源数据/当天汇总/历史累积)`;

  await transporter.sendMail({
    from: `"经纪商行情汇总" <${emailUser}>`,
    to: emailTo,
    subject: subject,
    html: `
      <h2>债立方 · 经纪商行情「成交行情」数据汇总</h2>
      <p><strong>成交日期：</strong>${dateStr}</p>
      <p><strong>工作簿含 3 个 Sheet：</strong></p>
      <ul>
        <li>Sheet1 成交行情_下框：下框原始全字段</li>
        <li>Sheet2 成交行情汇总：当天固定表头，发行人/区域来自债券详情页</li>
        <li>Sheet3 历史累积汇总：每日 Sheet2 逐日累积</li>
      </ul>
      <p><strong>数据来源：</strong>债立方 web.innodealing.com 经纪商行情（仅含平均成交 ≥2.8 的债券，上午放宽至 ≥2.4）</p>
      <hr>
      <p><small>由 GitHub Actions 自动生成 | ${new Date().toISOString()}</small></p>
    `,
    attachments: [{
      filename: `经纪商行情成交汇总_${dateStr}.xlsx`,
      path: excelPath,
    }],
  });

  console.log(`[邮件] 发送成功: ${emailTo}`);
  return true;
}

// 定位要发送的 Excel：优先当天文件，否则取最新的 broker_quote_*.xlsx
const brokerDir = process.env.BROKER_DIR || path.join(__dirname, 'data', 'broker');
let excelPath = process.env.EXCEL_PATH || path.join(brokerDir, `broker_quote_${BJ_DATE}.xlsx`);
if (!fs.existsSync(excelPath)) {
  const cands = fs.readdirSync(brokerDir)
    .filter(f => /^broker_quote_.*\.xlsx$/.test(f))
    .sort().reverse();
  if (cands.length) excelPath = path.join(brokerDir, cands[0]);
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
