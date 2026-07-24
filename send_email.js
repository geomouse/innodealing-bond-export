#!/usr/bin/env node
// 发送汇总邮件（带附件）
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

async function sendEmail(excelPath, dateStr, stats) {
  const emailUser = process.env.EMAIL_USER;
  const emailPass = process.env.EMAIL_PASS;
  const emailTo = process.env.EMAIL_TO;

  if (!emailUser || !emailPass || !emailTo) {
    console.log('邮件未配置，跳过发送');
    return;
  }

  console.log(`[邮件] 发送到 ${emailTo}...`);

  const transporter = nodemailer.createTransport({
    host: 'smtp.qq.com',
    port: 465,
    secure: true,
    auth: { user: emailUser, pass: emailPass },
  });

  const [todayCount, totalCount, rateFilled, rateTotal] = stats;
  const ratePct = rateTotal > 0 ? (rateFilled / rateTotal * 100).toFixed(0) : '0';

  let subject = `[信用债发行汇总] ${dateStr} | 今日${todayCount}条 | 累计${totalCount}只`;
  if (process.env.RETRY === '1') {
    subject = `[补录重试] ${subject}`;
  }

  await transporter.sendMail({
    from: `"信用债汇总" <${emailUser}>`,
    to: emailTo,
    subject: subject,
    html: `
      <h2>信用债发行数据汇总</h2>
      <p><strong>日期：</strong>${dateStr}</p>
      <p><strong>今日新增：</strong>${todayCount} 条</p>
      <p><strong>累计债券：</strong>${totalCount} 只</p>
      <p><strong>票面利率补全：</strong>${rateFilled}/${rateTotal} (${ratePct}%)</p>
      <p><strong>数据来源：</strong>债立方 web.innodealing.com</p>
      <hr>
      <p><small>由 GitHub Actions 自动生成 | ${new Date().toISOString()}</small></p>
    `,
    attachments: [{
      filename: `信用债发行汇总_${dateStr}.xlsx`,
      path: excelPath,
    }],
  });

  console.log(`[邮件] 发送成功: ${emailTo}`);
}

// 从命令行参数或环境变量获取信息
const dateStr = process.env.DATE_STR || new Date().toISOString().split('T')[0];
const excelPath = process.env.EXCEL_PATH || path.join(__dirname, 'data', `信用债发行汇总_${dateStr}.xlsx`);

// 解析统计信息
const statsRaw = process.env.STATS || '0:0:0:0';
const stats = statsRaw.split(':');

if (!fs.existsSync(excelPath)) {
  console.log(`[邮件] Excel 文件不存在: ${excelPath}，跳过邮件发送`);
  process.exit(0);
}

sendEmail(excelPath, dateStr, stats).catch(err => {
  console.error('[邮件] 发送失败:', err.message);
  process.exit(1);
});
