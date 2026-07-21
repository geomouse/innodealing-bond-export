# 本地兜底运行器：债立方「我的关注」二级行情成交汇总
# 设计：
#  - 先 git pull 获取云端已产出的历史/今日数据（与云端同步）
#  - 若今日 daily json 已存在（说明云端 GitHub 或 cron 已成功产出）→ 不重复发邮件，仅本地再生 xlsx
#  - 若今日 json 不存在（云端当天未成功）→ 本地兜底产出数据并发送邮件
#  - 最后把本地新增的 daily json 提交回仓库，保证历史不丢、下次云端/本地都能读到
$ErrorActionPreference = 'Stop'
$repo = 'C:\Users\geo\WorkBuddy\2026-07-17-11-02-38\github_package'
Set-Location $repo

# 1) 同步云端最新数据
git pull --rebase 2>&1 | Out-Null

# 2) 计算北京时间日期
$bj = (Get-Date).ToUniversalTime().AddHours(8).ToString('yyyy-MM-dd')
$jsonPath = "data/secondary/daily/$bj.json"

# 3) 去重判断：云端今日是否已产出
$send = '1'
if (Test-Path $jsonPath) { $send = '0' }

# 4) 运行导出（本地用 headless + msedge：无弹窗且稳定；自带 chromium-headless-shell 在本机会崩溃）
$env:HEADLESS = 'true'
$env:BROWSER_CHANNEL = 'msedge'
$env:SECONDARY_DIR = "$repo\data\secondary"
Write-Host "[本地兜底] 日期=$bj 云端已产出=$($send -eq '0') -> 发邮件=$($send -eq '1')"
node innodealing_secondary.js

# 5) 邮件（仅当本地为今日首个产出方时发送，避免与云端双邮件）
if ($send -eq '1') {
  node send_email_secondary.js
}

# 6) 把今日新增的历史 json 提交回仓库（与云端同步）
git add data/secondary
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) {
  git -c user.name="local-fallback" -c user.email="local@fallback" commit -m "二级行情 $bj (本地兜底)"
  git push 2>&1 | Out-Null
  Write-Host "[本地兜底] 已提交并推送今日数据"
} else {
  Write-Host "[本地兜底] 无新增数据，跳过推送"
}
