# AI GitHub Star Daily

这个 GitHub Actions workflow 会每天北京时间 09:00 抓取 AI 相关 GitHub 项目的 Star 增长榜，并通过 Gmail 发送中文日报。

## 需要配置的 Secrets

在 GitHub 仓库里打开 `Settings -> Secrets and variables -> Actions -> New repository secret`，新增：

- `GMAIL_USER`：发信 Gmail 地址
- `GMAIL_APP_PASSWORD`：Gmail 应用专用密码
- `EMAIL_TO`：收件邮箱

`GITHUB_TOKEN` 不需要手动配置，GitHub Actions 会自动提供。

## Gmail 应用专用密码

Gmail SMTP 需要应用专用密码。账号需要先开启两步验证，然后在 Google 账号安全设置里创建 App Password。

## 手动测试

配置 Secrets 后，可以在 GitHub 仓库的 `Actions -> AI GitHub Star Daily -> Run workflow` 手动触发一次。

## 数据口径

脚本优先抓取 GitHub Trending 页面中的 `stars today` 指标；如果当天 AI 项目不足 10 个，会用 GitHub Search API 按 AI 相关 topic 兜底，并在邮件中标注为估算或无精确 24h 增量。

