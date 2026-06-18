import html
import json
import os
import re
import smtplib
import ssl
import sys
from dataclasses import dataclass
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo


TRENDING_URLS = [
    "https://github.com/trending",
    "https://github.com/trending/python",
    "https://github.com/trending/typescript",
    "https://github.com/trending/javascript",
    "https://github.com/trending/go",
    "https://github.com/trending/rust",
]

AI_KEYWORDS = [
    "ai",
    "agent",
    "agents",
    "llm",
    "rag",
    "mcp",
    "model",
    "models",
    "openai",
    "claude",
    "codex",
    "chatgpt",
    "gpt",
    "transformer",
    "embedding",
    "vector",
    "inference",
    "pytorch",
    "tensorflow",
    "diffusion",
    "multimodal",
    "vision",
    "ocr",
    "speech",
    "tts",
    "machine learning",
    "deep learning",
    "neural",
]

SEARCH_FALLBACK_QUERIES = [
    "topic:artificial-intelligence stars:>1000",
    "topic:llm stars:>1000",
    "topic:ai-agent stars:>100",
    "topic:mcp stars:>50",
    "topic:rag stars:>100",
]


@dataclass
class Repo:
    full_name: str
    url: str
    description: str = ""
    language: str = ""
    stars: int | None = None
    stars_today: int | None = None
    source: str = ""
    topics: tuple[str, ...] = ()
    license_name: str = ""
    pushed_at: str = ""
    estimated: bool = False


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def http_get(url: str, accept: str = "text/html") -> str:
    headers = {
        "Accept": accept,
        "User-Agent": "codex-ai-star-daily/1.0",
    }
    token = os.environ.get("GITHUB_TOKEN")
    if token and "api.github.com" in url:
        headers["Authorization"] = f"Bearer {token}"
        headers["X-GitHub-Api-Version"] = "2022-11-28"

    request = Request(url, headers=headers)
    with urlopen(request, timeout=25) as response:
        return response.read().decode("utf-8", errors="replace")


def clean_text(value: str) -> str:
    value = re.sub(r"<[^>]+>", " ", value)
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def parse_count(value: str) -> int:
    return int(value.replace(",", "").strip())


def parse_trending_page(url: str) -> list[Repo]:
    try:
        page = http_get(url)
    except (HTTPError, URLError, TimeoutError) as exc:
        print(f"Warning: failed to fetch {url}: {exc}", file=sys.stderr)
        return []

    repos: list[Repo] = []
    articles = re.findall(r"<article[^>]*class=\"[^\"]*Box-row[^\"]*\"[^>]*>(.*?)</article>", page, re.S)
    for article in articles:
        link_match = re.search(r"<h2[^>]*>.*?<a[^>]+href=\"/([^\"]+)\"", article, re.S)
        if not link_match:
            continue

        full_name = clean_text(link_match.group(1)).replace(" ", "")
        if full_name.count("/") != 1:
            continue

        desc_match = re.search(r"<p[^>]*class=\"[^\"]*col-9[^\"]*\"[^>]*>(.*?)</p>", article, re.S)
        language_match = re.search(r'itemprop="programmingLanguage"[^>]*>(.*?)</span>', article, re.S)
        stars_match = re.search(r'href="/' + re.escape(full_name) + r'/stargazers"[^>]*>\s*([\d,]+)', article, re.S)
        today_match = re.search(r"([\d,]+)\s+stars?\s+today", clean_text(article), re.I)

        repos.append(
            Repo(
                full_name=full_name,
                url=f"https://github.com/{full_name}",
                description=clean_text(desc_match.group(1)) if desc_match else "",
                language=clean_text(language_match.group(1)) if language_match else "",
                stars=parse_count(stars_match.group(1)) if stars_match else None,
                stars_today=parse_count(today_match.group(1)) if today_match else None,
                source=url,
            )
        )
    return repos


def looks_ai_related(repo: Repo) -> bool:
    text = " ".join([repo.full_name, repo.description, repo.language, " ".join(repo.topics)]).lower()
    for keyword in AI_KEYWORDS:
        if keyword in {"ai", "gpt", "rag", "mcp", "ocr", "tts", "llm"}:
            if re.search(rf"(^|[^a-z0-9]){re.escape(keyword)}([^a-z0-9]|$)", text):
                return True
        elif keyword in text:
            return True
    return False


def github_api_repo(full_name: str) -> dict:
    if not os.environ.get("GITHUB_TOKEN"):
        return {}

    url = f"https://api.github.com/repos/{quote(full_name, safe='/')}"
    try:
        return json.loads(http_get(url, accept="application/vnd.github+json"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        print(f"Warning: failed to enrich {full_name}: {exc}", file=sys.stderr)
        return {}


def github_search_fallback(existing: set[str]) -> list[Repo]:
    results: list[Repo] = []
    for query in SEARCH_FALLBACK_QUERIES:
        url = (
            "https://api.github.com/search/repositories"
            f"?q={quote(query)}&sort=stars&order=desc&per_page=10"
        )
        try:
            data = json.loads(http_get(url, accept="application/vnd.github+json"))
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
            print(f"Warning: fallback search failed for {query}: {exc}", file=sys.stderr)
            continue

        for item in data.get("items", []):
            full_name = item.get("full_name", "")
            if not full_name or full_name in existing:
                continue
            existing.add(full_name)
            results.append(
                Repo(
                    full_name=full_name,
                    url=item.get("html_url") or f"https://github.com/{full_name}",
                    description=item.get("description") or "",
                    language=item.get("language") or "",
                    stars=item.get("stargazers_count"),
                    stars_today=None,
                    source=f"GitHub Search: {query}",
                    topics=tuple(item.get("topics") or ()),
                    license_name=(item.get("license") or {}).get("spdx_id") or "",
                    pushed_at=item.get("pushed_at") or "",
                    estimated=True,
                )
            )
    return results


def collect_repos() -> list[Repo]:
    by_name: dict[str, Repo] = {}
    for url in TRENDING_URLS:
        for repo in parse_trending_page(url):
            if not looks_ai_related(repo):
                continue
            existing = by_name.get(repo.full_name)
            if existing is None or (repo.stars_today or 0) > (existing.stars_today or 0):
                by_name[repo.full_name] = repo

    candidates = list(by_name.values())
    candidates.sort(key=lambda item: (item.stars_today or 0, item.stars or 0), reverse=True)

    enriched: dict[str, Repo] = {}
    for repo in candidates[:20]:
        details = github_api_repo(repo.full_name)
        topics = tuple(details.get("topics") or ())
        merged = Repo(
            full_name=repo.full_name,
            url=repo.url,
            description=details.get("description") or repo.description,
            language=details.get("language") or repo.language,
            stars=details.get("stargazers_count") or repo.stars,
            stars_today=repo.stars_today,
            source=repo.source,
            topics=topics,
            license_name=(details.get("license") or {}).get("spdx_id") or "",
            pushed_at=details.get("pushed_at") or "",
        )
        if looks_ai_related(merged):
            enriched[merged.full_name] = merged

    if len(enriched) < 10:
        for repo in github_search_fallback(set(enriched)):
            if looks_ai_related(repo):
                enriched[repo.full_name] = repo
            if len(enriched) >= 14:
                break

    repos = list(enriched.values())
    repos.sort(key=lambda item: (item.stars_today is not None, item.stars_today or 0, item.stars or 0), reverse=True)
    return repos[:10]


def codex_help(repo: Repo) -> tuple[str, list[str], list[str]]:
    text = " ".join([repo.full_name, repo.description, repo.language, " ".join(repo.topics)]).lower()
    reasons: list[str] = []
    actions: list[str] = []

    if any(word in text for word in ["mcp", "model context protocol"]):
        reasons.append("可能直接扩展 Codex 工具边界，尤其适合作为 MCP Server 或 MCP 参考实现。")
        actions.append("优先查看 README 的安装方式，评估是否能接入为本地 MCP 工具。")
    if any(word in text for word in ["agent", "agents", "skill", "skills", "workflow"]):
        reasons.append("可能提供 agent skills、任务模板或多步骤工作流设计，可转化为 Codex skill。")
        actions.append("抽取其目录结构、prompt 模板和任务编排方式，沉淀为 Codex skill 草案。")
    if any(word in text for word in ["code", "codebase", "repository", "index", "graph"]):
        reasons.append("可能增强代码库理解、索引、检索或跨文件记忆。")
        actions.append("用一个小型仓库试跑，比较它对代码搜索、依赖图和上下文召回的帮助。")
    if any(word in text for word in ["rag", "retrieval", "vector", "embedding", "search"]):
        reasons.append("可能改善 Codex 的文档检索、知识库问答或长期上下文。")
        actions.append("评估是否适合接入本地文档、项目 README 和历史决策记录。")
    if any(word in text for word in ["ocr", "pdf", "document", "vision", "multimodal", "speech", "tts", "voice"]):
        reasons.append("可能扩展 Codex 对文档、图片、语音或多模态资料的处理能力。")
        actions.append("挑一个真实资料样本测试输入输出质量和自动化成本。")

    if reasons:
        priority = "高" if len(reasons) >= 2 or any("MCP" in reason or "代码库" in reason for reason in reasons) else "中"
    else:
        priority = "低"
        reasons.append("更偏通用 AI 项目或学习资料，对 Codex 的直接接入价值暂不明显。")
        actions.append("先收藏观察，等 README、示例和生态采用度更清晰后再评估。")

    if repo.license_name:
        actions.append(f"落地前检查许可证：{repo.license_name}。")
    else:
        actions.append("落地前补查许可证和维护活跃度。")

    return priority, reasons, actions[:3]


def format_number(value: int | None) -> str:
    return f"{value:,}" if value is not None else "未知"


def chinese_description(repo: Repo) -> str:
    name = repo.full_name.lower()
    text = " ".join([name, repo.description.lower(), " ".join(repo.topics)]).lower()

    if "mcp" in text or "model context protocol" in text:
        if any(word in text for word in ["code", "codebase", "repository", "memory", "graph"]):
            return "面向代码库理解的 MCP 工具，可帮助 AI 助手建立代码索引、记忆和检索能力。"
        if any(word in text for word in ["browser", "chrome", "devtools", "web"]):
            return "面向浏览器或开发者工具的 MCP 服务，可把网页和调试能力接入 AI 工作流。"
        return "基于 MCP 协议的 AI 工具服务，可扩展 Codex 或其他智能体的外部能力。"

    if any(word in text for word in ["skill", "skills"]):
        return "面向 AI 智能体的技能集合或技能框架，可复用任务模板、提示词和工作流结构。"

    if any(word in text for word in ["agent", "agents", "agentic"]):
        if any(word in text for word in ["social", "twitter", "reddit", "youtube", "bilibili", "xiaohongshu"]):
            return "为 AI 智能体提供跨平台信息读取和搜索能力，适合增强外部资料获取。"
        return "面向 AI 智能体的开发框架或工具，帮助构建多步骤自动化任务。"

    if any(word in text for word in ["rag", "retrieval", "vector", "embedding"]):
        return "面向 RAG、向量检索或知识库问答的项目，可改善文档检索和上下文召回。"

    if any(word in text for word in ["ocr", "pdf", "document"]):
        return "面向文档识别和结构化处理的工具，可把 PDF、图片等资料转成可分析内容。"

    if any(word in text for word in ["voice", "speech", "tts", "audio"]):
        return "面向语音、音频或文本转语音的 AI 项目，可扩展多模态输入输出能力。"

    if any(word in text for word in ["vision", "image", "multimodal", "detector", "detection"]):
        return "面向视觉理解或多模态处理的 AI 项目，可用于图片识别、检测或视觉自动化。"

    if any(word in text for word in ["time series", "forecast", "forecasting"]):
        return "面向时间序列预测的 AI 模型或工具，适合业务指标、趋势和预测类场景。"

    if any(word in text for word in ["engineering", "from scratch", "tutorial", "course", "learn"]):
        return "面向 AI 工程实践的学习资料或教程，适合提炼项目模板和实现路线。"

    if any(word in text for word in ["claude", "codex", "openai", "chatgpt", "gpt"]):
        return "围绕主流 AI 编程助手或大模型生态的工具，可作为 Codex 工作流参考。"

    if any(word in text for word in ["llm", "model", "inference", "transformer"]):
        return "面向大模型、推理或模型应用的开源项目，可作为 AI 工程能力参考。"

    return "AI 相关开源项目，建议结合 README、示例和许可证进一步评估实际价值。"


def build_report(repos: list[Repo]) -> tuple[str, str]:
    now = datetime.now(ZoneInfo("Asia/Shanghai"))
    date_text = now.strftime("%Y-%m-%d")
    subject = f"AI GitHub Star 增长日报 - {date_text}"

    plain_lines = [
        f"红豆，今天的 AI GitHub Star 增长 Top 10 来了。",
        "",
        f"采集时间：{now.strftime('%Y-%m-%d %H:%M:%S')} Asia/Shanghai",
        "数据口径：优先使用 GitHub Trending 的 stars today；不足 10 个时用 GitHub Search 兜底并标注为估算。",
        "",
        "Top 10 项目榜单",
    ]

    html_rows = []
    priority_groups: dict[str, list[tuple[Repo, list[str], list[str]]]] = {"高": [], "中": [], "低": []}

    for index, repo in enumerate(repos, start=1):
        stars_today = format_number(repo.stars_today)
        growth_note = "估算/无精确 24h 增量" if repo.estimated or repo.stars_today is None else stars_today
        priority, reasons, actions = codex_help(repo)
        priority_groups[priority].append((repo, reasons, actions))

        plain_lines.extend(
            [
                f"{index}. {repo.full_name}",
                f"   链接：{repo.url}",
                f"   今日新增 Star：{growth_note}",
                f"   当前 Star：{format_number(repo.stars)}",
                f"   技术栈：{repo.language or '未知'}",
                f"   中文简介：{chinese_description(repo)}",
                f"   来源：{repo.source}",
                "",
            ]
        )

        html_rows.append(
            "<tr>"
            f"<td>{index}</td>"
            f"<td><a href=\"{html.escape(repo.url)}\">{html.escape(repo.full_name)}</a></td>"
            f"<td>{html.escape(growth_note)}</td>"
            f"<td>{format_number(repo.stars)}</td>"
            f"<td>{html.escape(repo.language or '未知')}</td>"
            f"<td>{html.escape(chinese_description(repo))}</td>"
            f"<td>{html.escape(repo.source)}</td>"
            "</tr>"
        )

    plain_lines.extend(["Codex 帮助度分析", ""])
    analysis_html = []
    for priority in ("高", "中", "低"):
        items = priority_groups[priority]
        if not items:
            continue
        plain_lines.append(f"{priority}优先级")
        analysis_html.append(f"<h3>{priority}优先级</h3><ul>")
        for repo, reasons, actions in items:
            plain_lines.append(f"- {repo.full_name}")
            plain_lines.append(f"  帮助方式：{'；'.join(reasons)}")
            plain_lines.append(f"  落地建议：{'；'.join(actions)}")
            analysis_html.append(
                "<li>"
                f"<strong><a href=\"{html.escape(repo.url)}\">{html.escape(repo.full_name)}</a></strong>"
                f"<br><b>帮助方式：</b>{html.escape('；'.join(reasons))}"
                f"<br><b>落地建议：</b>{html.escape('；'.join(actions))}"
                "</li>"
            )
        analysis_html.append("</ul>")
        plain_lines.append("")

    sources = sorted({repo.source for repo in repos})
    plain_lines.extend(["数据来源", *[f"- {source}" for source in sources]])

    html_body = f"""
<!doctype html>
<html>
  <body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;line-height:1.55;color:#1f2933;">
    <p>红豆，今天的 AI GitHub Star 增长 Top 10 来了。</p>
    <p><b>采集时间：</b>{html.escape(now.strftime('%Y-%m-%d %H:%M:%S'))} Asia/Shanghai<br>
    <b>数据口径：</b>优先使用 GitHub Trending 的 stars today；不足 10 个时用 GitHub Search 兜底并标注为估算。</p>
    <h2>Top 10 项目榜单</h2>
    <table cellpadding="8" cellspacing="0" border="1" style="border-collapse:collapse;border-color:#d7dee8;">
      <thead>
        <tr>
          <th>排名</th><th>项目</th><th>今日新增 Star</th><th>当前 Star</th><th>技术栈</th><th>中文简介</th><th>来源</th>
        </tr>
      </thead>
      <tbody>
        {''.join(html_rows)}
      </tbody>
    </table>
    <h2>Codex 帮助度分析</h2>
    {''.join(analysis_html)}
    <h2>数据来源</h2>
    <ul>{''.join(f'<li>{html.escape(source)}</li>' for source in sources)}</ul>
  </body>
</html>
"""
    return subject, "\n".join(plain_lines), html_body


def send_email(subject: str, plain_body: str, html_body: str) -> None:
    gmail_user = require_env("GMAIL_USER")
    gmail_password = require_env("GMAIL_APP_PASSWORD")
    email_to = require_env("EMAIL_TO")

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = gmail_user
    message["To"] = email_to
    message.attach(MIMEText(plain_body, "plain", "utf-8"))
    message.attach(MIMEText(html_body, "html", "utf-8"))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, [email_to], message.as_string())


def main() -> None:
    repos = collect_repos()
    if not repos:
        raise SystemExit("No AI repositories collected; aborting email to avoid sending an empty report.")

    subject, plain_body, html_body = build_report(repos)
    send_email(subject, plain_body, html_body)
    print(f"Sent report: {subject} ({len(repos)} repos)")


if __name__ == "__main__":
    main()
