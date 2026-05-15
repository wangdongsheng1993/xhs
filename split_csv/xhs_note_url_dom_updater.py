import argparse
import csv
import os
import random
import re
import shutil
import subprocess
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Callable
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_PROFILE_DIR = BASE_DIR / "chrome_xhs_profile"
DEFAULT_PORT = 9222
NOTE_URL_HEADER = "笔记官方地址"
TITLE_HEADER = "笔记标题"
HOMEPAGE_HEADER = "主页链接"
NO_TITLE_VALUE = "笔记暂未设置标题"
INTERVAL_JITTER_SEC = 20
MANUAL_RECOVERY_TIMEOUT_SEC = 300


def log(message: str) -> None:
    print(message, flush=True)


def format_duration(seconds: float) -> str:
    total_seconds = max(0, int(round(seconds)))
    minutes, secs = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours}小时{minutes}分{secs}秒"
    if minutes:
        return f"{minutes}分{secs}秒"
    return f"{secs}秒"


def wait_with_stop(total_seconds: float, stop_event: Event | None) -> bool:
    deadline = time.time() + max(0.0, total_seconds)
    while time.time() < deadline:
        if stop_event and stop_event.is_set():
            return False
        time.sleep(min(0.5, max(0.0, deadline - time.time())))
    return True


def normalize_text(value: str) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[\s\u200b\u200c\u200d\ufeff]+", "", text)


def normalize_url(url: str) -> str:
    value = str(url or "").strip()
    if not value:
        return ""
    if value.startswith("//"):
        return "https:" + value
    if value.startswith("/"):
        return urljoin("https://www.xiaohongshu.com", value)
    if not re.match(r"^https?://", value, flags=re.I):
        return "https://" + value
    return value


def read_csv_rows(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with path.open("r", encoding=encoding, newline="") as file:
                reader = csv.DictReader(file)
                return list(reader.fieldnames or []), list(reader)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(f"无法识别 CSV 编码: {path}") from last_error


def has_valid_xsec_token(value: str) -> bool:
    return bool(build_standard_note_url(value))


def find_resume_anchor(rows: list[dict[str, str]]) -> int:
    for index in range(len(rows) - 1, -1, -1):
        value = str(rows[index].get(NOTE_URL_HEADER, "") or "").strip()
        if has_valid_xsec_token(value):
            return index + 1
    return 0


def read_resume_rows(source: Path, output: Path) -> tuple[list[str], list[dict[str, str]], int]:
    if output.exists():
        source_fieldnames, source_rows = read_csv_rows(source)
        output_fieldnames, output_rows = read_csv_rows(output)
        if source_fieldnames != output_fieldnames:
            return source_fieldnames, source_rows, 0
        if len(output_rows) > len(source_rows):
            return source_fieldnames, source_rows, 0
        resumed_rows = [dict(row) for row in source_rows]
        for index, output_row in enumerate(output_rows):
            resumed_rows[index].update(output_row)
        return source_fieldnames, resumed_rows, find_resume_anchor(resumed_rows)
    fieldnames, rows = read_csv_rows(source)
    return fieldnames, rows, 0


def summarize_resume_state(rows: list[dict[str, str]], resume_anchor: int) -> tuple[int, int, int]:
    completed = 0
    failed_before_anchor = 0
    pending = 0

    for index, row in enumerate(rows):
        title = str(row.get(TITLE_HEADER, "") or "").strip()
        homepage = str(row.get(HOMEPAGE_HEADER, "") or "").strip()
        note_url = str(row.get(NOTE_URL_HEADER, "") or "").strip()

        if has_valid_xsec_token(note_url):
            completed += 1
            continue

        can_process = bool(title and homepage and title != NO_TITLE_VALUE)
        if not can_process:
            continue

        if index < resume_anchor:
            failed_before_anchor += 1
        else:
            pending += 1

    return completed, failed_before_anchor, pending


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def output_path_for(source: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_笔记地址已更新_{timestamp}{source.suffix}")


def failed_path_for(output_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_path.with_name(f"{output_path.stem}_失败数据_{timestamp}.csv")


def find_browser_executable() -> str:
    candidates = []
    env_browser = os.environ.get("XHS_BROWSER_PATH")
    if env_browser:
        candidates.append(env_browser)

    if os.name == "nt":
        local_app = os.environ.get("LOCALAPPDATA", "")
        program_files = os.environ.get("PROGRAMFILES", r"C:\Program Files")
        program_files_x86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
        candidates.extend(
            [
                rf"{program_files}\Google\Chrome\Application\chrome.exe",
                rf"{program_files_x86}\Google\Chrome\Application\chrome.exe",
                rf"{local_app}\Google\Chrome\Application\chrome.exe",
                rf"{program_files}\Microsoft\Edge\Application\msedge.exe",
                rf"{program_files_x86}\Microsoft\Edge\Application\msedge.exe",
            ]
        )
    else:
        for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge"):
            found = shutil.which(name)
            if found:
                candidates.append(found)

    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(candidate)
    raise RuntimeError("未找到本机 Chrome/Edge。可设置环境变量 XHS_BROWSER_PATH 指向浏览器可执行文件。")


def is_cdp_ready(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1.5) as response:
            return response.status == 200
    except Exception:
        return False


def launch_browser(
    port: int,
    profile_dir: Path,
    start_url: str = "https://www.xiaohongshu.com/explore",
    logger: Callable[[str], None] = log,
) -> None:
    if is_cdp_ready(port):
        logger(f"浏览器调试端口已可用: {port}")
        return

    browser = find_browser_executable()
    profile_dir.mkdir(parents=True, exist_ok=True)
    command = [
        browser,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        start_url,
    ]
    logger("启动本机浏览器: " + browser)
    subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    deadline = time.time() + 15
    while time.time() < deadline:
        if is_cdp_ready(port):
            logger(f"浏览器已启动，调试端口: {port}")
            return
        time.sleep(0.5)
    raise RuntimeError(f"浏览器已启动但调试端口 {port} 未就绪，请确认 Chrome/Edge 没有限制远程调试。")


def click_notes_tab_if_visible(page) -> None:
    selectors = [
        "[role='tab']",
        ".reds-tabs-list .tab",
        ".user-tabs .tab",
        ".tabs-list .tab",
        "button:has-text('笔记')",
        "a:has-text('笔记')",
        "text=笔记",
    ]
    for selector in selectors:
        try:
            locator = page.locator(selector).filter(has_text="笔记").first
            if locator.count() and locator.is_visible(timeout=500):
                locator.click(timeout=1500)
                page.wait_for_timeout(800)
                return
        except Exception:
            continue


def wait_for_notes_dom_ready(page, timeout_ms: int = 15000) -> bool:
    try:
        page.wait_for_function(
            """
            () => {
                const cards = Array.from(document.querySelectorAll(
                    'section.note-item, .note-item, [class*="note-item"], [class*="note-card"]'
                ));
                if (!cards.length) return false;
                return cards.some((card) => {
                    const titleNode = card.querySelector('a.title, [class*="title"]');
                    const titleText = titleNode ? (titleNode.innerText || titleNode.textContent || '') : '';
                    const cardText = card.innerText || card.textContent || '';
                    const profileLink = card.querySelector('a[href*="/user/profile/"]');
                    const exploreLink = card.querySelector('a[href*="/explore/"]');
                    const coverLink = card.querySelector('a.cover');
                    return Boolean(
                        titleText.trim()
                        || cardText.trim()
                        || profileLink
                        || exploreLink
                        || coverLink
                    );
                });
            }
            """,
            timeout=timeout_ms,
        )
        return True
    except PlaywrightTimeoutError:
        return False


def wait_for_manual_recovery(
    page,
    timeout_sec: int,
    stop_event: Event | None,
    logger: Callable[[str], None],
    wait_reason: str,
) -> str:
    logger(f"  - {wait_reason}")
    logger(f"  - 将在当前页等待最多 {format_duration(timeout_sec)}，恢复到笔记卡片页面后继续当前条。")
    deadline = time.time() + max(0, timeout_sec)
    while time.time() < deadline:
        if stop_event and stop_event.is_set():
            return "stopped"
        gate_status, _ = detect_page_gate_status(page)
        if gate_status == "none" and wait_for_notes_dom_ready(page, timeout_ms=1200):
            return "recovered"
        time.sleep(1)
    return "timeout"


def build_standard_note_url(href: str) -> str:
    value = normalize_url(href).replace("&amp;", "&")
    if "xsec_token=" not in value:
        return ""

    parts = urlsplit(value)
    path_parts = [part for part in parts.path.split("/") if part]
    note_id = ""
    if len(path_parts) >= 2 and path_parts[0] == "explore":
        note_id = path_parts[1]
    elif len(path_parts) >= 4 and path_parts[0] == "user" and path_parts[1] == "profile":
        note_id = path_parts[3]

    if not note_id or len(note_id) < 8:
        return ""

    token = ""
    source = "pc_user"
    for item in parts.query.split("&"):
        if item.startswith("xsec_token="):
            token = item.split("=", 1)[1]
        elif item.startswith("xsec_source="):
            source = item.split("=", 1)[1] or source
    if not token:
        return ""
    return f"https://www.xiaohongshu.com/explore/{note_id}?xsec_token={token}&xsec_source={source}"


def extract_note_url_from_homepage(page, title: str, max_scrolls: int) -> tuple[str, str]:
    target = normalize_text(title)
    if not target:
        return "", "标题为空"

    result = page.evaluate(
        """
        ({ target }) => {
            const normalize = (value) => String(value || '')
                .trim()
                .toLowerCase()
                .replace(/[\\s\\u200b\\u200c\\u200d\\ufeff]+/g, '');
            const scoreText = (text) => {
                const value = normalize(text);
                if (!value) return 0;
                if (value === target) return 100;
                if (value.includes(target)) return 90;
                if (target.includes(value) && value.length >= Math.min(10, target.length)) return 70;
                return 0;
            };
            const cleanHref = (href) => {
                if (!href) return '';
                try {
                    return new URL(href, location.href).href;
                } catch (error) {
                    return href;
                }
            };
            const cards = Array.from(document.querySelectorAll(
                'section.note-item, .note-item, [class*="note-item"], [class*="note-card"]'
            ));
            const candidates = [];
            for (const card of cards) {
                const titleEl = card.querySelector('a.title, [class*="title"]');
                const titleText = titleEl ? (titleEl.innerText || titleEl.textContent || '') : (card.innerText || card.textContent || '');
                const score = scoreText(titleText);
                if (!score) continue;
                const links = [
                    card.querySelector('a.title'),
                    card.querySelector('a.cover'),
                    ...Array.from(card.querySelectorAll('a[href*="/user/profile/"]')),
                    ...Array.from(card.querySelectorAll('a[href*="/explore/"]')),
                ].filter(Boolean);
                const hrefs = [];
                for (const link of links) {
                    const href = cleanHref(link.getAttribute('href') || link.href || '');
                    if (href) hrefs.push(href);
                }
                candidates.push({
                    score,
                    titleText,
                    hrefs,
                });
            }
            candidates.sort((left, right) => right.score - left.score);
            const best = candidates[0];
            if (!best) {
                return {
                    status: 'title_not_found',
                    href: '',
                    matchedTitle: '',
                    cardCount: cards.length,
                };
            }
            const hrefWithToken = best.hrefs.find(
                (href) => href.includes('xsec_token=') && !href.includes('xsec_token=&')
            ) || '';
            return {
                status: hrefWithToken ? 'ok' : 'token_missing',
                href: hrefWithToken,
                matchedTitle: best.titleText || '',
                cardCount: cards.length,
            };
        }
        """,
        {"target": target},
    )
    href = result.get("href") if result else ""
    if href:
        standard_url = build_standard_note_url(href)
        if standard_url:
            return standard_url, ""
        return "", f"匹配到卡片但链接无法转换: {href[:120]}"

    if (result or {}).get("status") == "token_missing":
        return "", "找到标题匹配卡片但链接没有 xsec_token"
    return "", "未找到标题匹配卡片"


def detect_page_gate_status(page) -> tuple[str, str]:
    try:
        result = page.evaluate(
            """
            () => {
                const text = document.body ? document.body.innerText || '' : '';
                const rateLimitKeywords = [
                    '访问频繁',
                    '操作频繁',
                    '频次异常',
                    '稍后再试',
                    '安全验证',
                    '二次校验',
                ];
                const loginKeywords = [
                    '登录即可查看',
                    '扫码验证身份',
                    '请先登录',
                ];
                const verifyKeywords = [
                    '请完成验证',
                    '请通过验证',
                ];
                const matchedRate = rateLimitKeywords.find((item) => text.includes(item)) || '';
                if (matchedRate) {
                    return { status: 'rate_limited', keyword: matchedRate };
                }
                const matchedLogin = loginKeywords.find((item) => text.includes(item)) || '';
                if (matchedLogin || (text.includes('手机号登录') && text.includes('获取验证码'))) {
                    return { status: 'login_or_verify', keyword: matchedLogin || '手机号登录' };
                }
                const matchedVerify = verifyKeywords.find((item) => text.includes(item)) || '';
                if (matchedVerify) {
                    return { status: 'login_or_verify', keyword: matchedVerify };
                }
                return { status: 'none', keyword: '' };
            }
            """
        )
        return str(result.get("status") or "none"), str(result.get("keyword") or "")
    except Exception:
        return "none", ""


def compute_interval_delay(interval_sec: int) -> int:
    base = max(0, int(interval_sec))
    return random.randint(base, base + INTERVAL_JITTER_SEC)


def update_note_urls(
    source: Path,
    output: Path | None,
    port: int,
    profile_dir: Path,
    skip_existing_xsec: bool,
    interval_sec: int,
    max_scrolls: int,
    logger: Callable[[str], None] = log,
    stop_event: Event | None = None,
) -> tuple[Path, Path | None]:
    source = source.resolve()
    if not source.exists():
        raise FileNotFoundError(f"源 CSV 不存在: {source}")
    output_path = output.resolve() if output else output_path_for(source)
    failed_path = failed_path_for(output_path)
    fieldnames, rows, resume_anchor = read_resume_rows(source, output_path)
    required = [NOTE_URL_HEADER, TITLE_HEADER, HOMEPAGE_HEADER]
    missing = [name for name in required if name not in fieldnames]
    if missing:
        raise ValueError("CSV 缺少列: " + "、".join(missing))

    failed_rows: list[dict[str, str]] = []

    launch_browser(port, profile_dir, logger=logger)
    endpoint = f"http://127.0.0.1:{port}"

    updated = 0
    skipped = 0
    failed = 0
    processed_homepages = 0
    stopped = False
    total_started_at = time.monotonic()
    if output_path.exists():
        completed_count, failed_before_anchor_count, pending_count = summarize_resume_state(rows, resume_anchor)
        if resume_anchor > 0:
            logger("检测到已有更新结果，将从最后一个有效 xsec_token 的下一行继续。")
            logger(f"恢复锚点: 第 {resume_anchor + 2} 行；本次从第 {resume_anchor + 3} 行开始。")
        else:
            logger("检测到已有更新结果，但未检测到有效 xsec_token，将从首条开始处理。")
        logger(
            f"续跑摘要: 已完成 {completed_count} 条，锚点前失败 {failed_before_anchor_count} 条，待处理 {pending_count} 条。"
        )

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(endpoint)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()

        for index, row in enumerate(rows, start=2):
            if index - 2 < resume_anchor:
                skipped += 1
                continue
            if stop_event and stop_event.is_set():
                stopped = True
                logger("收到停止请求，保存已处理数据后退出。")
                break

            title = str(row.get(TITLE_HEADER, "") or "").strip()
            homepage = str(row.get(HOMEPAGE_HEADER, "") or "").strip()
            old_url = str(row.get(NOTE_URL_HEADER, "") or "").strip()

            if not title or not homepage or title == NO_TITLE_VALUE:
                skipped += 1
                continue
            if skip_existing_xsec and has_valid_xsec_token(old_url):
                skipped += 1
                continue

            if processed_homepages > 0 and interval_sec > 0:
                delay_seconds = compute_interval_delay(interval_sec)
                if delay_seconds > 0:
                    logger(f"等待 {delay_seconds} 秒后处理下一条...")
                if not wait_with_stop(delay_seconds, stop_event):
                    stopped = True
                    logger("等待期间收到停止请求，保存已处理数据后退出。")
                    break
            processed_homepages += 1

            logger(f"[{index - 1}/{len(rows)}] 第 {index} 行: {title[:60]}")
            row_started_at = time.monotonic()
            try:
                page.goto(normalize_url(homepage), wait_until="domcontentloaded", timeout=45000)
                gate_status, gate_keyword = detect_page_gate_status(page)
                if gate_status == "rate_limited":
                    if gate_keyword:
                        logger(f"  - 命中页面提示: {gate_keyword}")
                    recovery = wait_for_manual_recovery(
                        page,
                        timeout_sec=MANUAL_RECOVERY_TIMEOUT_SEC,
                        stop_event=stop_event,
                        logger=logger,
                        wait_reason="检测到请求过于频繁页面，请手动刷新或等待恢复。",
                    )
                    if recovery == "stopped":
                        stopped = True
                        logger("人工恢复等待期间收到停止请求，保存已处理数据后退出。")
                        break
                    if recovery == "timeout":
                        reason = "请求过于频繁页面等待超时"
                        failed += 1
                        failed_rows.append(build_failed_row(index, title, homepage, old_url, reason))
                        logger(f"  - 失败: {reason}")
                        continue
                    logger("  - 页面已恢复，继续处理当前条。")
                elif gate_status == "login_or_verify":
                    if gate_keyword:
                        logger(f"  - 命中页面提示: {gate_keyword}")
                    recovery = wait_for_manual_recovery(
                        page,
                        timeout_sec=MANUAL_RECOVERY_TIMEOUT_SEC,
                        stop_event=stop_event,
                        logger=logger,
                        wait_reason="检测到登录/验证页面，请先扫码登录或完成验证。",
                    )
                    if recovery == "stopped":
                        stopped = True
                        logger("人工恢复等待期间收到停止请求，保存已处理数据后退出。")
                        break
                    if recovery == "timeout":
                        reason = "登录/验证页面等待超时"
                        failed += 1
                        failed_rows.append(build_failed_row(index, title, homepage, old_url, reason))
                        logger(f"  - 失败: {reason}")
                        continue
                    logger("  - 页面已恢复，继续处理当前条。")
                if not wait_for_notes_dom_ready(page, timeout_ms=15000):
                    reason = "笔记卡片 DOM 未就绪"
                    failed += 1
                    failed_rows.append(build_failed_row(index, title, homepage, old_url, reason))
                    logger("  - " + reason)
                    continue
                note_url, reason = extract_note_url_from_homepage(page, title, max_scrolls=max_scrolls)
                if note_url:
                    row[NOTE_URL_HEADER] = note_url
                    updated += 1
                    logger(f"  - 已更新: {note_url}")
                else:
                    failed += 1
                    failed_rows.append(build_failed_row(index, title, homepage, old_url, reason))
                    logger(f"  - 失败: {reason}")
            except Exception as exc:
                failed += 1
                reason = f"处理异常: {exc}"
                failed_rows.append(build_failed_row(index, title, homepage, old_url, reason))
                logger(f"  - 失败: {reason}")
            finally:
                logger(f"  - 本条耗时: {format_duration(time.monotonic() - row_started_at)}")

        browser.close()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_csv(output_path, fieldnames, rows)
    logger(f"已保存更新后 CSV: {output_path}")
    logger(f"OUTPUT_CSV: {output_path}")

    actual_failed_path = None
    if failed_rows:
        failed_path.parent.mkdir(parents=True, exist_ok=True)
        with failed_path.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["行号", "笔记标题", "主页链接", "原笔记地址", "失败原因"])
            writer.writeheader()
            writer.writerows(failed_rows)
        actual_failed_path = failed_path
        logger(f"失败数据: {failed_path}")
        logger(f"FAILED_CSV: {failed_path}")

    if stopped:
        logger(f"已停止: 更新={updated}, 跳过={skipped}, 失败={failed}, 总耗时={format_duration(time.monotonic() - total_started_at)}")
    else:
        logger(f"完成: 更新={updated}, 跳过={skipped}, 失败={failed}, 总耗时={format_duration(time.monotonic() - total_started_at)}")
    return output_path, actual_failed_path


def build_failed_row(row_index: int, title: str, homepage: str, old_url: str, reason: str) -> dict[str, str]:
    return {
        "行号": str(row_index),
        "笔记标题": title,
        "主页链接": homepage,
        "原笔记地址": old_url,
        "失败原因": reason,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="只访问主页DOM，提取带xsec_token的xhs笔记官方地址。")
    parser.add_argument("--source", help="源 CSV 文件")
    parser.add_argument("--output", default="", help="输出 CSV 文件；默认生成新文件")
    parser.add_argument("--open-browser", action="store_true", help="只打开/连接本机 Chrome/Edge 浏览器")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="Chrome 调试端口")
    parser.add_argument("--profile-dir", default=str(DEFAULT_PROFILE_DIR), help="Chrome/Edge 专用用户目录")
    parser.add_argument("--interval-sec", type=int, default=20, help="每次打开主页前的基础间隔秒数；实际会随机增加 0-20 秒")
    parser.add_argument("--max-scrolls", type=int, default=0, help="兼容保留参数；主页内不再滚动重试")
    parser.add_argument("--no-skip-existing-xsec", action="store_true", help="已有xsec_token时也重新覆盖")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        profile_dir = Path(args.profile_dir)
        if args.open_browser:
            launch_browser(args.port, profile_dir)
            log("请在打开的浏览器中登录xhs；登录后可回到GUI执行更新。")
            return 0
        if not args.source:
            raise ValueError("缺少 --source")
        update_note_urls(
            source=Path(args.source),
            output=Path(args.output) if args.output else None,
            port=args.port,
            profile_dir=profile_dir,
            skip_existing_xsec=not args.no_skip_existing_xsec,
            interval_sec=args.interval_sec,
            max_scrolls=args.max_scrolls,
        )
        return 0
    except Exception as exc:
        log(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
