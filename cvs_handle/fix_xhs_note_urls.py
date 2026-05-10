import argparse
import csv
import os
import random
import re
import shutil
import sys
import time
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlsplit, urlunsplit

from openpyxl import load_workbook, Workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BASE_DIR)
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "小红书笔记列表_规范Excel版.xlsx")
DEFAULT_SHEET_NAME = "小红书笔记列表"
DEFAULT_USER_DATA_DIR = os.path.join(REPO_DIR, "browser_session")

TITLE_HEADER = "笔记标题"
HOMEPAGE_HEADER = "主页链接"
NOTE_URL_HEADER = "笔记官方地址"


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def is_blank(value):
    return value is None or str(value).strip() == ""


def normalize_text(value):
    text = str(value or "").strip().lower()
    text = re.sub(r"[\s\u200b\u200c\u200d\ufeff]+", "", text)
    return text


def parse_note_date(date_str, current_year=None):
    if not date_str:
        return None
    date_str = str(date_str).strip()
    today = datetime.now()
    year = current_year or today.year
    
    if "昨天" in date_str:
        return today.replace(hour=0, minute=0, second=0, microsecond=0)
    if "前天" in date_str:
        from datetime import timedelta
        return (today - timedelta(days=2)).replace(hour=0, minute=0, second=0, microsecond=0)
    
    hours_match = re.search(r"(\d+)\s*小时前", date_str)
    if hours_match:
        hours = int(hours_match.group(1))
        return today.replace(hour=0, minute=0, second=0, microsecond=0)
    
    days_match = re.search(r"(\d+)\s*天前", date_str)
    if days_match:
        days = int(days_match.group(1))
        from datetime import timedelta
        return (today - timedelta(days=days)).replace(hour=0, minute=0, second=0, microsecond=0)
    
    match = re.search(r"(\d{1,2})[-月/](\d{1,2})", date_str)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        try:
            return datetime(year=year, month=month, day=day)
        except ValueError:
            return None
    
    return None


def parse_rows_spec(spec, min_row, max_row):
    if not spec:
        return list(range(min_row, max_row + 1))

    rows = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if end < start:
                start, end = end, start
            rows.extend(range(start, end + 1))
        else:
            rows.append(int(part))

    seen = set()
    result = []
    for row in rows:
        if row < min_row or row > max_row or row in seen:
            continue
        seen.add(row)
        result.append(row)
    return result


def header_map(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if value:
            headers[str(value).strip()] = col
    return headers


def require_headers(headers, required):
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError("Excel 缺少列: " + "、".join(missing))


def extract_note_id(url):
    url = str(url or "")
    match = re.search(r"/explore/([^/?#]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"/user/profile/[^/?#]+/([^/?#]+)", url)
    return match.group(1) if match else ""


def is_valid_note_link(url):
    note_id = extract_note_id(url)
    return bool(note_id and len(note_id) >= 8)


def is_clickable_profile_card_link(url):
    url = normalize_note_url(url)
    return (
        "/user/profile/" in url
        and "xsec_source=pc_user" in url
        and "xsec_token=" in url
        and "xsec_token=&" not in url
    )


def is_clickable_candidate_link(url):
    return is_valid_note_link(url) or is_clickable_profile_card_link(url)


def is_valid_address_bar_note_url(url, require_xsec_token=True):
    url = normalize_note_url(url)
    if not is_valid_note_link(url):
        return False
    if require_xsec_token and "xsec_token=" not in url:
        return False
    return True


def normalize_homepage_url(url):
    url = str(url or "").strip()
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/"):
        url = urljoin("https://www.xiaohongshu.com", url)
    elif not re.match(r"^https?://", url, flags=re.I):
        url = "https://" + url

    parts = urlsplit(url)
    path = parts.path.rstrip("/")
    return urlunsplit((parts.scheme or "https", parts.netloc, path, "", ""))


def normalize_note_url(url):
    url = str(url or "").strip()
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/"):
        url = urljoin("https://www.xiaohongshu.com", url)
    return url


def make_backup(path):
    if not os.path.exists(path):
        return ""
    folder, filename = os.path.split(path)
    name, ext = os.path.splitext(filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(folder, f"{name}_备份_{timestamp}{ext}")
    shutil.copy2(path, backup_path)
    return backup_path


def fallback_output_path(output_path):
    folder, filename = os.path.split(output_path)
    name, ext = os.path.splitext(filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(folder, f"{name}_自动保存_{timestamp}{ext or '.xlsx'}")


def save_workbook(wb, output_path):
    try:
        wb.save(output_path)
        return output_path
    except PermissionError:
        fallback_path = fallback_output_path(output_path)
        wb.save(fallback_path)
        print(f"  - 目标文件被占用，已改存到: {fallback_path}", flush=True)
        return fallback_path


def goto_with_retry(page, url, timeout_ms=45000, retries=2):
    last_error = None
    for attempt in range(retries + 1):
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            status = resp.status if resp else "无响应对象"
            final_url = page.url
            print(f"  - 页面加载完成: status={status}, 最终URL={final_url[:100]}", flush=True)
            return True
        except Exception as exc:
            last_error = exc
            if attempt >= retries:
                print(f"  - 页面打开失败(已重试{retries}次): {exc}", flush=True)
                raise
            print(f"  - 页面打开失败，准备重试({attempt + 1}/{retries}): {exc}", flush=True)
            page.wait_for_timeout(1200)
    if last_error:
        raise last_error
    return False


def click_notes_tab_if_visible(page):
    selectors = [
        "[role='tab']",
        ".reds-tabs-list .tab",
        ".user-tabs .tab",
        ".tabs-list .tab",
    ]
    for selector in selectors:
        try:
            tabs = page.locator(selector).filter(has_text="笔记")
            count = min(tabs.count(), 3)
            if count > 0:
                print(f"  - [笔记Tab] 选择器 '{selector}' 找到 {count} 个匹配", flush=True)
            for index in range(count):
                tab = tabs.nth(index)
                if tab.is_visible(timeout=500):
                    tab.click(timeout=1500)
                    page.wait_for_timeout(800)
                    print(f"  - [笔记Tab] 已点击选择器 '{selector}' 第{index}个tab", flush=True)
                    return True
        except Exception as exc:
            print(f"  - [笔记Tab] 选择器 '{selector}' 查找/点击异常: {exc}", flush=True)
    print(f"  - [笔记Tab] 未找到可见的笔记Tab (尝试了 {len(selectors)} 个选择器)", flush=True)
    return False


def find_note_candidate(page, title):
    target = normalize_text(title)
    if not target:
        return None

    result = page.evaluate(
        """
        ({ target }) => {
            const normalize = (value) => String(value || '')
                .trim()
                .toLowerCase()
                .replace(/[\\s\\u200b\\u200c\\u200d\\ufeff]+/g, '');

            const cleanHref = (href) => {
                if (!href) return '';
                try {
                    return new URL(href, location.href).href;
                } catch (error) {
                    return href;
                }
            };

            const extractNoteId = (href) => {
                const value = cleanHref(href);
                const match = value.match(/\\/explore\\/([^/?#]+)/)
                    || value.match(/\\/user\\/profile\\/[^/?#]+\\/([^/?#]+)/);
                return match ? match[1] : '';
            };

            const isValidNoteHref = (href) => {
                const noteId = extractNoteId(href);
                return Boolean(noteId && noteId.length >= 8);
            };

            const isClickableProfileCardHref = (href) => {
                const value = cleanHref(href);
                return value.includes('/user/profile/')
                    && value.includes('xsec_source=pc_user')
                    && value.includes('xsec_token=')
                    && !value.includes('xsec_token=&');
            };

            const isClickableCandidateHref = (href) => isValidNoteHref(href) || isClickableProfileCardHref(href);

            const scoreText = (text) => {
                const value = normalize(text);
                if (!value) return 0;
                if (value === target) return 100;
                if (value.includes(target)) return 90;
                if (target.includes(value) && value.length >= Math.min(10, target.length)) return 70;
                // 模糊匹配：计算最长公共子序列长度
                const lcs = (a, b) => {
                    const m = a.length, n = b.length;
                    const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
                    for (let i = 1; i <= m; i++) {
                        for (let j = 1; j <= n; j++) {
                            if (a[i - 1] === b[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
                            else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
                        }
                    }
                    return dp[m][n];
                };
                const commonLen = lcs(value, target);
                const similarity = commonLen / Math.max(value.length, target.length);
                if (similarity >= 0.6 && value.length >= 6) return Math.round(50 + similarity * 30);
                return 0;
            };

            const findNoteAnchor = (node) => {
                let current = node;
                for (let depth = 0; current && depth < 8; depth += 1) {
                    if (current.matches && current.matches('a[href*="/explore/"], a[href*="/user/profile/"]')) {
                        const href = current.getAttribute('href') || current.href;
                        if (isClickableCandidateHref(href)) return current;
                    }
                    if (current.querySelector) {
                        const anchors = Array.from(current.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]'));
                        const anchor = anchors.find((item) => isValidNoteHref(item.getAttribute('href') || item.href))
                            || anchors.find((item) => isClickableProfileCardHref(item.getAttribute('href') || item.href));
                        if (anchor) return anchor;
                    }
                    current = current.parentElement;
                }
                return null;
            };

            const findClickableNode = (anchor) => {
                const selectors = [
                    '.note-item',
                    '.note-card',
                    '.note-card-wrapper',
                    '.feeds-page .note-item',
                    '[class*="note-item"]',
                    '[class*="note-card"]',
                    '[class*="cover"]',
                    '[class*="card"]',
                ];
                for (const selector of selectors) {
                    const node = anchor.closest(selector);
                    if (!node) continue;
                    const rect = node.getBoundingClientRect();
                    if (rect.width >= 80 && rect.height >= 80) {
                        return node;
                    }
                }
                return anchor;
            };

            const buildCandidate = (anchor, text, score) => {
                const href = cleanHref(anchor.getAttribute('href') || anchor.href);
                const clickable = findClickableNode(anchor);
                const rect = clickable.getBoundingClientRect();
                const anchorRect = anchor.getBoundingClientRect();
                const targetRect = rect.width >= 20 && rect.height >= 20 ? rect : anchorRect;
                
                const extractDate = (node) => {
                    if (!node) return '';
                    const selectors = ['.date', '[class*="date"]', 'time', '.time', '[class*="time"]', 'span', 'div'];
                    for (const sel of selectors) {
                        const el = node.querySelector(sel);
                        if (!el) continue;
                        const txt = (el.innerText || el.textContent || '').trim();
                        if (txt && /\\d{1,2}[-月/]\\d{1,2}/.test(txt)) return txt;
                        if (txt && /昨天|前天|\\d+小时前|\\d+天前|\\d+分钟前/.test(txt)) return txt;
                    }
                    const allText = (node.innerText || node.textContent || '');
                    const dateMatch = allText.match(/(\\d{1,2}[-月/]\\d{1,2}日?)/);
                    if (dateMatch) return dateMatch[1];
                    return '';
                };
                
                const dateStr = extractDate(clickable) || extractDate(anchor);
                
                return {
                    href,
                    noteId: extractNoteId(href),
                    text,
                    score,
                    x: targetRect.left + targetRect.width / 2,
                    y: targetRect.top + Math.min(targetRect.height * 0.42, targetRect.height / 2),
                    width: targetRect.width,
                    height: targetRect.height,
                    dateStr,
                };
            };

            const candidates = [];
            const noteAnchors = Array.from(document.querySelectorAll(
                'a[href*="/explore/"], a[href*="/user/profile/"]'
            )).filter((anchor) => {
                const href = anchor.getAttribute('href') || anchor.href;
                if (!isClickableCandidateHref(href)) return false;
                const rect = anchor.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            });

            for (const anchor of noteAnchors) {
                const rawHref = anchor.getAttribute('href') || anchor.href;
                if (!isClickableCandidateHref(rawHref)) continue;
                let current = anchor;
                for (let depth = 0; current && depth < 6; depth += 1) {
                    const text = current.innerText || current.textContent || '';
                    const score = scoreText(text);
                    if (score) {
                        candidates.push(buildCandidate(anchor, text, score));
                        break;
                    }
                    current = current.parentElement;
                }
            }

            const textNodes = Array.from(document.querySelectorAll(
                '.title, [class*="title"], span, p, div'
            ));
            for (const node of textNodes) {
                const text = node.innerText || node.textContent || '';
                const score = scoreText(text);
                if (!score) continue;
                const anchor = findNoteAnchor(node);
                if (!anchor) continue;
                candidates.push(buildCandidate(anchor, text, score + 5));
            }

            candidates.sort((left, right) => right.score - left.score);
            const best = candidates.find((item) => item.href && item.width > 0 && item.height > 0);
            return {
                best: best || null,
                totalAnchors: noteAnchors.length,
                totalCandidates: candidates.length,
                topScores: candidates.slice(0, 5).map((c) => ({ score: c.score, noteId: c.noteId, text: (c.text || '').substring(0, 30) })),
            };
        }
        """,
        {"target": target},
    )
    if result:
        best = result.get("best")
        total_anchors = result.get("totalAnchors", 0)
        total_candidates = result.get("totalCandidates", 0)
        top_scores = result.get("topScores", [])
        if best:
            print(f"  - [查找卡片] 找到匹配: score={best.get('score', 0)} noteId={best.get('noteId', '')} href={normalize_note_url(best.get('href', ''))[:80]} 坐标=({best.get('x', 0):.0f},{best.get('y', 0):.0f}) 尺寸={best.get('width', 0):.0f}x{best.get('height', 0):.0f}", flush=True)
        else:
            print(f"  - [查找卡片] 未找到匹配: 页面锚点数={total_anchors}, 候选数={total_candidates}, 目标标题={target[:40]}", flush=True)
        if top_scores:
            scores_desc = ", ".join(f"[score={s['score']} noteId={s['noteId']} text={s['text'][:20]}]" for s in top_scores)
            print(f"  - [查找卡片] 候选排名: {scores_desc}", flush=True)
        return best
    print(f"  - [查找卡片] JS执行返回空, 目标标题={target[:40]}", flush=True)
    return None


def click_note_candidate(page, candidate):
    if not candidate or not is_clickable_candidate_link(candidate.get("href")):
        print(f"  - [点击卡片] 候选无效或链接不可点击: href={normalize_note_url(candidate.get('href', ''))[:80] if candidate else 'None'}", flush=True)
        return False

    viewport = page.viewport_size or {}
    viewport_width = viewport.get("width", 0)
    viewport_height = viewport.get("height", 0)
    if (
        candidate.get("x", 0) > 0
        and candidate.get("y", 0) > 0
        and (not viewport_width or candidate["x"] < viewport_width)
        and (not viewport_height or candidate["y"] < viewport_height)
    ):
        print(f"  - [点击卡片] 方式1-坐标点击: ({candidate['x']:.0f}, {candidate['y']:.0f}), viewport=({viewport_width}x{viewport_height})", flush=True)
        page.mouse.click(float(candidate["x"]), float(candidate["y"]))
        return True

    print(f"  - [点击卡片] 方式1失败(坐标越界或为0: x={candidate.get('x', 0):.0f}, y={candidate.get('y', 0):.0f}, viewport={viewport_width}x{viewport_height}), 尝试方式2-JS查找元素点击", flush=True)

    point = page.evaluate(
        """
        ({ href, text }) => {
            const cleanHref = (value) => {
                try {
                    return new URL(value, location.href).href;
                } catch (error) {
                    return value || '';
                }
            };
            const target = cleanHref(href);
            const normalize = (value) => String(value || '')
                .trim()
                .toLowerCase()
                .replace(/[\\s\\u200b\\u200c\\u200d\\ufeff]+/g, '');
            const targetText = normalize(text);
            const anchors = Array.from(document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]'));
            const sameHrefAnchors = anchors.filter((item) => cleanHref(item.getAttribute('href') || item.href) === target);
            const anchor = sameHrefAnchors.find((item) => targetText && normalize(item.innerText || item.textContent).includes(targetText))
                || anchors.find((item) => cleanHref(item.getAttribute('href') || item.href) === target)
                || anchors.find((item) => cleanHref(item.getAttribute('href') || item.href).split('?')[0] === target.split('?')[0]);
            if (!anchor) return null;

            const selectors = [
                '.note-item',
                '.note-card',
                '.note-card-wrapper',
                '.feeds-page .note-item',
                '[class*="note-item"]',
                '[class*="note-card"]',
                '[class*="cover"]',
                '[class*="card"]',
            ];
            let clickable = anchor;
            for (const selector of selectors) {
                const node = anchor.closest(selector);
                if (!node) continue;
                const rect = node.getBoundingClientRect();
                if (rect.width >= 80 && rect.height >= 80) {
                    clickable = node;
                    break;
                }
            }

            clickable.scrollIntoView({ block: 'center', inline: 'center' });
            const rect = clickable.getBoundingClientRect();
            const x = rect.left + rect.width / 2;
            const y = rect.top + Math.min(rect.height * 0.42, rect.height / 2);
            return { x, y, width: rect.width, height: rect.height };
        }
        """,
        {"href": candidate.get("href"), "text": candidate.get("text", "")},
    )

    if point and point.get("x", 0) > 0 and point.get("y", 0) > 0:
        print(f"  - [点击卡片] 方式2-JS定位点击: ({point['x']:.0f}, {point['y']:.0f}), 尺寸={point.get('width', 0):.0f}x{point.get('height', 0):.0f}", flush=True)
        page.wait_for_timeout(300)
        page.mouse.click(float(point["x"]), float(point["y"]))
        return True

    print(f"  - [点击卡片] 方式2失败(JS未找到有效坐标), 尝试方式3-直接anchor.click()", flush=True)
    js_result = page.evaluate(
        """
        ({ href }) => {
            const cleanHref = (value) => {
                try {
                    return new URL(value, location.href).href;
                } catch (error) {
                    return value || '';
                }
            };
            const target = cleanHref(href);
            const anchors = Array.from(document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]'));
            const anchor = anchors.find((item) => cleanHref(item.getAttribute('href') || item.href) === target)
                || anchors.find((item) => cleanHref(item.getAttribute('href') || item.href).split('?')[0] === target.split('?')[0]);
            if (!anchor) return false;
            anchor.scrollIntoView({ block: 'center', inline: 'center' });
            anchor.click();
            return true;
        }
        """,
        {"href": candidate.get("href")},
    )
    print(f"  - [点击卡片] 方式3-anchor.click()结果: {js_result}", flush=True)
    return js_result


def click_note_by_href(page, href):
    if not href or not is_valid_note_link(href):
        return False
    return page.evaluate(
        """
        ({ href }) => {
            const normalizeHref = (value) => {
                try {
                    return new URL(value, location.href).href;
                } catch (error) {
                    return value || '';
                }
            };
            const target = normalizeHref(href);
            const anchors = Array.from(document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]'));
            const anchor = anchors.find((item) => normalizeHref(item.getAttribute('href') || item.href) === target)
                || anchors.find((item) => normalizeHref(item.getAttribute('href') || item.href).split('?')[0] === target.split('?')[0]);
            if (!anchor) return false;
            anchor.scrollIntoView({ block: 'center', inline: 'center' });
            anchor.click();
            return true;
        }
        """,
        {"href": href},
    )


def wait_for_current_note_url(page, expected_note_id="", timeout_ms=6000, require_xsec_token=True):
    deadline = time.time() + timeout_ms / 1000
    best = ""
    expected_note_id = str(expected_note_id or "").strip()
    check_count = 0
    first_url = normalize_note_url(page.url)
    print(f"  - [等待地址栏] 开始等待: 当前URL={first_url[:80]}, 期望noteId={expected_note_id}, 超时={timeout_ms}ms, 需要xsec_token={require_xsec_token}", flush=True)
    while time.time() < deadline:
        current_url = normalize_note_url(page.url)
        current_note_id = extract_note_id(current_url)
        check_count += 1
        if is_valid_address_bar_note_url(current_url, require_xsec_token=False):
            if not expected_note_id or current_note_id == expected_note_id:
                best = current_url
                if is_valid_address_bar_note_url(current_url, require_xsec_token=require_xsec_token):
                    print(f"  - [等待地址栏] 成功: 检查{check_count}次, URL={current_url[:80]}", flush=True)
                    return current_url
        page.wait_for_timeout(250)
    elapsed = time.time() - (deadline - timeout_ms / 1000)
    if best:
        print(f"  - [等待地址栏] 超时{elapsed:.1f}s/{timeout_ms/1000:.1f}s 检查{check_count}次, 有基础URL但无xsec_token: {best[:80]}", flush=True)
    else:
        final_url = normalize_note_url(page.url)
        final_note_id = extract_note_id(final_url)
        print(f"  - [等待地址栏] 超时{elapsed:.1f}s/{timeout_ms/1000:.1f}s 检查{check_count}次, 地址栏未跳转到笔记页, 最终URL={final_url[:80]}, noteId={final_note_id}", flush=True)
    if not require_xsec_token and best:
        return best
    return ""


def check_verification_popup(page, max_wait_sec=120):
    try:
        indicators = page.evaluate(
            """
            () => {
                const captchaSelectors = [
                    '[class*="captcha"]',
                    '[class*="verify-modal"]',
                    '[class*="slider-verify"]',
                    '[class*="puzzle-verify"]',
                    '.secsdk-captcha-wrapper',
                    '#secsdk-captcha-wrapper',
                    '[id*="captcha"]'
                ];
                let hasCaptcha = false;
                for (const sel of captchaSelectors) {
                    const el = document.querySelector(sel);
                    if (el && el.offsetParent !== null) {
                        hasCaptcha = true;
                        break;
                    }
                }
                
                const loginModal = document.querySelector('.login-modal, [class*="login-modal"], .login-container');
                const hasLoginModal = loginModal && loginModal.offsetParent !== null;
                
                const body = document.body ? (document.body.innerText || '') : '';
                const hasCaptchaText = (
                    body.includes('请完成验证')
                    || body.includes('请通过验证')
                    || body.includes('滑动验证')
                ) && hasCaptcha;
                
                return { hasCaptcha: hasCaptcha || hasCaptchaText, hasLoginModal };
            }
            """
        )
        if not indicators.get("hasCaptcha") and not indicators.get("hasLoginModal"):
            return True

        if indicators.get("hasCaptcha"):
            print(f"  - [风控检测] 检测到验证弹窗！请在浏览器中手动完成验证，最多等待 {max_wait_sec} 秒...", flush=True)
        if indicators.get("hasLoginModal"):
            print(f"  - [风控检测] 检测到登录弹窗！请在浏览器中完成登录，最多等待 {max_wait_sec} 秒...", flush=True)

        deadline = time.time() + max_wait_sec
        while time.time() < deadline:
            page.wait_for_timeout(2000)
            try:
                recheck = page.evaluate(
                    """
                    () => {
                        const captchaSelectors = [
                            '[class*="captcha"]',
                            '[class*="verify-modal"]',
                            '[class*="slider-verify"]',
                            '[class*="puzzle-verify"]',
                            '.secsdk-captcha-wrapper',
                            '#secsdk-captcha-wrapper',
                            '[id*="captcha"]'
                        ];
                        let hasCaptcha = false;
                        for (const sel of captchaSelectors) {
                            const el = document.querySelector(sel);
                            if (el && el.offsetParent !== null) {
                                hasCaptcha = true;
                                break;
                            }
                        }
                        
                        const loginModal = document.querySelector('.login-modal, [class*="login-modal"], .login-container');
                        const hasLoginModal = loginModal && loginModal.offsetParent !== null;
                        
                        return { hasCaptcha, hasLoginModal };
                    }
                    """
                )
                if not recheck.get("hasCaptcha") and not recheck.get("hasLoginModal"):
                    print("  - [风控检测] 验证/登录已完成，继续执行", flush=True)
                    page.wait_for_timeout(random.randint(1000, 2000))
                    return True
            except Exception:
                pass
            remaining = int(deadline - time.time())
            if remaining > 0 and remaining % 10 == 0:
                print(f"  - [风控检测] 仍在等待验证... 剩余 {remaining} 秒", flush=True)

        print("  - [风控检测] 验证等待超时，继续执行（可能仍被限制）", flush=True)
        return False
    except Exception as exc:
        print(f"  - [风控检测] 检测异常: {exc}", flush=True)
        return True


def random_delay(min_ms=500, max_ms=2000):
    delay = random.randint(min_ms, max_ms)
    time.sleep(delay / 1000.0)


def locate_note_url(
    page,
    homepage_url,
    title,
    max_scrolls,
    scroll_pixels,
    detail_wait_ms,
    require_xsec_token=True,
    per_item_timeout_sec=20,
    target_date=None,
):
    homepage_url = normalize_homepage_url(homepage_url)
    target_date_str = target_date.strftime("%Y-%m-%d") if target_date else "未设置"
    print(f"  - [定位流程] 开始: 主页={homepage_url} 查找标题={title[:40]} max_scrolls={max_scrolls} detail_wait_ms={detail_wait_ms} require_xsec_token={require_xsec_token} per_item_timeout={per_item_timeout_sec}s target_date={target_date_str}", flush=True)
    start_time = time.time()
    goto_with_retry(page, homepage_url)
    page.wait_for_timeout(random.randint(1500, 2500))
    
    check_verification_popup(page)
    
    tab_clicked = click_notes_tab_if_visible(page)
    if not tab_clicked:
        print(f"  - [定位流程] 未点击到笔记Tab，页面可能不是用户主页或Tab结构已变", flush=True)

    last_height = 0
    no_change_count = 0
    max_no_change = 3
    scroll_count = 0
    early_stop_due_to_date = False

    for scroll_index in range(max_scrolls + 1):
        elapsed = time.time() - start_time
        if elapsed > per_item_timeout_sec:
            print(f"  - [定位流程] 已超时 ({elapsed:.1f}秒 > {per_item_timeout_sec}秒)，停止查找 (已滚动{scroll_count}次)", flush=True)
            return ""

        candidate = find_note_candidate(page, title)
        if candidate and candidate.get("href"):
            print(f"  - [定位流程] 滚动第{scroll_index}次: 找到候选 score={candidate.get('score', 0)} href={normalize_note_url(candidate['href'])[:80]}", flush=True)
            href = normalize_note_url(candidate["href"])
            note_id = extract_note_id(href)
            if is_clickable_candidate_link(href):
                print(f"  - [定位流程] 找到匹配卡片，准备点击: {href} noteId={note_id}", flush=True)
                clicked = False
                try:
                    clicked = click_note_candidate(page, candidate)
                except Exception as exc:
                    print(f"  - [定位流程] 坐标点击失败，尝试链接点击: {exc}", flush=True)
                    try:
                        clicked = click_note_by_href(page, href)
                    except Exception as exc2:
                        print(f"  - [定位流程] 链接点击也失败: {exc2}", flush=True)
                        clicked = False

                if clicked:
                    resolved_url = wait_for_current_note_url(
                        page,
                        expected_note_id=note_id,
                        timeout_ms=detail_wait_ms,
                        require_xsec_token=require_xsec_token,
                    )
                    if resolved_url:
                        print(f"  - [定位流程] 成功获取地址: {resolved_url}", flush=True)
                        return resolved_url
                    if is_valid_address_bar_note_url(href, require_xsec_token=require_xsec_token):
                        print(f"  - [定位流程] 地址栏未变化，使用卡片链接: {href}", flush=True)
                        return href
                    current_page_url = normalize_note_url(page.url)
                    has_note_id = bool(extract_note_id(current_page_url))
                    print(f"  - [定位流程] 点击后未拿到xsec_token地址: 当前URL={current_page_url[:100]} 是否含笔记ID={has_note_id} 卡片href={href[:80]}", flush=True)
                    return ""
                print(f"  - [定位流程] 找到卡片但点击失败: href={normalize_note_url(candidate.get('href', ''))[:80]}", flush=True)
                return ""
            print(f"  - [定位流程] 忽略无效候选地址(非可点击链接): {href[:80]}", flush=True)
        elif candidate:
            print(f"  - [定位流程] 滚动第{scroll_index}次: 找到候选但无有效href score={candidate.get('score', 0)} text={str(candidate.get('text', ''))[:30]}", flush=True)
        else:
            if scroll_index % 3 == 0:
                print(f"  - [定位流程] 滚动第{scroll_index}次: 未找到任何候选", flush=True)

        if target_date and scroll_index > 0 and scroll_index % 2 == 0:
            page_dates = page.evaluate(
                """
                () => {
                    const dates = [];
                    const selectors = ['.date', '[class*="date"]', 'time', '.time', '[class*="time"]'];
                    for (const sel of selectors) {
                        const els = document.querySelectorAll(sel);
                        for (const el of els) {
                            const txt = (el.innerText || el.textContent || '').trim();
                            if (txt && /\\d{1,2}[-月/]\\d{1,2}/.test(txt)) {
                                dates.push(txt);
                            }
                        }
                    }
                    return dates.slice(0, 10);
                }
                """
            )
            if page_dates:
                parsed_dates = []
                for d in page_dates:
                    pd = parse_note_date(d)
                    if pd:
                        parsed_dates.append(pd)
                if parsed_dates:
                    min_date = min(parsed_dates)
                    if min_date < target_date:
                        print(f"  - [定位流程] 检测到笔记日期早于目标日期: 最早日期={min_date.strftime('%Y-%m-%d')} < 目标日期={target_date.strftime('%Y-%m-%d')}, 提前停止滚动", flush=True)
                        early_stop_due_to_date = True
                        break

        current_height = page.evaluate(
            "() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        if scroll_index >= max_scrolls:
            break
        page.mouse.wheel(0, scroll_pixels)
        page.wait_for_timeout(random.randint(600, 1200))
        scroll_count += 1

        next_height = page.evaluate(
            "() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        if next_height == last_height == current_height:
            no_change_count += 1
            if no_change_count >= max_no_change:
                print(f"  - [定位流程] 页面已到底部，连续 {no_change_count} 次无变化，提前退出 (已滚动{scroll_count}次, 页面高度={current_height})", flush=True)
                break
            page.wait_for_timeout(400)
        else:
            no_change_count = 0
        last_height = next_height

    total_elapsed = time.time() - start_time
    if early_stop_due_to_date:
        print(f"  - [定位流程] 失败(日期早于目标): 滚动{scroll_count}次后未找到匹配卡片, 耗时{total_elapsed:.1f}s, title={title[:40]}", flush=True)
    else:
        print(f"  - [定位流程] 失败: 滚动{scroll_count}次后未找到匹配卡片, 耗时{total_elapsed:.1f}s, title={title[:40]}", flush=True)
    return ""


def build_rows(ws, headers, rows):
    items = []
    title_col = headers[TITLE_HEADER]
    homepage_col = headers[HOMEPAGE_HEADER]
    url_col = headers[NOTE_URL_HEADER]
    for row_idx in rows:
        title = ws.cell(row=row_idx, column=title_col).value
        homepage_url = ws.cell(row=row_idx, column=homepage_col).value
        old_url = ws.cell(row=row_idx, column=url_col).value
        items.append(
            {
                "row": row_idx,
                "title": str(title or "").strip(),
                "homepage_url": str(homepage_url or "").strip(),
                "old_url": str(old_url or "").strip(),
            }
        )
    return items


def export_failed_items(failed_items, output_path):
    if not failed_items:
        return ""
    folder = os.path.dirname(output_path)
    name = os.path.splitext(os.path.basename(output_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    failed_path = os.path.join(folder, f"{name}_失败数据_{timestamp}.csv")
    try:
        with open(failed_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["行号", "笔记标题", "主页链接", "原笔记地址", "失败原因"])
            writer.writeheader()
            for item in failed_items:
                writer.writerow({
                    "行号": item["row"],
                    "笔记标题": item["title"],
                    "主页链接": item["homepage_url"],
                    "原笔记地址": item["old_url"],
                    "失败原因": item["reason"],
                })
        return failed_path
    except Exception as exc:
        print(f"导出失败数据出错: {exc}", flush=True)
        return ""


def update_note_urls(args):
    excel_path = os.path.abspath(args.excel_path)
    output_path = os.path.abspath(args.output_path or args.excel_path)
    sheet_name = args.sheet_name

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到文件: {excel_path}")

    is_csv = excel_path.lower().endswith(".csv")

    if is_csv:
        return update_note_urls_csv(args, excel_path, output_path)
    else:
        return update_note_urls_excel(args, excel_path, output_path, sheet_name)


def update_note_urls_csv(args, csv_path, output_path):
    skip_no_title = args.skip_no_title and not args.no_skip_no_title
    
    target_date = None
    if args.target_date:
        try:
            if "-" in args.target_date:
                parts = args.target_date.split("-")
                if len(parts) == 3:
                    target_date = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                elif len(parts) == 2:
                    target_date = datetime(datetime.now().year, int(parts[0]), int(parts[1]))
        except Exception as e:
            print(f"警告: 无法解析目标日期 '{args.target_date}': {e}", flush=True)
    
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows_data = list(reader)
        headers = {h: idx for idx, h in enumerate(reader.fieldnames or [])}

    require_headers(headers, [TITLE_HEADER, HOMEPAGE_HEADER, NOTE_URL_HEADER])

    total_rows = len(rows_data)
    all_row_indices = list(range(2, total_rows + 2))  # CSV 数据行号从2开始（1是表头）
    if args.limit:
        all_row_indices = all_row_indices[: args.limit]
    if not all_row_indices:
        print("没有需要处理的行。", flush=True)
        return output_path

    row_indices = all_row_indices
    if args.parallel_index >= 0 and args.parallel_total > 1:
        chunk_size = (len(all_row_indices) + args.parallel_total - 1) // args.parallel_total
        start_idx = args.parallel_index * chunk_size
        end_idx = min(start_idx + chunk_size, len(all_row_indices))
        row_indices = all_row_indices[start_idx:end_idx]
        actual_start_row = row_indices[0] if row_indices else 0
        actual_end_row = row_indices[-1] if row_indices else 0
        print(f"并行模式: 任务 {args.parallel_index + 1}/{args.parallel_total}，处理行 {actual_start_row}-{actual_end_row}", flush=True)

    items = []
    for row_idx in row_indices:
        row_data = rows_data[row_idx - 2]
        items.append({
            "row": row_idx,
            "title": str(row_data.get(TITLE_HEADER, "") or "").strip(),
            "homepage_url": str(row_data.get(HOMEPAGE_HEADER, "") or "").strip(),
            "old_url": str(row_data.get(NOTE_URL_HEADER, "") or "").strip(),
        })

    if os.path.abspath(csv_path) == os.path.abspath(output_path):
        backup_path = make_backup(csv_path)
        if backup_path:
            print(f"已备份原文件: {backup_path}", flush=True)

    user_data_dir = args.user_data_dir
    if args.parallel_session:
        if f"parallel_{args.parallel_session}_{args.parallel_index}" not in args.user_data_dir:
            user_data_dir = os.path.join(args.user_data_dir, f"parallel_{args.parallel_session}_{args.parallel_index}")

    session_dirs = []
    if args.sessions:
        for s in args.sessions.split(","):
            s = s.strip()
            if s:
                if os.path.isabs(s):
                    session_dirs.append(s)
                else:
                    session_dirs.append(os.path.join(os.path.dirname(BASE_DIR), s))
    
    if not session_dirs:
        session_dirs = [user_data_dir]
    
    print(f"CSV: {csv_path}", flush=True)
    print(f"输出: {output_path}", flush=True)
    print(f"处理行数: {len(items)}", flush=True)
    print(f"使用 {len(session_dirs)} 个账号轮询:", flush=True)
    for i, sd in enumerate(session_dirs):
        print(f"  账号{i+1}: {sd}", flush=True)

    updated = 0
    skipped = 0
    failed = 0
    failed_items = []

    session_mode = getattr(args, 'session_mode', 'rotate')

    with sync_playwright() as p:
        contexts = []
        pages = []
        current_session_idx = 0
        
        if session_mode == "bind":
            print(f"\n[绑定模式] 同时打开 {len(session_dirs)} 个浏览器窗口...", flush=True)
        else:
            print(f"\n[轮询模式] 同时打开 {len(session_dirs)} 个浏览器窗口...", flush=True)
        
        for i, session_dir in enumerate(session_dirs):
            print(f"  - 启动账号{i+1}的浏览器...", flush=True)
            context = p.chromium.launch_persistent_context(
                user_data_dir=session_dir,
                headless=args.headless,
                slow_mo=args.slow_mo,
                viewport={"width": 1440, "height": 1000},
            )
            page = context.pages[0] if context.pages else context.new_page()
            
            if args.login_wait:
                first_homepage_url = ""
                for item in items:
                    first_homepage_url = normalize_homepage_url(item.get("homepage_url"))
                    if first_homepage_url:
                        break
                login_url = first_homepage_url or "https://www.xiaohongshu.com/explore"
                page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2000)
                check_verification_popup(page)
            
            contexts.append(context)
            pages.append(page)
        
        print(f"所有 {len(session_dirs)} 个浏览器窗口已就绪\n", flush=True)
        
        def get_current_page():
            return pages[current_session_idx % len(pages)]

        try:
            for index, item in enumerate(items, start=1):
                row_idx = item["row"]
                title = item["title"]
                homepage_url = item["homepage_url"]
                old_url = item["old_url"]
                old_note_id = extract_note_id(old_url)

                if len(session_dirs) > 1:
                    expected_session_idx = (index - 1) % len(session_dirs)
                    if expected_session_idx != current_session_idx:
                        current_session_idx = expected_session_idx
                        print(f"  - [切换账号] → 账号{current_session_idx + 1}", flush=True)

                page = get_current_page()
                print(f"\n[{index}/{len(items)}] 第 {row_idx} 行: {title}", flush=True)
                print(f"  - 使用账号{current_session_idx + 1}", flush=True)
                
                if index > 1:
                    random_delay(800, 2000)
                
                batch_size = getattr(args, 'batch_size', 0)
                batch_interval = getattr(args, 'batch_interval', 30)
                if batch_size > 0 and index > 1 and (index - 1) % batch_size == 0:
                    print(f"  - [批次休息] 已处理 {index - 1} 条，休息 {batch_interval} 秒避免风控...", flush=True)
                    time.sleep(batch_interval)
                
                check_verification_popup(page)
                
                if skip_no_title and ("笔记暂未设置标题" in title or not title):
                    skipped += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": "笔记无标题(不可恢复)",
                    })
                    print("  - 跳过：笔记无标题，无法匹配", flush=True)
                    continue
                
                if not title or not homepage_url:
                    skipped += 1
                    print("  - 跳过：标题或主页链接为空", flush=True)
                    continue
                if args.only_empty and not is_blank(old_url):
                    skipped += 1
                    print("  - 跳过：笔记官方地址已有内容", flush=True)
                    continue

                try:
                    note_url = locate_note_url(
                        page,
                        homepage_url,
                        title,
                        max_scrolls=args.max_scrolls,
                        scroll_pixels=args.scroll_pixels,
                        detail_wait_ms=args.detail_wait_ms,
                        require_xsec_token=not args.allow_basic_url,
                        per_item_timeout_sec=args.per_item_timeout,
                        target_date=target_date,
                    )
                except PlaywrightTimeoutError as exc:
                    note_url = ""
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"查找超时: {exc}",
                    })
                    print(f"  - 查找超时: {exc}", flush=True)
                    continue
                except Exception as exc:
                    note_url = ""
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"查找失败: {exc}",
                    })
                    print(f"  - 查找失败: {exc}", flush=True)
                    continue

                if not note_url:
                    failed += 1
                    old_url_has_note_id = bool(extract_note_id(old_url))
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}, 原地址={old_url[:80]})",
                    })
                    print(f"  - 未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}), 保持原值", flush=True)
                    continue

                new_note_id = extract_note_id(note_url)
                if not new_note_id:
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"地址没有笔记ID: {note_url}",
                    })
                    print(f"  - 找到的地址没有笔记ID，保持原值: {note_url}", flush=True)
                    continue

                rows_data[row_idx - 2][NOTE_URL_HEADER] = note_url
                updated += 1
                id_note = ""
                if old_note_id and new_note_id and old_note_id != new_note_id:
                    id_note = f"（注意：笔记ID从 {old_note_id} 变为 {new_note_id}）"
                print(f"  - 已更新: {note_url} {id_note}", flush=True)

                if args.save_every and updated % args.save_every == 0:
                    save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
                    print(f"  - 已保存进度: {output_path}", flush=True)
        except KeyboardInterrupt:
            print("\n\n用户中断，正在保存已处理数据...", flush=True)
            save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
            print(f"已保存 {updated} 条更新到: {output_path}", flush=True)
            raise
        finally:
            for ctx in contexts:
                try:
                    ctx.close()
                except Exception:
                    pass

    save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
    print("\n处理完成。", flush=True)
    print(f"更新: {updated}，跳过: {skipped}，失败: {failed}", flush=True)
    print(f"结果文件: {output_path}", flush=True)

    if failed_items:
        print("\n" + "=" * 60, flush=True)
        print("失败数据明细：", flush=True)
        print("=" * 60, flush=True)
        for item in failed_items:
            print(f"\n行号: {item['row']}", flush=True)
            print(f"  标题: {item['title']}", flush=True)
            print(f"  主页: {item['homepage_url']}", flush=True)
            print(f"  原地址: {item['old_url']}", flush=True)
            print(f"  失败原因: {item['reason']}", flush=True)

        failed_path = export_failed_items(failed_items, output_path)
        if failed_path:
            print(f"\n失败数据已导出到: {failed_path}", flush=True)

    return output_path


def save_csv(rows_data, output_path, fieldnames, original_csv_path=None):
    is_csv_output = output_path.lower().endswith(".csv")
    if is_csv_output:
        if original_csv_path and os.path.exists(original_csv_path):
            # 读取原始文件，只更新目标列，保持其他列原样
            with open(original_csv_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                original_rows = list(reader)
            
            # 找到笔记官方地址列的索引
            header_row = original_rows[0]
            try:
                url_col_idx = header_row.index(NOTE_URL_HEADER)
            except ValueError:
                url_col_idx = -1
            
            if url_col_idx >= 0:
                # 更新对应行的笔记官方地址
                for row_idx_in_data, row_data in enumerate(rows_data, start=1):
                    if row_idx_in_data < len(original_rows):
                        original_rows[row_idx_in_data][url_col_idx] = row_data.get(NOTE_URL_HEADER, "")
                
                with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerows(original_rows)
                return
        
        # 回退到原始方式
        with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows_data)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "数据"
        for col_idx, header in enumerate(fieldnames, 1):
            ws.cell(row=1, column=col_idx, value=header)
        for row_idx, row_data in enumerate(rows_data, 2):
            for col_idx, header in enumerate(fieldnames, 1):
                ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ""))
        wb.save(output_path)


def update_note_urls_excel(args, excel_path, output_path, sheet_name):
    skip_no_title = args.skip_no_title and not args.no_skip_no_title
    
    target_date = None
    if args.target_date:
        try:
            if "-" in args.target_date:
                parts = args.target_date.split("-")
                if len(parts) == 3:
                    target_date = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                elif len(parts) == 2:
                    target_date = datetime(datetime.now().year, int(parts[0]), int(parts[1]))
        except Exception as e:
            print(f"警告: 无法解析目标日期 '{args.target_date}': {e}", flush=True)
    
    wb = load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到 sheet: {sheet_name}，当前有: {', '.join(wb.sheetnames)}")
    ws = wb[sheet_name]
    headers = header_map(ws)
    require_headers(headers, [TITLE_HEADER, HOMEPAGE_HEADER, NOTE_URL_HEADER])

    rows = parse_rows_spec(args.rows, 2, ws.max_row)
    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        print("没有需要处理的行。", flush=True)
        return output_path

    if args.parallel_index >= 0 and args.parallel_total > 1:
        chunk_size = (len(rows) + args.parallel_total - 1) // args.parallel_total
        start_idx = args.parallel_index * chunk_size
        end_idx = min(start_idx + chunk_size, len(rows))
        rows = rows[start_idx:end_idx]
        print(f"并行模式: 任务 {args.parallel_index + 1}/{args.parallel_total}，处理行 {start_idx + 1}-{end_idx}", flush=True)

    items = build_rows(ws, headers, rows)
    url_col = headers[NOTE_URL_HEADER]

    if os.path.abspath(excel_path) == os.path.abspath(output_path):
        backup_path = make_backup(excel_path)
        if backup_path:
            print(f"已备份原文件: {backup_path}", flush=True)

    user_data_dir = args.user_data_dir
    if args.parallel_session:
        if f"parallel_{args.parallel_session}_{args.parallel_index}" not in args.user_data_dir:
            user_data_dir = os.path.join(args.user_data_dir, f"parallel_{args.parallel_session}_{args.parallel_index}")

    session_dirs = []
    if args.sessions:
        for s in args.sessions.split(","):
            s = s.strip()
            if s:
                if os.path.isabs(s):
                    session_dirs.append(s)
                else:
                    session_dirs.append(os.path.join(os.path.dirname(BASE_DIR), s))
    
    if not session_dirs:
        session_dirs = [user_data_dir]
    
    print(f"Excel: {excel_path}", flush=True)
    print(f"Sheet: {sheet_name}", flush=True)
    print(f"输出: {output_path}", flush=True)
    print(f"处理行数: {len(items)}", flush=True)
    print(f"使用 {len(session_dirs)} 个账号轮询:", flush=True)
    for i, sd in enumerate(session_dirs):
        print(f"  账号{i+1}: {sd}", flush=True)

    updated = 0
    skipped = 0
    failed = 0
    actual_output_path = output_path
    failed_items = []

    with sync_playwright() as p:
        contexts = []
        pages = []
        current_session_idx = 0
        
        print(f"\n同时打开 {len(session_dirs)} 个浏览器窗口...", flush=True)
        for i, session_dir in enumerate(session_dirs):
            print(f"  - 启动账号{i+1}的浏览器...", flush=True)
            context = p.chromium.launch_persistent_context(
                user_data_dir=session_dir,
                headless=args.headless,
                slow_mo=args.slow_mo,
                viewport={"width": 1440, "height": 1000},
            )
            page = context.pages[0] if context.pages else context.new_page()
            
            if args.login_wait:
                first_homepage_url = ""
                for item in items:
                    first_homepage_url = normalize_homepage_url(item.get("homepage_url"))
                    if first_homepage_url:
                        break
                login_url = first_homepage_url or "https://www.xiaohongshu.com/explore"
                page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2000)
                check_verification_popup(page)
            
            contexts.append(context)
            pages.append(page)
        
        print(f"所有 {len(session_dirs)} 个浏览器窗口已就绪\n", flush=True)
        
        def get_current_page():
            return pages[current_session_idx % len(pages)]

        try:
            for index, item in enumerate(items, start=1):
                row_idx = item["row"]
                title = item["title"]
                homepage_url = item["homepage_url"]
                old_url = item["old_url"]
                old_note_id = extract_note_id(old_url)

                if len(session_dirs) > 1:
                    expected_session_idx = (index - 1) % len(session_dirs)
                    if expected_session_idx != current_session_idx:
                        current_session_idx = expected_session_idx
                        print(f"  - [切换账号] → 账号{current_session_idx + 1}", flush=True)

                page = get_current_page()
                print(f"\n[{index}/{len(items)}] 第 {row_idx} 行: {title}", flush=True)
                print(f"  - 使用账号{current_session_idx + 1}", flush=True)
                
                if index > 1:
                    random_delay(800, 2000)
                
                batch_size = getattr(args, 'batch_size', 0)
                batch_interval = getattr(args, 'batch_interval', 30)
                if batch_size > 0 and index > 1 and (index - 1) % batch_size == 0:
                    print(f"  - [批次休息] 已处理 {index - 1} 条，休息 {batch_interval} 秒避免风控...", flush=True)
                    time.sleep(batch_interval)
                
                check_verification_popup(page)
                
                if skip_no_title and ("笔记暂未设置标题" in title or not title):
                    skipped += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": "笔记无标题(不可恢复)",
                    })
                    print("  - 跳过：笔记无标题，无法匹配", flush=True)
                    continue
                
                if not title or not homepage_url:
                    skipped += 1
                    print("  - 跳过：标题或主页链接为空", flush=True)
                    continue
                if args.only_empty and not is_blank(old_url):
                    skipped += 1
                    print("  - 跳过：笔记官方地址已有内容", flush=True)
                    continue

                try:
                    note_url = locate_note_url(
                        page,
                        homepage_url,
                        title,
                        max_scrolls=args.max_scrolls,
                        scroll_pixels=args.scroll_pixels,
                        detail_wait_ms=args.detail_wait_ms,
                        require_xsec_token=not args.allow_basic_url,
                        per_item_timeout_sec=args.per_item_timeout,
                        target_date=target_date,
                    )
                except PlaywrightTimeoutError as exc:
                    note_url = ""
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"查找超时: {exc}",
                    })
                    print(f"  - 查找超时: {exc}", flush=True)
                    continue
                except Exception as exc:
                    note_url = ""
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"查找失败: {exc}",
                    })
                    print(f"  - 查找失败: {exc}", flush=True)
                    continue

                if not note_url:
                    failed += 1
                    old_url_has_note_id = bool(extract_note_id(old_url))
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}, 原地址={old_url[:80]})",
                    })
                    print(f"  - 未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}), 保持原值", flush=True)
                    continue

                new_note_id = extract_note_id(note_url)
                if not new_note_id:
                    failed += 1
                    failed_items.append({
                        "row": row_idx,
                        "title": title,
                        "homepage_url": homepage_url,
                        "old_url": old_url,
                        "reason": f"地址没有笔记ID: {note_url}",
                    })
                    print(f"  - 找到的地址没有笔记ID，保持原值: {note_url}", flush=True)
                    continue

                ws.cell(row=row_idx, column=url_col).value = note_url
                updated += 1
                id_note = ""
                if old_note_id and new_note_id and old_note_id != new_note_id:
                    id_note = f"（注意：笔记ID从 {old_note_id} 变为 {new_note_id}）"
                print(f"  - 已更新: {note_url} {id_note}", flush=True)

                if args.save_every and updated % args.save_every == 0:
                    actual_output_path = save_workbook(wb, actual_output_path)
                    print(f"  - 已保存进度: {actual_output_path}", flush=True)
        except KeyboardInterrupt:
            print("\n\n用户中断，正在保存已处理数据...", flush=True)
            actual_output_path = save_workbook(wb, actual_output_path)
            print(f"已保存 {updated} 条更新到: {actual_output_path}", flush=True)
            raise
        finally:
            for ctx in contexts:
                try:
                    ctx.close()
                except Exception:
                    pass

    actual_output_path = save_workbook(wb, actual_output_path)
    print("\n处理完成。", flush=True)
    print(f"更新: {updated}，跳过: {skipped}，失败: {failed}", flush=True)
    print(f"结果文件: {actual_output_path}", flush=True)

    if failed_items:
        print("\n" + "=" * 60, flush=True)
        print("失败数据明细：", flush=True)
        print("=" * 60, flush=True)
        for item in failed_items:
            print(f"\n行号: {item['row']}", flush=True)
            print(f"  标题: {item['title']}", flush=True)
            print(f"  主页: {item['homepage_url']}", flush=True)
            print(f"  原地址: {item['old_url']}", flush=True)
            print(f"  失败原因: {item['reason']}", flush=True)

        failed_path = export_failed_items(failed_items, actual_output_path)
        if failed_path:
            print(f"\n失败数据已导出到: {failed_path}", flush=True)

    return actual_output_path


def build_arg_parser():
    default_target_date = (datetime.now() - timedelta(days=3)).strftime("%m.%d")
    
    parser = argparse.ArgumentParser(description="按主页链接+笔记标题，更新小红书笔记官方地址。")
    parser.add_argument("--excel", dest="excel_path", default=DEFAULT_EXCEL_PATH, help="输入 Excel 路径")
    parser.add_argument("--output", dest="output_path", default="", help="输出 Excel 路径，默认直接更新输入文件")
    parser.add_argument("--sheet", dest="sheet_name", default=DEFAULT_SHEET_NAME, help="Sheet 名")
    parser.add_argument("--rows", default="", help="行号，如 2 / 2-20 / 2,5,9-12；留空处理全部数据行")
    parser.add_argument("--limit", type=int, default=0, help="最多处理多少行，调试用")
    parser.add_argument("--only-empty", action="store_true", help="只处理笔记官方地址为空的行")
    parser.add_argument("--max-scrolls", type=int, default=15, help="每个主页最多向下滚动次数，默认 15")
    parser.add_argument("--scroll-pixels", type=int, default=900, help="每次滚动像素")
    parser.add_argument("--detail-wait-ms", type=int, default=8000, help="点击笔记后等待地址栏变化的毫秒数")
    parser.add_argument("--per-item-timeout", type=int, default=20, help="每条数据的最大处理时间（秒），超时后跳过")
    parser.add_argument("--save-every", type=int, default=20, help="每更新多少行保存一次，0 表示只最后保存")
    parser.add_argument("--login-wait", type=int, default=50, help="启动浏览器后等待登录/页面准备的秒数，默认 50")
    parser.add_argument("--allow-basic-url", action="store_true", help="没有 xsec_token 时允许写入基础 explore 地址")
    parser.add_argument("--headless", action="store_true", help="无头模式运行，不建议首次使用")
    parser.add_argument("--slow-mo", type=int, default=80, help="Playwright 操作延迟毫秒")
    parser.add_argument("--user-data-dir", default=DEFAULT_USER_DATA_DIR, help="浏览器会话目录")
    parser.add_argument("--parallel-index", type=int, default=-1, help="并行任务索引（0开始），-1表示非并行模式")
    parser.add_argument("--parallel-total", type=int, default=1, help="并行任务总数")
    parser.add_argument("--parallel-session", default="", help="并行会话ID，用于区分浏览器目录")
    parser.add_argument("--skip-no-title", action="store_true", default=True, help="跳过无标题笔记（默认开启）")
    parser.add_argument("--no-skip-no-title", action="store_true", help="不跳过无标题笔记")
    parser.add_argument("--target-date", default=default_target_date, help=f"目标日期（格式：MM.DD），默认 {default_target_date}（当前日期往前推3天）")
    parser.add_argument("--batch-size", type=int, default=30, help="每处理多少条后休息一次，0表示不休息，默认30")
    parser.add_argument("--batch-interval", type=int, default=30, help="批次休息秒数，默认30")
    parser.add_argument("--sessions", default="", help="多个session目录，用逗号分隔，用于每条数据轮询账号")
    parser.add_argument("--session-mode", choices=["rotate", "bind"], default="rotate", help="账号模式: rotate=所有任务共享所有账号, bind=每个任务使用分配到的账号组")
    return parser


def main():
    parser = build_arg_parser()
    args = parser.parse_args()
    try:
        update_note_urls(args)
    except KeyboardInterrupt:
        print("\n用户中断。", flush=True)
        raise SystemExit(130)
    except Exception as exc:
        print(f"错误: {exc}", flush=True)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
