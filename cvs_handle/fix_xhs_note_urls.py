import argparse
import csv
import logging
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
LOG_DIR = os.path.join(BASE_DIR, "logs")

_logger = None
_log_file_path = None


def setup_logger(log_path=None):
    global _logger, _log_file_path
    if _logger:
        return _logger
    os.makedirs(LOG_DIR, exist_ok=True)
    if not log_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(LOG_DIR, f"fix_urls_{ts}.log")
    _log_file_path = log_path
    os.makedirs(os.path.dirname(log_path), exist_ok=True)

    logger = logging.getLogger("fix_xhs_note_urls")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-5s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)
    _logger = logger
    logger.info(f"日志文件: {log_path}")
    return logger


def log(msg, level="info"):
    if _logger is None:
        setup_logger()
    getattr(_logger, level, _logger.info)(msg)


DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "小红书笔记列表_规范Excel版.xlsx")
DEFAULT_SHEET_NAME = "小红书笔记列表"
DEFAULT_USER_DATA_DIR = os.path.join(REPO_DIR, "browser_session")

TITLE_HEADER = "笔记标题"
HOMEPAGE_HEADER = "主页链接"
NOTE_URL_HEADER = "笔记官方地址"

ANTI_DETECTION_CONFIG = {
    "min_delay_ms": 1500,
    "max_delay_ms": 4000,
    "scroll_delay_min_ms": 800,
    "scroll_delay_max_ms": 2000,
    "batch_interval_default": 60,
    "batch_size_default": 15,
    "mouse_move_steps": 10,
    "random_scroll_pixels": True,
    "simulate_human_typing": True,
    "viewport_variance": True,
    "viewport_width_base": 1440,
    "viewport_height_base": 1000,
    "viewport_width_jitter": 90,
    "viewport_height_jitter": 70,
    "page_warmup_points_min": 2,
    "page_warmup_points_max": 4,
    "page_warmup_pause_min_ms": 300,
    "page_warmup_pause_max_ms": 900,
    "page_warmup_dwell_min_ms": 700,
    "page_warmup_dwell_max_ms": 1800,
    "verification_hit_suppression_sec": 15,
    "verification_cooldown_threshold": 2,
    "verification_cooldown_min_sec": 75,
    "verification_cooldown_max_sec": 135,
    "verification_cooldown_gap_sec": 180,
    "between_item_delay_min_ms": 900,
    "between_item_delay_max_ms": 2600,
    "micro_rest_every_min": 4,
    "micro_rest_every_max": 7,
    "micro_rest_min_sec": 8,
    "micro_rest_max_sec": 18,
    "failure_cooldown_threshold": 3,
    "failure_cooldown_min_sec": 45,
    "failure_cooldown_max_sec": 95,
    "failure_cooldown_gap_sec": 180,
    "session_batch_min": 10,
    "session_batch_max": 15,
    "session_switch_rest_min_sec": 60,
    "session_switch_rest_max_sec": 180,
    "mouse_pre_hover_probability": 0.45,
    "mouse_hover_pause_min_ms": 120,
    "mouse_hover_pause_max_ms": 420,
    "light_warmup_probability": 0.75,
    "light_warmup_points_min": 1,
    "light_warmup_points_max": 2,
    "light_warmup_pause_min_ms": 120,
    "light_warmup_pause_max_ms": 420,
    "light_warmup_dwell_min_ms": 150,
    "light_warmup_dwell_max_ms": 450,
}

FIXED_SESSION_BATCH_SIZE = 100
FIXED_SESSION_SWITCH_REST_SEC = 120

RISK_CONTROL_STATE = {
    "consecutive_verification_hits": 0,
    "last_verification_detected_at": 0.0,
    "last_cooldown_at": 0.0,
}

PACE_CONTROL_STATE = {
    "items_since_micro_rest": 0,
    "next_micro_rest_after": 0,
    "consecutive_failures": 0,
    "last_failure_cooldown_at": 0.0,
}

MOUSE_TRACK_STATE = {}
PAGE_VISIT_STATE = {}

SPEED_PROFILE = "default"


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
    
    match = re.search(r"(\d{1,2})[-月/.](\d{1,2})", date_str)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        try:
            return datetime(year=year, month=month, day=day)
        except ValueError:
            return None
    
    return None


def parse_target_date(value, now=None):
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    base = now or datetime.now()
    normalized = (
        text.replace("年", "-")
        .replace("月", "-")
        .replace("日", "")
        .replace("/", "-")
        .replace(".", "-")
    )
    match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", normalized)
    if match:
        year, month, day = match.groups()
        try:
            return datetime(int(year), int(month), int(day))
        except ValueError:
            return None
    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})", normalized)
    if match:
        month, day = match.groups()
        try:
            return datetime(base.year, int(month), int(day))
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
        log(f"  - 目标文件被占用，已改存到: {fallback_path}")
        return fallback_path


def goto_with_retry(page, url, timeout_ms=45000, retries=2):
    last_error = None
    for attempt in range(retries + 1):
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            status = resp.status if resp else "无响应对象"
            final_url = page.url
            log(f"  - 页面加载完成: status={status}, 最终URL={final_url[:100]}")
            return True
        except Exception as exc:
            last_error = exc
            if attempt >= retries:
                log(f"  - 页面打开失败(已重试{retries}次): {exc}")
                raise
            log(f"  - 页面打开失败，准备重试({attempt + 1}/{retries}): {exc}")
            page.wait_for_timeout(1200)
    if last_error:
        raise last_error
    return False


def click_notes_tab_if_visible(page):
    selectors = [
        "[role='tab']",
        "[role='tablist'] [role='tab']",
        ".reds-tabs-list .tab",
        ".user-tabs .tab",
        ".tabs-list .tab",
        "button:has-text('笔记')",
        "a:has-text('笔记')",
        "div:has-text('笔记')",
        "span:has-text('笔记')",
        "text=笔记",
    ]
    best_tab = None
    best_box = None
    for selector in selectors:
        try:
            tabs = page.locator(selector).filter(has_text="笔记")
            count = min(tabs.count(), 6)
            if count > 0:
                log(f"  - [笔记Tab] 选择器 '{selector}' 找到 {count} 个匹配")
            for index in range(count):
                tab = tabs.nth(index)
                if not tab.is_visible(timeout=300):
                    continue
                box = None
                try:
                    handle = tab.element_handle()
                    if handle:
                        box = handle.bounding_box()
                except Exception:
                    box = None
                if not box:
                    try:
                        tab.click(timeout=1500)
                        page.wait_for_timeout(800)
                        log(f"  - [笔记Tab] 已点击选择器 '{selector}' 第{index}个tab")
                        return True
                    except Exception:
                        continue
                if box.get("height", 0) <= 0 or box.get("width", 0) <= 0:
                    continue
                if box.get("y", 0) > 700 or box.get("height", 0) > 120 or box.get("width", 0) > 320:
                    continue
                if not best_box or (box["y"], box["x"]) < (best_box["y"], best_box["x"]):
                    best_tab = tab
                    best_box = box
        except Exception as exc:
            log(f"  - [笔记Tab] 选择器 '{selector}' 查找/点击异常: {exc}")
    if best_tab:
        try:
            best_tab.click(timeout=1500)
            page.wait_for_timeout(800)
            log(f"  - [笔记Tab] 已点击(最佳匹配) 坐标=({best_box.get('x',0):.0f},{best_box.get('y',0):.0f}) 尺寸={best_box.get('width',0):.0f}x{best_box.get('height',0):.0f}")
            return True
        except Exception as exc:
            log(f"  - [笔记Tab] 点击最佳匹配失败: {exc}")
    log(f"  - [笔记Tab] 未找到可见的笔记Tab (尝试了 {len(selectors)} 个选择器)")
    return False


def find_note_candidate(page, title):
    target = normalize_text(title)
    if not target:
        return None
    log(f"  - [查找卡片] 目标标题: raw='{title[:60]}' normalized='{target[:60]}' len={len(target)}")

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
                if (value.length >= 8 && target.length >= 8) {
                    if (value.slice(0, 8) === target.slice(0, 8)) return 65;
                    if (value.slice(-8) === target.slice(-8)) return 60;
                }
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
            log(f"  - [查找卡片] 找到匹配: score={best.get('score', 0)} noteId={best.get('noteId', '')} href={normalize_note_url(best.get('href', ''))[:80]} 坐标=({best.get('x', 0):.0f},{best.get('y', 0):.0f}) 尺寸={best.get('width', 0):.0f}x{best.get('height', 0):.0f}")
        else:
            log(f"  - [查找卡片] 未找到匹配: 页面锚点数={total_anchors}, 候选数={total_candidates}, 目标标题={target[:40]}")
        if top_scores:
            scores_desc = ", ".join(f"[score={s['score']} noteId={s['noteId']} text={s['text'][:20]}]" for s in top_scores)
            log(f"  - [查找卡片] 候选排名: {scores_desc}")
        return best
    log(f"  - [查找卡片] JS执行返回空, 目标标题={target[:40]}")
    return None


def find_note_candidate_by_note_id(page, note_id):
    note_id = str(note_id or "").strip()
    if not note_id:
        return None
    result = page.evaluate(
        """
        ({ noteId }) => {
            const cleanHref = (href) => {
                if (!href) return '';
                try {
                    return new URL(href, location.href).href;
                } catch (error) {
                    return href;
                }
            };
            const matchesNoteId = (href) => {
                const value = cleanHref(href);
                return value.includes(`/explore/${noteId}`) || value.includes(`/${noteId}?`) || value.endsWith(`/${noteId}`);
            };
            const extractNoteId = (href) => {
                const value = cleanHref(href);
                const match = value.match(/\\/explore\\/([^/?#]+)/)
                    || value.match(/\\/user\\/profile\\/[^/?#]+\\/([^/?#]+)/);
                return match ? match[1] : '';
            };
            const isClickableProfileCardHref = (href) => {
                const value = cleanHref(href);
                return value.includes('/user/profile/')
                    && value.includes('xsec_source=pc_user')
                    && value.includes('xsec_token=')
                    && !value.includes('xsec_token=&');
            };
            const isClickableCandidateHref = (href) => {
                const extracted = extractNoteId(href);
                return Boolean(extracted && extracted.length >= 8) || isClickableProfileCardHref(href);
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
                    if (rect.width >= 80 && rect.height >= 80) return node;
                }
                return anchor;
            };
            const buildCandidate = (anchor) => {
                const href = cleanHref(anchor.getAttribute('href') || anchor.href);
                const clickable = findClickableNode(anchor);
                const rect = clickable.getBoundingClientRect();
                const anchorRect = anchor.getBoundingClientRect();
                const targetRect = rect.width >= 20 && rect.height >= 20 ? rect : anchorRect;
                let score = 110;
                if (href.includes('xsec_token=')) score += 10;
                if (href.includes('xsec_source=pc_user')) score += 5;
                return {
                    href,
                    noteId: extractNoteId(href),
                    text: (anchor.innerText || anchor.textContent || '').trim(),
                    score,
                    x: targetRect.left + targetRect.width / 2,
                    y: targetRect.top + Math.min(targetRect.height * 0.42, targetRect.height / 2),
                    width: targetRect.width,
                    height: targetRect.height,
                };
            };
            const anchors = Array.from(document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]'))
                .filter((a) => {
                    const href = a.getAttribute('href') || a.href || '';
                    if (!matchesNoteId(href)) return false;
                    if (!isClickableCandidateHref(href)) return false;
                    const rect = a.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0;
                });
            const candidates = anchors.map(buildCandidate);
            candidates.sort((l, r) => r.score - l.score);
            const best = candidates[0] || null;
            return {
                best,
                totalAnchors: anchors.length,
                topScores: candidates.slice(0, 5).map((c) => ({ score: c.score, noteId: c.noteId, href: (c.href || '').substring(0, 80) })),
            };
        }
        """,
        {"noteId": note_id},
    )
    if result:
        best = result.get("best")
        total_anchors = result.get("totalAnchors", 0)
        top_scores = result.get("topScores", [])
        if best:
            log(f"  - [查找卡片] 按noteId找到: noteId={note_id} score={best.get('score', 0)} href={normalize_note_url(best.get('href', ''))[:80]} 坐标=({best.get('x', 0):.0f},{best.get('y', 0):.0f}) 尺寸={best.get('width', 0):.0f}x{best.get('height', 0):.0f}")
        else:
            log(f"  - [查找卡片] 按noteId未找到: noteId={note_id} 候选锚点数={total_anchors}")
        if top_scores:
            scores_desc = ", ".join(f"[score={s['score']} noteId={s['noteId']} href={s['href'][:30]}]" for s in top_scores)
            log(f"  - [查找卡片] noteId候选: {scores_desc}")
        return best
    log(f"  - [查找卡片] 按noteId JS执行返回空: noteId={note_id}")
    return None


def click_note_candidate(page, candidate):
    if not candidate or not is_clickable_candidate_link(candidate.get("href")):
        log(f"  - [点击卡片] 候选无效或链接不可点击: href={normalize_note_url(candidate.get('href', ''))[:80] if candidate else 'None'}")
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
        log(f"  - [点击卡片] 方式1-坐标点击: ({candidate['x']:.0f}, {candidate['y']:.0f}), viewport=({viewport_width}x{viewport_height})")
        human_like_click(page, float(candidate["x"]), float(candidate["y"]))
        return True

    log(f"  - [点击卡片] 方式1失败(坐标越界或为0: x={candidate.get('x', 0):.0f}, y={candidate.get('y', 0):.0f}, viewport={viewport_width}x{viewport_height}), 尝试方式2-JS查找元素点击")

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
        log(f"  - [点击卡片] 方式2-JS定位点击: ({point['x']:.0f}, {point['y']:.0f}), 尺寸={point.get('width', 0):.0f}x{point.get('height', 0):.0f}")
        page.wait_for_timeout(300)
        human_like_click(page, float(point["x"]), float(point["y"]))
        return True

    log(f"  - [点击卡片] 方式2失败(JS未找到有效坐标), 尝试方式3-直接anchor.click()")
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
    log(f"  - [点击卡片] 方式3-anchor.click()结果: {js_result}")
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
    poll_interval_ms = get_note_url_poll_interval_ms()
    log(f"  - [等待地址栏] 开始等待: 当前URL={first_url[:80]}, 期望noteId={expected_note_id}, 超时={timeout_ms}ms, 需要xsec_token={require_xsec_token}")
    while time.time() < deadline:
        current_url = normalize_note_url(page.url)
        current_note_id = extract_note_id(current_url)
        check_count += 1
        if is_valid_address_bar_note_url(current_url, require_xsec_token=False):
            if not expected_note_id or current_note_id == expected_note_id:
                best = current_url
                if is_valid_address_bar_note_url(current_url, require_xsec_token=require_xsec_token):
                    log(f"  - [等待地址栏] 成功: 检查{check_count}次, URL={current_url[:80]}")
                    return current_url
        page.wait_for_timeout(poll_interval_ms)
    elapsed = time.time() - (deadline - timeout_ms / 1000)
    if best:
        log(f"  - [等待地址栏] 超时{elapsed:.1f}s/{timeout_ms/1000:.1f}s 检查{check_count}次, 有基础URL但无xsec_token: {best[:80]}")
    else:
        final_url = normalize_note_url(page.url)
        final_note_id = extract_note_id(final_url)
        log(f"  - [等待地址栏] 超时{elapsed:.1f}s/{timeout_ms/1000:.1f}s 检查{check_count}次, 地址栏未跳转到笔记页, 最终URL={final_url[:80]}, noteId={final_note_id}")
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
            if RISK_CONTROL_STATE["consecutive_verification_hits"]:
                log("  - [风控检测] 页面恢复正常，清空连续风控计数", level="debug")
            RISK_CONTROL_STATE["consecutive_verification_hits"] = 0
            log(f"  - [风控检测] 无验证弹窗/登录弹窗，页面正常", level="debug")
            return True

        now = time.time()
        if now - RISK_CONTROL_STATE["last_verification_detected_at"] > ANTI_DETECTION_CONFIG["verification_hit_suppression_sec"]:
            RISK_CONTROL_STATE["consecutive_verification_hits"] += 1
            RISK_CONTROL_STATE["last_verification_detected_at"] = now
        log(
            f"  - [风控检测] 连续触发计数: {RISK_CONTROL_STATE['consecutive_verification_hits']}",
            level="debug",
        )
        maybe_take_risk_cooldown()

        if indicators.get("hasCaptcha"):
            log(f"  - [风控检测] 检测到验证弹窗！请在浏览器中手动完成验证，最多等待 {max_wait_sec} 秒...")
        if indicators.get("hasLoginModal"):
            log(f"  - [风控检测] 检测到登录弹窗！请在浏览器中完成登录，最多等待 {max_wait_sec} 秒...")

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
                    RISK_CONTROL_STATE["consecutive_verification_hits"] = 0
                    log("  - [风控检测] 验证/登录已完成，继续执行")
                    page.wait_for_timeout(random.randint(1000, 2000))
                    return True
            except Exception:
                pass
            remaining = int(deadline - time.time())
            if remaining > 0 and remaining % 10 == 0:
                log(f"  - [风控检测] 仍在等待验证... 剩余 {remaining} 秒")

        log("  - [风控检测] 验证等待超时，继续执行（可能仍被限制）")
        return False
    except Exception as exc:
        log(f"  - [风控检测] 检测异常: {exc}")
        return True


def random_delay(min_ms=None, max_ms=None):
    if min_ms is None:
        min_ms = ANTI_DETECTION_CONFIG["min_delay_ms"]
    if max_ms is None:
        max_ms = ANTI_DETECTION_CONFIG["max_delay_ms"]
    delay = random.randint(min_ms, max_ms)
    time.sleep(delay / 1000.0)


def is_turbo_speed_mode(speed_mode):
    return str(speed_mode or "").strip().lower() in {"turbo", "fastest"}


def format_duration(seconds):
    try:
        total = int(max(0, seconds))
    except Exception:
        total = 0
    hours = total // 3600
    minutes = (total % 3600) // 60
    sec = total % 60
    if hours:
        return f"{hours}:{minutes:02d}:{sec:02d}"
    return f"{minutes:02d}:{sec:02d}"


def reset_pace_state():
    PACE_CONTROL_STATE["items_since_micro_rest"] = 0
    PACE_CONTROL_STATE["next_micro_rest_after"] = random.randint(
        ANTI_DETECTION_CONFIG["micro_rest_every_min"],
        ANTI_DETECTION_CONFIG["micro_rest_every_max"],
    )
    PACE_CONTROL_STATE["consecutive_failures"] = 0
    PACE_CONTROL_STATE["last_failure_cooldown_at"] = 0.0


def apply_between_item_pacing():
    random_delay(
        ANTI_DETECTION_CONFIG["between_item_delay_min_ms"],
        ANTI_DETECTION_CONFIG["between_item_delay_max_ms"],
    )
    PACE_CONTROL_STATE["items_since_micro_rest"] += 1
    threshold = PACE_CONTROL_STATE["next_micro_rest_after"]
    if threshold <= 0:
        threshold = random.randint(
            ANTI_DETECTION_CONFIG["micro_rest_every_min"],
            ANTI_DETECTION_CONFIG["micro_rest_every_max"],
        )
        PACE_CONTROL_STATE["next_micro_rest_after"] = threshold

    if PACE_CONTROL_STATE["items_since_micro_rest"] < threshold:
        return

    rest_sec = random.randint(
        ANTI_DETECTION_CONFIG["micro_rest_min_sec"],
        ANTI_DETECTION_CONFIG["micro_rest_max_sec"],
    )
    log(f"  - [节奏控制] 已连续处理 {PACE_CONTROL_STATE['items_since_micro_rest']} 条，短休息 {rest_sec} 秒")
    time.sleep(rest_sec)
    PACE_CONTROL_STATE["items_since_micro_rest"] = 0
    PACE_CONTROL_STATE["next_micro_rest_after"] = random.randint(
        ANTI_DETECTION_CONFIG["micro_rest_every_min"],
        ANTI_DETECTION_CONFIG["micro_rest_every_max"],
    )


def register_processing_outcome(success, row_idx=None, reason=""):
    if success:
        if PACE_CONTROL_STATE["consecutive_failures"]:
            log(
                f"  - [节奏控制] 连续失败计数已清零（此前 {PACE_CONTROL_STATE['consecutive_failures']} 次）",
                level="debug",
            )
        PACE_CONTROL_STATE["consecutive_failures"] = 0
        return

    PACE_CONTROL_STATE["consecutive_failures"] += 1
    log(
        f"  - [节奏控制] 连续失败计数: {PACE_CONTROL_STATE['consecutive_failures']}"
        + (f"，行{row_idx}" if row_idx else "")
        + (f"，原因: {reason}" if reason else ""),
        level="debug",
    )
    maybe_take_failure_cooldown()


def maybe_take_failure_cooldown():
    failures = PACE_CONTROL_STATE["consecutive_failures"]
    if failures < ANTI_DETECTION_CONFIG["failure_cooldown_threshold"]:
        return

    now = time.time()
    if now - PACE_CONTROL_STATE["last_failure_cooldown_at"] < ANTI_DETECTION_CONFIG["failure_cooldown_gap_sec"]:
        return

    cooldown_sec = random.randint(
        ANTI_DETECTION_CONFIG["failure_cooldown_min_sec"],
        ANTI_DETECTION_CONFIG["failure_cooldown_max_sec"],
    )
    PACE_CONTROL_STATE["last_failure_cooldown_at"] = now
    log(f"  - [失败保护] 连续失败 {failures} 次，自动冷却 {cooldown_sec} 秒后继续")
    time.sleep(cooldown_sec)


def get_session_rotation_settings(args):
    batch_min = max(1, int(getattr(args, "session_batch_min", ANTI_DETECTION_CONFIG["session_batch_min"])))
    batch_max = max(batch_min, int(getattr(args, "session_batch_max", ANTI_DETECTION_CONFIG["session_batch_max"])))
    rest_min = max(0, int(getattr(args, "session_switch_rest_min", ANTI_DETECTION_CONFIG["session_switch_rest_min_sec"])))
    rest_max = max(rest_min, int(getattr(args, "session_switch_rest_max", ANTI_DETECTION_CONFIG["session_switch_rest_max_sec"])))
    return batch_min, batch_max, rest_min, rest_max


def draw_session_batch_size(batch_min, batch_max):
    return random.randint(batch_min, batch_max)


def draw_session_switch_rest(rest_min, rest_max):
    return random.randint(rest_min, rest_max)


def get_mouse_position(page):
    state = MOUSE_TRACK_STATE.get(id(page))
    if state:
        return state["x"], state["y"]

    viewport = page.viewport_size or build_human_viewport()
    width = max(int(viewport.get("width", 0) or 0), 1280)
    height = max(int(viewport.get("height", 0) or 0), 860)
    x = random.randint(int(width * 0.35), int(width * 0.65))
    y = random.randint(int(height * 0.28), int(height * 0.62))
    MOUSE_TRACK_STATE[id(page)] = {"x": x, "y": y}
    return x, y


def remember_mouse_position(page, x, y):
    MOUSE_TRACK_STATE[id(page)] = {"x": int(x), "y": int(y)}


def clamp_point(page, x, y):
    viewport = page.viewport_size or build_human_viewport()
    width = max(int(viewport.get("width", 0) or 0), 1280)
    height = max(int(viewport.get("height", 0) or 0), 860)
    safe_x = max(5, min(int(x), width - 5))
    safe_y = max(5, min(int(y), height - 5))
    return safe_x, safe_y


def move_mouse_naturally(page, target_x, target_y, min_steps=8, max_steps=18):
    start_x, start_y = get_mouse_position(page)
    end_x, end_y = clamp_point(page, target_x, target_y)
    steps = random.randint(min_steps, max_steps)

    for i in range(steps):
        progress = (i + 1) / steps
        eased = 1 - (1 - progress) * (1 - progress)
        jitter_scale = max(1, int((1 - progress) * 7))
        x = start_x + (end_x - start_x) * eased + random.randint(-jitter_scale, jitter_scale)
        y = start_y + (end_y - start_y) * eased + random.randint(-jitter_scale, jitter_scale)
        move_x, move_y = clamp_point(page, x, y)
        page.mouse.move(move_x, move_y)
        remember_mouse_position(page, move_x, move_y)
        time.sleep(random.randint(12, 55) / 1000.0)

    page.mouse.move(end_x, end_y)
    remember_mouse_position(page, end_x, end_y)


def human_like_scroll(page, pixels=None):
    if pixels is None:
        pixels = random.randint(300, 800)

    if SPEED_PROFILE == "turbo":
        steps = random.randint(1, 2)
    else:
        steps = random.randint(3, 6)
    pixels_per_step = pixels // steps
    remainder = pixels % steps
    
    for i in range(steps):
        step_pixels = pixels_per_step + (remainder if i == steps - 1 else 0)
        page.mouse.wheel(0, step_pixels)
        delay = random.randint(
            ANTI_DETECTION_CONFIG["scroll_delay_min_ms"],
            ANTI_DETECTION_CONFIG["scroll_delay_max_ms"]
        )
        time.sleep(delay / 1000.0)


def get_note_url_poll_interval_ms():
    if SPEED_PROFILE == "turbo":
        return 80
    if SPEED_PROFILE == "fast_batch":
        return 140
    return 250


def get_post_navigation_settle_wait_ms():
    if SPEED_PROFILE == "turbo":
        return random.randint(120, 260)
    if SPEED_PROFILE == "fast_batch":
        return random.randint(220, 420)
    return random.randint(600, 1200)


def get_post_selector_settle_wait_ms():
    if SPEED_PROFILE == "turbo":
        return random.randint(80, 180)
    if SPEED_PROFILE == "fast_batch":
        return random.randint(160, 320)
    return random.randint(500, 1000)


def get_scroll_stuck_wait_ms():
    if SPEED_PROFILE == "turbo":
        return 100
    if SPEED_PROFILE == "fast_batch":
        return 180
    return 400


def get_effective_slow_mo(args):
    slow_mo = int(getattr(args, "slow_mo", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")) and slow_mo == 80:
        return 0
    return slow_mo


def get_effective_detail_wait_ms(args):
    detail_wait_ms = int(getattr(args, "detail_wait_ms", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")) and detail_wait_ms >= 8000:
        return 2500
    return detail_wait_ms


def get_effective_max_scrolls(args):
    max_scrolls = int(getattr(args, "max_scrolls", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")) and max_scrolls >= 12:
        return 8
    return max_scrolls


def get_effective_per_item_timeout(args):
    per_item_timeout = int(getattr(args, "per_item_timeout", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")) and per_item_timeout >= 20:
        return 12
    return per_item_timeout


def get_effective_batch_pause(args):
    batch_size = int(getattr(args, "batch_size", 0) or 0)
    batch_interval = int(getattr(args, "batch_interval", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")):
        if batch_size > 0 and batch_interval >= 20 and batch_size <= 30:
            log(
                f"  - [速度模式] turbo 下已忽略重批次休息配置: batch_size={batch_size}, batch_interval={batch_interval}",
                level="debug",
            )
            return 0, 0
        if batch_size > 0 and batch_size <= 30 and batch_interval > 3:
            batch_interval = 3
    return batch_size, batch_interval


def get_effective_save_every(args):
    save_every = int(getattr(args, "save_every", 0) or 0)
    if is_turbo_speed_mode(getattr(args, "speed_mode", "")) and save_every == 20:
        log("  - [速度模式] turbo 下关闭过程高频保存，仅在结束/中断时保存", level="debug")
        return 0
    if str(getattr(args, "speed_mode", "")).strip().lower() == "auto" and save_every == 20:
        return 100
    return save_every


def build_human_viewport():
    width = ANTI_DETECTION_CONFIG["viewport_width_base"]
    height = ANTI_DETECTION_CONFIG["viewport_height_base"]
    if ANTI_DETECTION_CONFIG["viewport_variance"]:
        width += random.randint(
            -ANTI_DETECTION_CONFIG["viewport_width_jitter"],
            ANTI_DETECTION_CONFIG["viewport_width_jitter"],
        )
        height += random.randint(
            -ANTI_DETECTION_CONFIG["viewport_height_jitter"],
            ANTI_DETECTION_CONFIG["viewport_height_jitter"],
        )
    return {"width": max(width, 1280), "height": max(height, 860)}


def choose_warmup_mode(page, homepage_url):
    state = PAGE_VISIT_STATE.setdefault(id(page), {})
    last_homepage = state.get("last_homepage_url", "")
    mode = "full"
    if SPEED_PROFILE == "fast_batch":
        mode = "light" if random.random() < 0.88 else "skip"
    elif SPEED_PROFILE == "turbo":
        mode = "skip" if random.random() < 0.92 else "light"
    else:
        mode = "full"
    if homepage_url and homepage_url == last_homepage:
        mode = "light" if random.random() < ANTI_DETECTION_CONFIG["light_warmup_probability"] else "skip"
    state["last_homepage_url"] = homepage_url
    return mode


def apply_speed_profile_for_large_batch(total_items, session_mode="rotate", speed_mode="auto"):
    global SPEED_PROFILE
    speed_mode = str(speed_mode or "auto").strip().lower()
    if speed_mode in {"off", "default", "normal"}:
        SPEED_PROFILE = "default"
        return
    if speed_mode in {"turbo", "fastest"}:
        SPEED_PROFILE = "turbo"
        ANTI_DETECTION_CONFIG["between_item_delay_min_ms"] = 10
        ANTI_DETECTION_CONFIG["between_item_delay_max_ms"] = 60
        ANTI_DETECTION_CONFIG["micro_rest_every_min"] = 80
        ANTI_DETECTION_CONFIG["micro_rest_every_max"] = 140
        ANTI_DETECTION_CONFIG["micro_rest_min_sec"] = 0
        ANTI_DETECTION_CONFIG["micro_rest_max_sec"] = 1
        ANTI_DETECTION_CONFIG["scroll_delay_min_ms"] = 80
        ANTI_DETECTION_CONFIG["scroll_delay_max_ms"] = 220
        ANTI_DETECTION_CONFIG["page_warmup_points_min"] = 0
        ANTI_DETECTION_CONFIG["page_warmup_points_max"] = 1
        ANTI_DETECTION_CONFIG["page_warmup_pause_min_ms"] = 20
        ANTI_DETECTION_CONFIG["page_warmup_pause_max_ms"] = 80
        ANTI_DETECTION_CONFIG["page_warmup_dwell_min_ms"] = 0
        ANTI_DETECTION_CONFIG["page_warmup_dwell_max_ms"] = 80
        ANTI_DETECTION_CONFIG["light_warmup_probability"] = 0.995
        ANTI_DETECTION_CONFIG["light_warmup_points_min"] = 0
        ANTI_DETECTION_CONFIG["light_warmup_points_max"] = 1
        ANTI_DETECTION_CONFIG["light_warmup_pause_min_ms"] = 20
        ANTI_DETECTION_CONFIG["light_warmup_pause_max_ms"] = 60
        ANTI_DETECTION_CONFIG["light_warmup_dwell_min_ms"] = 0
        ANTI_DETECTION_CONFIG["light_warmup_dwell_max_ms"] = 60
        ANTI_DETECTION_CONFIG["mouse_pre_hover_probability"] = 0.05
        ANTI_DETECTION_CONFIG["failure_cooldown_threshold"] = 6
        ANTI_DETECTION_CONFIG["failure_cooldown_min_sec"] = 8
        ANTI_DETECTION_CONFIG["failure_cooldown_max_sec"] = 18
        ANTI_DETECTION_CONFIG["failure_cooldown_gap_sec"] = 60
        ANTI_DETECTION_CONFIG["verification_cooldown_threshold"] = 4
        ANTI_DETECTION_CONFIG["verification_cooldown_min_sec"] = 20
        ANTI_DETECTION_CONFIG["verification_cooldown_max_sec"] = 45
        ANTI_DETECTION_CONFIG["verification_cooldown_gap_sec"] = 90
        if session_mode == "rotate":
            ANTI_DETECTION_CONFIG["session_switch_rest_min_sec"] = 5
            ANTI_DETECTION_CONFIG["session_switch_rest_max_sec"] = 15
        log("  - [速度模式] 已启用 turbo：轻风控，优先提速")
        return
    if total_items < 800:
        SPEED_PROFILE = "default"
        return
    SPEED_PROFILE = "fast_batch"
    ANTI_DETECTION_CONFIG["between_item_delay_min_ms"] = 320
    ANTI_DETECTION_CONFIG["between_item_delay_max_ms"] = 950
    ANTI_DETECTION_CONFIG["micro_rest_every_min"] = 10
    ANTI_DETECTION_CONFIG["micro_rest_every_max"] = 16
    ANTI_DETECTION_CONFIG["micro_rest_min_sec"] = 4
    ANTI_DETECTION_CONFIG["micro_rest_max_sec"] = 10
    ANTI_DETECTION_CONFIG["page_warmup_points_min"] = 1
    ANTI_DETECTION_CONFIG["page_warmup_points_max"] = 2
    ANTI_DETECTION_CONFIG["page_warmup_pause_min_ms"] = 120
    ANTI_DETECTION_CONFIG["page_warmup_pause_max_ms"] = 420
    ANTI_DETECTION_CONFIG["page_warmup_dwell_min_ms"] = 250
    ANTI_DETECTION_CONFIG["page_warmup_dwell_max_ms"] = 650
    ANTI_DETECTION_CONFIG["light_warmup_probability"] = 0.92
    ANTI_DETECTION_CONFIG["light_warmup_points_min"] = 1
    ANTI_DETECTION_CONFIG["light_warmup_points_max"] = 1
    ANTI_DETECTION_CONFIG["light_warmup_pause_min_ms"] = 80
    ANTI_DETECTION_CONFIG["light_warmup_pause_max_ms"] = 240
    ANTI_DETECTION_CONFIG["light_warmup_dwell_min_ms"] = 120
    ANTI_DETECTION_CONFIG["light_warmup_dwell_max_ms"] = 320
    if session_mode == "rotate":
        ANTI_DETECTION_CONFIG["session_switch_rest_min_sec"] = 35
        ANTI_DETECTION_CONFIG["session_switch_rest_max_sec"] = 90
    log("  - [速度模式] 已启用 fast_batch：大批量提速（偏保守）")


def warm_up_page(page, mode="full"):
    viewport = page.viewport_size or build_human_viewport()
    width = max(int(viewport.get("width", 0) or 0), 1280)
    height = max(int(viewport.get("height", 0) or 0), 860)

    if mode == "skip":
        log("  - [反检测] 同主页连续处理，本次跳过页面预热", level="debug")
        return

    if mode == "light":
        point_count = random.randint(
            ANTI_DETECTION_CONFIG["light_warmup_points_min"],
            ANTI_DETECTION_CONFIG["light_warmup_points_max"],
        )
        pause_min = ANTI_DETECTION_CONFIG["light_warmup_pause_min_ms"]
        pause_max = ANTI_DETECTION_CONFIG["light_warmup_pause_max_ms"]
        dwell_ms = random.randint(
            ANTI_DETECTION_CONFIG["light_warmup_dwell_min_ms"],
            ANTI_DETECTION_CONFIG["light_warmup_dwell_max_ms"],
        )
    else:
        point_count = random.randint(
            ANTI_DETECTION_CONFIG["page_warmup_points_min"],
            ANTI_DETECTION_CONFIG["page_warmup_points_max"],
        )
        pause_min = ANTI_DETECTION_CONFIG["page_warmup_pause_min_ms"]
        pause_max = ANTI_DETECTION_CONFIG["page_warmup_pause_max_ms"]
        dwell_ms = random.randint(
            ANTI_DETECTION_CONFIG["page_warmup_dwell_min_ms"],
            ANTI_DETECTION_CONFIG["page_warmup_dwell_max_ms"],
        )

    for _ in range(point_count):
        if mode == "light":
            x = random.randint(int(width * 0.42), int(width * 0.78))
            y = random.randint(int(height * 0.22), int(height * 0.58))
        else:
            x = random.randint(int(width * 0.18), int(width * 0.82))
            y = random.randint(int(height * 0.18), int(height * 0.72))
        move_mouse_naturally(page, x, y, min_steps=5 if mode == "light" else 8, max_steps=10 if mode == "light" else 16)
        time.sleep(random.randint(pause_min, pause_max) / 1000.0)

    page.wait_for_timeout(dwell_ms)
    log(f"  - [反检测] 页面预热完成: mode={mode} 移动{point_count}次, 停留{dwell_ms}ms", level="debug")


def human_like_click(page, x, y):
    target_x, target_y = clamp_point(page, x, y)
    if random.random() < ANTI_DETECTION_CONFIG["mouse_pre_hover_probability"]:
        hover_x, hover_y = clamp_point(
            page,
            target_x + random.randint(-22, 22),
            target_y + random.randint(-16, 16),
        )
        move_mouse_naturally(page, hover_x, hover_y)
        time.sleep(
            random.randint(
                ANTI_DETECTION_CONFIG["mouse_hover_pause_min_ms"],
                ANTI_DETECTION_CONFIG["mouse_hover_pause_max_ms"],
            )
            / 1000.0
        )

    move_mouse_naturally(page, target_x, target_y, min_steps=4, max_steps=max(6, ANTI_DETECTION_CONFIG["mouse_move_steps"]))
    time.sleep(random.randint(60, 180) / 1000.0)
    page.mouse.click(target_x, target_y, delay=random.randint(40, 110))
    remember_mouse_position(page, target_x, target_y)


STEALTH_INIT_SCRIPT = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
Object.defineProperty(navigator, 'languages', { get: () => ['zh-CN', 'zh', 'en'] });
if (!window.chrome) {
    window.chrome = {};
}
if (!window.chrome.runtime) {
    window.chrome.runtime = {};
}
if (navigator.permissions && navigator.permissions.query) {
    const originalQuery = navigator.permissions.query.bind(navigator.permissions);
    navigator.permissions.query = (parameters) => (
        parameters && parameters.name === 'notifications'
            ? Promise.resolve({ state: Notification.permission })
            : originalQuery(parameters)
    );
}
"""


def inject_stealth_scripts(context, page=None):
    try:
        context.add_init_script(script=STEALTH_INIT_SCRIPT)
        log("  - [反检测] 已注册启动前隐身脚本", level="debug")
    except Exception as exc:
        log(f"  - [反检测] 注册启动前隐身脚本失败: {exc}", level="debug")

    if page is None:
        return

    try:
        page.evaluate(f"() => {{ {STEALTH_INIT_SCRIPT} }}")
        log("  - [反检测] 已补注入当前页面隐身脚本", level="debug")
    except Exception as exc:
        log(f"  - [反检测] 补注入当前页面隐身脚本失败: {exc}", level="debug")


def maybe_take_risk_cooldown():
    hits = RISK_CONTROL_STATE["consecutive_verification_hits"]
    if hits < ANTI_DETECTION_CONFIG["verification_cooldown_threshold"]:
        return

    now = time.time()
    if now - RISK_CONTROL_STATE["last_cooldown_at"] < ANTI_DETECTION_CONFIG["verification_cooldown_gap_sec"]:
        return

    cooldown_sec = random.randint(
        ANTI_DETECTION_CONFIG["verification_cooldown_min_sec"],
        ANTI_DETECTION_CONFIG["verification_cooldown_max_sec"],
    )
    RISK_CONTROL_STATE["last_cooldown_at"] = now
    log(f"  - [风控熔断] 连续检测到验证/登录弹窗 {hits} 次，自动休息 {cooldown_sec} 秒")
    time.sleep(cooldown_sec)


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
    note_id_hint=None,
):
    homepage_url = normalize_homepage_url(homepage_url)
    target_date_str = target_date.strftime("%Y-%m-%d") if target_date else "未设置"
    note_id_hint = str(note_id_hint or "").strip() or None
    log(f"  - [定位流程] 开始: 主页={homepage_url} 查找标题={title[:40]} noteIdHint={note_id_hint or '无'} max_scrolls={max_scrolls} detail_wait_ms={detail_wait_ms} require_xsec_token={require_xsec_token} per_item_timeout={per_item_timeout_sec}s target_date={target_date_str}")
    nav_start = time.time()
    goto_timeout_ms = 45000
    if per_item_timeout_sec and per_item_timeout_sec > 0:
        goto_timeout_ms = min(45000, max(15000, int(per_item_timeout_sec * 1000)))
    if note_id_hint:
        try:
            direct_url = f"https://www.xiaohongshu.com/explore/{note_id_hint}"
            goto_with_retry(page, direct_url, timeout_ms=min(20000, goto_timeout_ms), retries=0)
            if check_verification_popup(page):
                resolved_url = wait_for_current_note_url(
                    page,
                    expected_note_id=note_id_hint,
                    timeout_ms=detail_wait_ms,
                    require_xsec_token=require_xsec_token,
                )
                if resolved_url:
                    log(f"  - [定位流程] 直达笔记成功: {resolved_url}")
                    return resolved_url
        except Exception as exc:
            log(f"  - [定位流程] 直达笔记失败，回退主页查找: {exc}", level="debug")
    goto_with_retry(page, homepage_url, timeout_ms=goto_timeout_ms, retries=1 if per_item_timeout_sec and per_item_timeout_sec <= 30 else 2)
    page.wait_for_timeout(get_post_navigation_settle_wait_ms())
    warm_up_mode = choose_warmup_mode(page, homepage_url)
    warm_up_page(page, mode=warm_up_mode)
    nav_elapsed = time.time() - nav_start
    log(f"  - [定位流程] 导航耗时: {nav_elapsed:.1f}s (goto_timeout_ms={goto_timeout_ms})", level="debug")
    
    page_status = page.evaluate("""
    () => {
        const body = document.body || {};
        const text = (body.innerText || '').substring(0, 300);
        const hasLoginPrompt = text.includes('登录') && (text.includes('注册') || text.includes('手机号'));
        const hasCaptcha = text.includes('请完成验证') || text.includes('请通过验证') || text.includes('滑动验证');
        const hasNoteCards = document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]').length;
        const currentUrl = location.href;
        const title = document.title || '';
        return { hasLoginPrompt, hasCaptcha, hasNoteCards, currentUrl, title, textPreview: text.substring(0, 150) };
    }
    """)
    log(f"  - [页面状态] URL={page_status.get('currentUrl', '')[:80]} 标题='{page_status.get('title', '')}' 登录提示={page_status.get('hasLoginPrompt')} 验证码={page_status.get('hasCaptcha')} 笔记卡片数={page_status.get('hasNoteCards')}", level="debug")
    if page_status.get("hasLoginPrompt"):
        log(f"  - [页面状态] ⚠️ 检测到登录提示! 页面文本: {page_status.get('textPreview', '')[:100]}", level="warning")
    if page_status.get("hasCaptcha"):
        log(f"  - [页面状态] ⚠️ 检测到验证码! 页面文本: {page_status.get('textPreview', '')[:100]}", level="warning")
    
    verification_ok = check_verification_popup(page)
    if not verification_ok:
        log(f"  - [定位流程] 验证/登录未完成，页面可能无法正常加载笔记，跳过此条")
        return ""
    
    tab_clicked = click_notes_tab_if_visible(page)
    if not tab_clicked:
        log(f"  - [定位流程] 未点击到笔记Tab，页面可能不是用户主页或Tab结构已变")
    
    try:
        selector_timeout_ms = 1500 if SPEED_PROFILE == "turbo" else 2500 if SPEED_PROFILE == "fast_batch" else 5000
        page.wait_for_selector('a[href*="/explore/"], a[href*="/user/profile/"]', timeout=selector_timeout_ms)
        page.wait_for_timeout(get_post_selector_settle_wait_ms())
    except Exception:
        log(f"  - [定位流程] 等待页面元素超时，继续尝试查找")

    search_start = time.time()
    last_height = 0
    no_change_count = 0
    max_no_change = 3
    scroll_count = 0
    early_stop_due_to_date = False

    for scroll_index in range(max_scrolls + 1):
        elapsed = time.time() - search_start
        if elapsed > per_item_timeout_sec:
            log(f"  - [定位流程] 已超时 ({elapsed:.1f}秒 > {per_item_timeout_sec}秒)，停止查找 (已滚动{scroll_count}次)")
            return ""

        candidate = None
        if note_id_hint:
            candidate = find_note_candidate_by_note_id(page, note_id_hint)
        if not candidate:
            candidate = find_note_candidate(page, title)
        if candidate and candidate.get("href"):
            log(f"  - [定位流程] 滚动第{scroll_index}次: 找到候选 score={candidate.get('score', 0)} href={normalize_note_url(candidate['href'])[:80]}")
            href = normalize_note_url(candidate["href"])
            note_id = extract_note_id(href)
            if not note_id and not is_clickable_candidate_link(href):
                log(f"  - [定位流程] 候选链接无效(无noteId且非可点击卡片): href={href[:100]}，可能页面未登录或被风控")
            elif is_clickable_candidate_link(href):
                log(f"  - [定位流程] 找到匹配卡片，准备点击: {href} noteId={note_id}")
                clicked = False
                try:
                    clicked = click_note_candidate(page, candidate)
                except Exception as exc:
                    log(f"  - [定位流程] 坐标点击失败，尝试链接点击: {exc}")
                    try:
                        clicked = click_note_by_href(page, href)
                    except Exception as exc2:
                        log(f"  - [定位流程] 链接点击也失败: {exc2}")
                        clicked = False

                if clicked:
                    resolved_url = wait_for_current_note_url(
                        page,
                        expected_note_id=note_id,
                        timeout_ms=detail_wait_ms,
                        require_xsec_token=require_xsec_token,
                    )
                    if resolved_url:
                        log(f"  - [定位流程] 成功获取地址: {resolved_url}")
                        return resolved_url
                    if is_valid_address_bar_note_url(href, require_xsec_token=require_xsec_token):
                        log(f"  - [定位流程] 地址栏未变化，使用卡片链接: {href}")
                        return href
                    current_page_url = normalize_note_url(page.url)
                    has_note_id = bool(extract_note_id(current_page_url))
                    log(f"  - [定位流程] 点击后未拿到xsec_token地址: 当前URL={current_page_url[:100]} 是否含笔记ID={has_note_id} 卡片href={href[:80]}")
                    return ""
                log(f"  - [定位流程] 找到卡片但点击失败: href={normalize_note_url(candidate.get('href', ''))[:80]}")
                return ""
            log(f"  - [定位流程] 忽略无效候选地址(非可点击链接): {href[:80]}")
        elif candidate:
            log(f"  - [定位流程] 滚动第{scroll_index}次: 找到候选但无有效href score={candidate.get('score', 0)} text={str(candidate.get('text', ''))[:30]}")
        else:
            if scroll_index % 3 == 0:
                log(f"  - [定位流程] 滚动第{scroll_index}次: 未找到任何候选")
                if scroll_index == 0:
                    try:
                        diag = page.evaluate("""
                        () => {
                            const body = (document.body || {}).innerText || '';
                            const hasLogin = body.includes('登录') && (body.includes('注册') || body.includes('手机号'));
                            const allAnchors = document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]').length;
                            const validAnchors = Array.from(document.querySelectorAll('a[href*="/explore/"], a[href*="/user/profile/"]')).filter(a => {
                                const href = a.getAttribute('href') || a.href || '';
                                const noteId = href.match(/\\/explore\\/([^/?#]+)/)?.[1] || href.match(/\\/user\\/profile\\/[^/?#]+\\/([^/?#]+)/)?.[1] || '';
                                return noteId && noteId.length >= 8;
                            }).length;
                            return { hasLogin, allAnchors, validAnchors, url: location.href };
                        }
                        """)
                        log(f"  - [诊断] 首次未找到候选: 登录提示={diag.get('hasLogin')} 全部锚点={diag.get('allAnchors')} 有效锚点={diag.get('validAnchors')} 当前URL={diag.get('url','')[:80]}", level="debug")
                        if diag.get("hasLogin"):
                            log(f"  - [诊断] ⚠️ 页面可能未登录，导致笔记卡片不可见", level="warning")
                        if diag.get("allAnchors", 0) > 0 and diag.get("validAnchors", 0) == 0:
                            log(f"  - [诊断] ⚠️ 页面有{diag.get('allAnchors')}个锚点但无有效笔记链接，可能session失效或被风控", level="warning")
                    except Exception as exc:
                        log(f"  - [诊断] 获取页面诊断信息失败: {exc}", level="debug")

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
                            if (txt && /\\d{1,2}[-月\\/.]\\d{1,2}/.test(txt)) {
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
                        log(f"  - [定位流程] 检测到笔记日期早于目标日期: 最早日期={min_date.strftime('%Y-%m-%d')} < 目标日期={target_date.strftime('%Y-%m-%d')}, 提前停止滚动")
                        early_stop_due_to_date = True
                        break

        current_height = page.evaluate(
            "() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        if scroll_index >= max_scrolls:
            break
        if ANTI_DETECTION_CONFIG["random_scroll_pixels"]:
            human_like_scroll(page, random.randint(int(scroll_pixels * 0.6), int(scroll_pixels * 1.4)))
        else:
            human_like_scroll(page, scroll_pixels)
        scroll_count += 1

        next_height = page.evaluate(
            "() => Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        if next_height == last_height == current_height:
            no_change_count += 1
            if no_change_count >= max_no_change:
                log(f"  - [定位流程] 页面已到底部，连续 {no_change_count} 次无变化，提前退出 (已滚动{scroll_count}次, 页面高度={current_height})")
                break
            page.wait_for_timeout(get_scroll_stuck_wait_ms())
        else:
            no_change_count = 0
        last_height = next_height

    total_elapsed = time.time() - search_start
    if early_stop_due_to_date:
        log(f"  - [定位流程] 失败(日期早于目标): 滚动{scroll_count}次后未找到匹配卡片, 耗时{total_elapsed:.1f}s, title={title[:40]}")
    else:
        log(f"  - [定位流程] 失败: 滚动{scroll_count}次后未找到匹配卡片, 耗时{total_elapsed:.1f}s, title={title[:40]}")
    if note_id_hint:
        try:
            direct_url = f"https://www.xiaohongshu.com/explore/{note_id_hint}"
            log(f"  - [定位流程] 回退：尝试直接打开笔记ID: {note_id_hint}", level="debug")
            goto_with_retry(page, direct_url, timeout_ms=min(30000, goto_timeout_ms), retries=0)
            resolved_url = wait_for_current_note_url(
                page,
                expected_note_id=note_id_hint,
                timeout_ms=detail_wait_ms,
                require_xsec_token=require_xsec_token,
            )
            if resolved_url:
                log(f"  - [定位流程] 回退成功(直达笔记): {resolved_url}")
                return resolved_url
        except Exception as exc:
            log(f"  - [定位流程] 回退失败(直达笔记): {exc}", level="debug")
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
        log(f"导出失败数据出错: {exc}")
        return ""


def update_note_urls(args):
    log_path = getattr(args, 'log_file', '') or None
    setup_logger(log_path)
    
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
        target_date = parse_target_date(args.target_date)
        if not target_date:
            log(f"警告: 无法解析目标日期 '{args.target_date}'")
    
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows_data = list(reader)
        headers = {h: idx for idx, h in enumerate(reader.fieldnames or [])}

    require_headers(headers, [TITLE_HEADER, HOMEPAGE_HEADER, NOTE_URL_HEADER])

    total_rows = len(rows_data)
    all_row_indices = parse_rows_spec(args.rows, 2, total_rows + 1)  # CSV 数据行号从2开始（1是表头）
    if args.limit:
        all_row_indices = all_row_indices[: args.limit]
    if not all_row_indices:
        log("没有需要处理的行。")
        return output_path

    row_indices = all_row_indices
    if args.parallel_index >= 0 and args.parallel_total > 1:
        chunk_size = (len(all_row_indices) + args.parallel_total - 1) // args.parallel_total
        start_idx = args.parallel_index * chunk_size
        end_idx = min(start_idx + chunk_size, len(all_row_indices))
        row_indices = all_row_indices[start_idx:end_idx]
        actual_start_row = row_indices[0] if row_indices else 0
        actual_end_row = row_indices[-1] if row_indices else 0
        log(f"并行模式: 任务 {args.parallel_index + 1}/{args.parallel_total}，处理行 {actual_start_row}-{actual_end_row}")

    items = []
    for row_idx in row_indices:
        row_data = rows_data[row_idx - 2]
        items.append({
            "row": row_idx,
            "title": str(row_data.get(TITLE_HEADER, "") or "").strip(),
            "homepage_url": str(row_data.get(HOMEPAGE_HEADER, "") or "").strip(),
            "old_url": str(row_data.get(NOTE_URL_HEADER, "") or "").strip(),
        })
    apply_speed_profile_for_large_batch(
        len(items),
        session_mode=getattr(args, "session_mode", "rotate"),
        speed_mode=getattr(args, "speed_mode", "auto"),
    )

    if os.path.abspath(csv_path) == os.path.abspath(output_path):
        backup_path = make_backup(csv_path)
        if backup_path:
            log(f"已备份原文件: {backup_path}")

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
    
    log(f"CSV: {csv_path}")
    log(f"输出: {output_path}")
    log(f"处理行数: {len(items)}")
    log(f"使用 {len(session_dirs)} 个账号轮询:")
    for i, sd in enumerate(session_dirs):
        log(f"  账号{i+1}: {sd}")

    updated = 0
    skipped = 0
    failed = 0
    failed_items = []
    reset_pace_state()
    detail_wait_ms = get_effective_detail_wait_ms(args)
    max_scrolls = get_effective_max_scrolls(args)
    per_item_timeout_sec = get_effective_per_item_timeout(args)
    batch_size, batch_interval = get_effective_batch_pause(args)
    save_every = get_effective_save_every(args)

    session_mode = getattr(args, 'session_mode', 'rotate')
    multi_session_fixed_rotate = session_mode == "rotate" and len(session_dirs) > 1
    session_batch_min, session_batch_max, session_switch_rest_min, session_switch_rest_max = get_session_rotation_settings(args)
    if multi_session_fixed_rotate:
        batch_size = 0
        batch_interval = 0

    with sync_playwright() as p:
        contexts = []
        pages = []
        current_session_idx = 0
        current_session_remaining = 0
        
        if session_mode == "bind":
            log(f"\n[绑定模式] 同时打开 {len(session_dirs)} 个浏览器窗口...")
        else:
            log(f"\n[轮换模式] 同时打开 {len(session_dirs)} 个浏览器窗口...")
        
        for i, session_dir in enumerate(session_dirs):
            log(f"  - 启动账号{i+1}的浏览器...")
            viewport = build_human_viewport()
            context = p.chromium.launch_persistent_context(
                user_data_dir=session_dir,
                headless=args.headless,
                slow_mo=get_effective_slow_mo(args),
                viewport=viewport,
            )
            page = context.pages[0] if context.pages else context.new_page()
            inject_stealth_scripts(context, page)
            log(f"  - [反检测] 浏览器窗口尺寸: {viewport['width']}x{viewport['height']}", level="debug")
            
            if args.login_wait:
                first_homepage_url = ""
                for item in items:
                    first_homepage_url = normalize_homepage_url(item.get("homepage_url"))
                    if first_homepage_url:
                        break
                login_url = first_homepage_url or "https://www.xiaohongshu.com/explore"
                page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2000)
                warm_up_page(page, mode="full")
                check_verification_popup(page)
            
            contexts.append(context)
            pages.append(page)
        
        log(f"所有 {len(session_dirs)} 个浏览器窗口已就绪\n")
        
        def get_current_page():
            return pages[current_session_idx % len(pages)]

        def rotate_session_if_needed():
            nonlocal current_session_idx, current_session_remaining
            if session_mode != "rotate" or len(session_dirs) <= 1:
                return
            if current_session_remaining > 0:
                return
            previous_idx = current_session_idx
            current_session_idx = (current_session_idx + 1) % len(session_dirs)
            rest_sec = FIXED_SESSION_SWITCH_REST_SEC if multi_session_fixed_rotate else draw_session_switch_rest(session_switch_rest_min, session_switch_rest_max)
            log(f"  - [切换账号] 账号{previous_idx + 1} 本轮结束，休息 {rest_sec} 秒后切换到账号{current_session_idx + 1}")
            time.sleep(rest_sec)
            current_session_remaining = FIXED_SESSION_BATCH_SIZE if multi_session_fixed_rotate else draw_session_batch_size(session_batch_min, session_batch_max)
            log(f"  - [账号轮换] 账号{current_session_idx + 1} 本轮计划处理 {current_session_remaining} 条")

        perf_total_sec = 0.0
        perf_count = 0

        def finish_session_item(item_start_time=None, item_index=None):
            nonlocal current_session_remaining, perf_total_sec, perf_count
            if session_mode == "rotate" and len(session_dirs) > 1 and current_session_remaining > 0:
                current_session_remaining -= 1
            if item_start_time and item_index:
                perf_count += 1
                spent = time.time() - item_start_time
                perf_total_sec += spent
                if perf_count == 1 or perf_count % 20 == 0:
                    avg = perf_total_sec / max(1, perf_count)
                    remaining = max(0, len(items) - int(item_index))
                    eta = avg * remaining
                    log(f"  - [效率] 本条={spent:.1f}s 平均={avg:.1f}s 预计剩余={format_duration(eta)}")

        if session_mode == "rotate" and len(session_dirs) > 1:
            current_session_remaining = FIXED_SESSION_BATCH_SIZE if multi_session_fixed_rotate else draw_session_batch_size(session_batch_min, session_batch_max)
            if multi_session_fixed_rotate:
                log(
                    f"[固定轮换] 多账号单任务顺序执行：每个账号固定处理 {FIXED_SESSION_BATCH_SIZE} 条，"
                    f"切换前休息 {FIXED_SESSION_SWITCH_REST_SEC} 秒"
                )
            else:
                log(
                    f"[最稳轮换] 多账号按批次切换：每个账号连续处理 {session_batch_min}-{session_batch_max} 条，"
                    f"切换前休息 {session_switch_rest_min}-{session_switch_rest_max} 秒"
                )
            log(f"  - [账号轮换] 账号1 本轮计划处理 {current_session_remaining} 条")

        try:
            for index, item in enumerate(items, start=1):
                item_start_time = time.time()
                row_idx = item["row"]
                title = item["title"]
                homepage_url = item["homepage_url"]
                old_url = item["old_url"]
                old_note_id = extract_note_id(old_url)

                rotate_session_if_needed()

                page = get_current_page()
                log(f"\n[{index}/{len(items)}] 第 {row_idx} 行: {title}")
                log(f"  - 使用账号{current_session_idx + 1}")
                
                if index > 1:
                    apply_between_item_pacing()
                
                if batch_size > 0 and index > 1 and (index - 1) % batch_size == 0:
                    log(f"  - [批次休息] 已处理 {index - 1} 条，休息 {batch_interval} 秒避免风控...")
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
                    log("  - 跳过：笔记无标题，无法匹配")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue
                
                if not title or not homepage_url:
                    skipped += 1
                    log("  - 跳过：标题或主页链接为空")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue
                if args.only_empty and not is_blank(old_url):
                    skipped += 1
                    log("  - 跳过：笔记官方地址已有内容")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue

                try:
                    note_url = locate_note_url(
                        page,
                        homepage_url,
                        title,
                        max_scrolls=max_scrolls,
                        scroll_pixels=args.scroll_pixels,
                        detail_wait_ms=detail_wait_ms,
                        require_xsec_token=not args.allow_basic_url,
                        per_item_timeout_sec=per_item_timeout_sec,
                        target_date=target_date,
                        note_id_hint=old_note_id,
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
                    log(f"  - 查找超时: {exc}")
                    register_processing_outcome(False, row_idx=row_idx, reason="查找超时")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 查找失败: {exc}")
                    register_processing_outcome(False, row_idx=row_idx, reason="查找失败")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}), 保持原值")
                    log(f"  - [结果汇总] 行{row_idx} 失败(未找到匹配) | 累计: 更新={updated} 失败={failed} 跳过={skipped}", level="debug")
                    register_processing_outcome(False, row_idx=row_idx, reason="未找到匹配笔记")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 找到的地址没有笔记ID，保持原值: {note_url}")
                    register_processing_outcome(False, row_idx=row_idx, reason="地址没有笔记ID")
                    finish_session_item(item_start_time, index)
                    continue

                rows_data[row_idx - 2][NOTE_URL_HEADER] = note_url
                updated += 1
                id_note = ""
                if old_note_id and new_note_id and old_note_id != new_note_id:
                    id_note = f"（注意：笔记ID从 {old_note_id} 变为 {new_note_id}）"
                log(f"  - 已更新: {note_url} {id_note}")
                log(f"  - [结果汇总] 行{row_idx} 成功 | 累计: 更新={updated} 失败={failed} 跳过={skipped}", level="debug")
                register_processing_outcome(True, row_idx=row_idx)
                finish_session_item(item_start_time, index)

                if save_every and updated % save_every == 0:
                    save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
                    log(f"  - 已保存进度: {output_path}")
        except KeyboardInterrupt:
            log("\n\n用户中断，正在保存已处理数据...")
            save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
            log(f"已保存 {updated} 条更新到: {output_path}")
            raise
        finally:
            for ctx in contexts:
                try:
                    ctx.close()
                except Exception:
                    pass

    save_csv(rows_data, output_path, list(headers.keys()), original_csv_path=csv_path)
    log("\n处理完成。")
    log(f"更新: {updated}，跳过: {skipped}，失败: {failed}")
    log(f"结果文件: {output_path}")

    if failed_items:
        log("\n" + "=" * 60)
        log("失败数据明细：")
        log("=" * 60)
        for item in failed_items:
            log(f"\n行号: {item['row']}")
            log(f"  标题: {item['title']}")
            log(f"  主页: {item['homepage_url']}")
            log(f"  原地址: {item['old_url']}")
            log(f"  失败原因: {item['reason']}")

        failed_path = export_failed_items(failed_items, output_path)
        if failed_path:
            log(f"\n失败数据已导出到: {failed_path}")

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
        target_date = parse_target_date(args.target_date)
        if not target_date:
            log(f"警告: 无法解析目标日期 '{args.target_date}'")
    
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
        log("没有需要处理的行。")
        return output_path

    if args.parallel_index >= 0 and args.parallel_total > 1:
        chunk_size = (len(rows) + args.parallel_total - 1) // args.parallel_total
        start_idx = args.parallel_index * chunk_size
        end_idx = min(start_idx + chunk_size, len(rows))
        rows = rows[start_idx:end_idx]
        log(f"并行模式: 任务 {args.parallel_index + 1}/{args.parallel_total}，处理行 {start_idx + 1}-{end_idx}")

    items = build_rows(ws, headers, rows)
    url_col = headers[NOTE_URL_HEADER]
    apply_speed_profile_for_large_batch(
        len(items),
        session_mode=getattr(args, "session_mode", "rotate"),
        speed_mode=getattr(args, "speed_mode", "auto"),
    )

    if os.path.abspath(excel_path) == os.path.abspath(output_path):
        backup_path = make_backup(excel_path)
        if backup_path:
            log(f"已备份原文件: {backup_path}")

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
    
    log(f"Excel: {excel_path}")
    log(f"Sheet: {sheet_name}")
    log(f"输出: {output_path}")
    log(f"处理行数: {len(items)}")
    log(f"使用 {len(session_dirs)} 个账号轮询:")
    for i, sd in enumerate(session_dirs):
        log(f"  账号{i+1}: {sd}")

    updated = 0
    skipped = 0
    failed = 0
    actual_output_path = output_path
    failed_items = []
    reset_pace_state()
    detail_wait_ms = get_effective_detail_wait_ms(args)
    max_scrolls = get_effective_max_scrolls(args)
    per_item_timeout_sec = get_effective_per_item_timeout(args)
    batch_size, batch_interval = get_effective_batch_pause(args)
    save_every = get_effective_save_every(args)
    session_mode = getattr(args, 'session_mode', 'rotate')
    multi_session_fixed_rotate = session_mode == "rotate" and len(session_dirs) > 1
    session_batch_min, session_batch_max, session_switch_rest_min, session_switch_rest_max = get_session_rotation_settings(args)
    if multi_session_fixed_rotate:
        batch_size = 0
        batch_interval = 0

    with sync_playwright() as p:
        contexts = []
        pages = []
        current_session_idx = 0
        current_session_remaining = 0
        
        if session_mode == "bind":
            log(f"\n[绑定模式] 同时打开 {len(session_dirs)} 个浏览器窗口...")
        else:
            log(f"\n[轮换模式] 同时打开 {len(session_dirs)} 个浏览器窗口...")
        for i, session_dir in enumerate(session_dirs):
            log(f"  - 启动账号{i+1}的浏览器...")
            viewport = build_human_viewport()
            context = p.chromium.launch_persistent_context(
                user_data_dir=session_dir,
                headless=args.headless,
                slow_mo=get_effective_slow_mo(args),
                viewport=viewport,
            )
            page = context.pages[0] if context.pages else context.new_page()
            inject_stealth_scripts(context, page)
            log(f"  - [反检测] 浏览器窗口尺寸: {viewport['width']}x{viewport['height']}", level="debug")
            
            if args.login_wait:
                first_homepage_url = ""
                for item in items:
                    first_homepage_url = normalize_homepage_url(item.get("homepage_url"))
                    if first_homepage_url:
                        break
                login_url = first_homepage_url or "https://www.xiaohongshu.com/explore"
                page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2000)
                warm_up_page(page, mode="full")
                check_verification_popup(page)
            
            contexts.append(context)
            pages.append(page)
        
        log(f"所有 {len(session_dirs)} 个浏览器窗口已就绪\n")
        
        def get_current_page():
            return pages[current_session_idx % len(pages)]

        def rotate_session_if_needed():
            nonlocal current_session_idx, current_session_remaining
            if session_mode != "rotate" or len(session_dirs) <= 1:
                return
            if current_session_remaining > 0:
                return
            previous_idx = current_session_idx
            current_session_idx = (current_session_idx + 1) % len(session_dirs)
            rest_sec = FIXED_SESSION_SWITCH_REST_SEC if multi_session_fixed_rotate else draw_session_switch_rest(session_switch_rest_min, session_switch_rest_max)
            log(f"  - [切换账号] 账号{previous_idx + 1} 本轮结束，休息 {rest_sec} 秒后切换到账号{current_session_idx + 1}")
            time.sleep(rest_sec)
            current_session_remaining = FIXED_SESSION_BATCH_SIZE if multi_session_fixed_rotate else draw_session_batch_size(session_batch_min, session_batch_max)
            log(f"  - [账号轮换] 账号{current_session_idx + 1} 本轮计划处理 {current_session_remaining} 条")

        perf_total_sec = 0.0
        perf_count = 0

        def finish_session_item(item_start_time=None, item_index=None):
            nonlocal current_session_remaining, perf_total_sec, perf_count
            if session_mode == "rotate" and len(session_dirs) > 1 and current_session_remaining > 0:
                current_session_remaining -= 1
            if item_start_time and item_index:
                perf_count += 1
                spent = time.time() - item_start_time
                perf_total_sec += spent
                if perf_count == 1 or perf_count % 20 == 0:
                    avg = perf_total_sec / max(1, perf_count)
                    remaining = max(0, len(items) - int(item_index))
                    eta = avg * remaining
                    log(f"  - [效率] 本条={spent:.1f}s 平均={avg:.1f}s 预计剩余={format_duration(eta)}")

        if session_mode == "rotate" and len(session_dirs) > 1:
            current_session_remaining = FIXED_SESSION_BATCH_SIZE if multi_session_fixed_rotate else draw_session_batch_size(session_batch_min, session_batch_max)
            if multi_session_fixed_rotate:
                log(
                    f"[固定轮换] 多账号单任务顺序执行：每个账号固定处理 {FIXED_SESSION_BATCH_SIZE} 条，"
                    f"切换前休息 {FIXED_SESSION_SWITCH_REST_SEC} 秒"
                )
            else:
                log(
                    f"[最稳轮换] 多账号按批次切换：每个账号连续处理 {session_batch_min}-{session_batch_max} 条，"
                    f"切换前休息 {session_switch_rest_min}-{session_switch_rest_max} 秒"
                )
            log(f"  - [账号轮换] 账号1 本轮计划处理 {current_session_remaining} 条")

        try:
            for index, item in enumerate(items, start=1):
                item_start_time = time.time()
                row_idx = item["row"]
                title = item["title"]
                homepage_url = item["homepage_url"]
                old_url = item["old_url"]
                old_note_id = extract_note_id(old_url)

                rotate_session_if_needed()

                page = get_current_page()
                log(f"\n[{index}/{len(items)}] 第 {row_idx} 行: {title}")
                log(f"  - 使用账号{current_session_idx + 1}")
                
                if index > 1:
                    apply_between_item_pacing()
                
                if batch_size > 0 and index > 1 and (index - 1) % batch_size == 0:
                    log(f"  - [批次休息] 已处理 {index - 1} 条，休息 {batch_interval} 秒避免风控...")
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
                    log("  - 跳过：笔记无标题，无法匹配")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue
                
                if not title or not homepage_url:
                    skipped += 1
                    log("  - 跳过：标题或主页链接为空")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue
                if args.only_empty and not is_blank(old_url):
                    skipped += 1
                    log("  - 跳过：笔记官方地址已有内容")
                    register_processing_outcome(True, row_idx=row_idx)
                    finish_session_item(item_start_time, index)
                    continue

                try:
                    note_url = locate_note_url(
                        page,
                        homepage_url,
                        title,
                        max_scrolls=max_scrolls,
                        scroll_pixels=args.scroll_pixels,
                        detail_wait_ms=detail_wait_ms,
                        require_xsec_token=not args.allow_basic_url,
                        per_item_timeout_sec=per_item_timeout_sec,
                        target_date=target_date,
                        note_id_hint=old_note_id,
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
                    log(f"  - 查找超时: {exc}")
                    register_processing_outcome(False, row_idx=row_idx, reason="查找超时")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 查找失败: {exc}")
                    register_processing_outcome(False, row_idx=row_idx, reason="查找失败")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}), 保持原值")
                    log(f"  - [结果汇总] 行{row_idx} 失败(未找到匹配) | 累计: 更新={updated} 失败={failed} 跳过={skipped}", level="debug")
                    register_processing_outcome(False, row_idx=row_idx, reason="未找到匹配笔记")
                    finish_session_item(item_start_time, index)
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
                    log(f"  - 找到的地址没有笔记ID，保持原值: {note_url}")
                    register_processing_outcome(False, row_idx=row_idx, reason="地址没有笔记ID")
                    finish_session_item(item_start_time, index)
                    continue

                ws.cell(row=row_idx, column=url_col).value = note_url
                updated += 1
                id_note = ""
                if old_note_id and new_note_id and old_note_id != new_note_id:
                    id_note = f"（注意：笔记ID从 {old_note_id} 变为 {new_note_id}）"
                log(f"  - 已更新: {note_url} {id_note}")
                log(f"  - [结果汇总] 行{row_idx} 成功 | 累计: 更新={updated} 失败={failed} 跳过={skipped}", level="debug")
                register_processing_outcome(True, row_idx=row_idx)
                finish_session_item(item_start_time, index)

                if save_every and updated % save_every == 0:
                    actual_output_path = save_workbook(wb, actual_output_path)
                    log(f"  - 已保存进度: {actual_output_path}")
        except KeyboardInterrupt:
            log("\n\n用户中断，正在保存已处理数据...")
            actual_output_path = save_workbook(wb, actual_output_path)
            log(f"已保存 {updated} 条更新到: {actual_output_path}")
            raise
        finally:
            for ctx in contexts:
                try:
                    ctx.close()
                except Exception:
                    pass
            if updated > 0 or skipped > 0 or failed > 0:
                try:
                    actual_output_path = save_workbook(wb, actual_output_path)
                    log(f"  - [停止保护] 强制退出前已保存 {updated} 条更新到: {actual_output_path}")
                except Exception:
                    pass

    if updated > 0 or skipped > 0 or failed > 0:
        actual_output_path = save_workbook(wb, actual_output_path)
    log("\n处理完成。")
    log(f"更新: {updated}，跳过: {skipped}，失败: {failed}")
    log(f"结果文件: {actual_output_path}")

    if failed_items:
        log("\n" + "=" * 60)
        log("失败数据明细：")
        log("=" * 60)
        for item in failed_items:
            log(f"\n行号: {item['row']}")
            log(f"  标题: {item['title']}")
            log(f"  主页: {item['homepage_url']}")
            log(f"  原地址: {item['old_url']}")
            log(f"  失败原因: {item['reason']}")

        failed_path = export_failed_items(failed_items, actual_output_path)
        if failed_path:
            log(f"\n失败数据已导出到: {failed_path}")

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
    parser.add_argument("--max-scrolls", type=int, default=12, help="每个主页最多向下滚动次数，默认 12")
    parser.add_argument("--scroll-pixels", type=int, default=900, help="每次滚动像素")
    parser.add_argument("--detail-wait-ms", type=int, default=8000, help="点击笔记后等待地址栏变化的毫秒数")
    parser.add_argument("--per-item-timeout", type=int, default=15, help="每条数据的最大处理时间（秒），超时后跳过")
    parser.add_argument("--save-every", type=int, default=20, help="每更新多少行保存一次，0 表示只最后保存；turbo 默认会自动关闭过程保存")
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
    parser.add_argument("--batch-size", type=int, default=100, help="每处理多少条后休息一次，0表示不休息，默认100")
    parser.add_argument("--batch-interval", type=int, default=120, help="批次休息秒数，默认120")
    parser.add_argument("--sessions", default="", help="多个session目录，用逗号分隔，多账号场景下推荐按批次轮换")
    parser.add_argument("--session-mode", choices=["rotate", "bind"], default="rotate", help="账号模式: rotate=账号按批次轮换(推荐), bind=每个任务使用分配到的账号组")
    parser.add_argument("--speed-mode", choices=["auto", "default", "turbo"], default="auto", help="速度模式: auto=大批量自动提速(默认), default=不提速, turbo=优先提速(轻风控)")
    parser.add_argument("--session-batch-min", type=int, default=ANTI_DETECTION_CONFIG["session_batch_min"], help="轮换模式下单账号最少连续处理多少条，默认10")
    parser.add_argument("--session-batch-max", type=int, default=ANTI_DETECTION_CONFIG["session_batch_max"], help="轮换模式下单账号最多连续处理多少条，默认15")
    parser.add_argument("--session-switch-rest-min", type=int, default=ANTI_DETECTION_CONFIG["session_switch_rest_min_sec"], help="轮换模式下切账号前最少休息秒数，默认60")
    parser.add_argument("--session-switch-rest-max", type=int, default=ANTI_DETECTION_CONFIG["session_switch_rest_max_sec"], help="轮换模式下切账号前最多休息秒数，默认180")
    parser.add_argument("--log-file", default="", help="日志文件路径，默认自动生成到 logs/ 目录")
    return parser


def main():
    parser = build_arg_parser()
    args = parser.parse_args()
    try:
        update_note_urls(args)
    except KeyboardInterrupt:
        log("\n用户中断。")
        raise SystemExit(130)
    except Exception as exc:
        log(f"错误: {exc}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
