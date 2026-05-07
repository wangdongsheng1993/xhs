import pandas as pd
import re
import time
import os
import sys
import statistics
import shutil
import urllib.request
from urllib.parse import parse_qs, unquote, urlsplit, urlunsplit
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, coordinate_to_tuple
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
from playwright.sync_api import sync_playwright


def get_env_int(name, default):
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(float(value))
    except Exception:
        return default


def get_env_float(name, default):
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return float(value)
    except Exception:
        return default


def get_env_bool(name, default):
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "y", "on"}


# --- 配置区 ---
EXCEL_PATH = os.getenv(
    "XHS_EXCEL_PATH", r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx"
).strip()
OUTPUT_PATH = os.getenv(
    "XHS_OUTPUT_PATH",
    r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx",
).strip()
SHEET_NAME = os.getenv("XHS_SHEET_NAME", "蒸烤KOL").strip()
# 浏览器数据目录，用于保存登录状态
USER_DATA_DIR = os.path.join(os.getcwd(), "browser_session")
# 厨房场景图保存目录
SCREENSHOT_DIR = os.path.join(os.getcwd(), "steam_kol_screenshots")
# 每个页面打开后等待的秒数
PAGE_WAIT_SECONDS = get_env_float("XHS_PAGE_WAIT_SECONDS", 2.0)
TAB_WAIT_MS = max(0, get_env_int("XHS_TAB_WAIT_MS", 600))
SAVE_EVERY_ROWS = max(1, get_env_int("XHS_SAVE_EVERY_ROWS", 5))
SLOW_MO_MS = max(0, get_env_int("XHS_SLOW_MO_MS", 0))
PAGE_READY_TIMEOUT_MS = max(500, get_env_int("XHS_PAGE_READY_TIMEOUT_MS", 3000))
HOMEPAGE_POPUP_WAIT_MS = max(300, get_env_int("XHS_HOMEPAGE_POPUP_WAIT_MS", 1200))
FORCE_REWRITE = get_env_bool("XHS_FORCE_REWRITE", False)
SKIP_COMPLETED_ROWS = get_env_bool("XHS_SKIP_COMPLETED_ROWS", True)
DEBUG_TARGET_NAME = os.getenv("XHS_DEBUG_NAME", "").strip()
DEBUG_TARGET_ROW = os.getenv("XHS_DEBUG_ROW", "").strip()
DEBUG_MAX_ROWS = os.getenv("XHS_DEBUG_MAX_ROWS", "").strip()
DEBUG_VERBOSE = os.getenv("XHS_DEBUG_VERBOSE", "0").strip() == "1"
LOGIN_WAIT_SECONDS = int(os.getenv("XHS_LOGIN_WAIT_SECONDS", "10").strip() or "10")
REQUIRE_ENTER_CONFIRM = os.getenv("XHS_REQUIRE_ENTER_CONFIRM", "0").strip() == "1"

KITCHEN_KEYWORDS = [
    "厨房",
    "厨电",
    "橱柜",
    "厨柜",
    "餐边柜",
    "冰箱",
    "蒸烤箱",
    "蒸箱",
    "烤箱",
    "洗碗机",
    "油烟机",
    "烟机",
    "燃气灶",
    "灶台",
    "集成灶",
    "水槽",
    "龙头",
]

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ============== 通用工具函数 ==============


def sanitize_filename(text):
    text = re.sub(r'[\\/:*?"<>|]+', "_", str(text))
    return text[:80] or "unknown"


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def cleanup_dir(path):
    if not path:
        return
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def header_map(ws):
    return {
        ws.cell(row=1, column=col).value: col for col in range(1, ws.max_column + 1)
    }


def is_blank(value):
    if value is None:
        return True
    try:
        if value != value:
            return True
    except Exception:
        pass
    if isinstance(value, str):
        return not value.strip()
    return False


def has_all_values(mapping, fields):
    for field in fields:
        if is_blank(mapping.get(field)):
            return False
    return True


def set_cell(ws, headers, row_idx, header, value):
    col = headers.get(header)
    if not col:
        return
    cell = ws.cell(row=row_idx, column=col)
    if not FORCE_REWRITE and not is_blank(cell.value):
        return
    cell.value = value


def get_cell_ref(headers, row_idx, header):
    col = headers.get(header)
    if not col:
        return None
    return f"{get_column_letter(col)}{row_idx}"


def prepare_sheet_layout(ws, headers):
    image_headers = ["厨房场景图"]
    for header in image_headers:
        col = headers.get(header)
        if col:
            ws.column_dimensions[get_column_letter(col)].width = 26


def queue_image_for_cell(
    ws, headers, image_jobs, row_idx, header, image_path, width, height
):
    cell_ref = get_cell_ref(headers, row_idx, header)
    if not cell_ref or not image_path or not os.path.exists(image_path):
        return
    if not FORCE_REWRITE and not is_blank(ws[cell_ref].value):
        return
    ws.row_dimensions[row_idx].height = max(
        ws.row_dimensions[row_idx].height or 15, height * 0.75
    )
    ws[cell_ref].value = None
    image_jobs.append(
        {
            "cell_ref": cell_ref,
            "image_path": image_path,
            "width": width,
            "height": height,
        }
    )


def apply_image_jobs(ws, image_jobs):
    ws._images = []
    for job in image_jobs:
        if not os.path.exists(job["image_path"]):
            continue
        try:
            img = XLImage(job["image_path"])
            img.width = job["width"]
            img.height = job["height"]
            ws.add_image(img, job["cell_ref"])
        except Exception:
            pass


def save_progress(wb, output_path, processed_count):
    image_jobs = getattr(wb, "_codex_image_jobs", [])
    for worksheet in wb.worksheets:
        apply_image_jobs(worksheet, image_jobs.get(worksheet.title, []))

    saved_path = output_path
    try:
        wb.save(output_path)
    except PermissionError:
        saved_path = OUTPUT_PATH.replace(".xlsx", "_自动保存.xlsx")
        wb.save(saved_path)
        print(f"  - 结果文件被占用，已改存到: {saved_path}")
    else:
        print(f"  - 已保存进度到: {saved_path}")

    if DEBUG_VERBOSE:
        print(f"  - 嵌入图片 {sum(len(jobs) for jobs in image_jobs.values())} 张")
    return saved_path


# ============== 复用自 xhs_extractor 的函数 ==============


def wait_for_profile_page_ready(page, timeout_ms=PAGE_READY_TIMEOUT_MS):
    selectors = [
        ".blogger-data__item",
        ".blogger-data__label",
        ".blogger-data__value",
        ".price-box",
    ]
    deadline = time.time() + max(timeout_ms, 0) / 1000
    while time.time() < deadline:
        for selector in selectors:
            try:
                locator = page.locator(selector).first
                if locator.count() and locator.is_visible(timeout=200):
                    return True
            except Exception:
                pass
        time.sleep(0.2)
    if PAGE_WAIT_SECONDS > 0:
        page.wait_for_timeout(int(PAGE_WAIT_SECONDS * 1000))
    return False


def parse_w_value(text):
    if not text or text == "-":
        return 0.0
    text = (
        text.replace("w", "")
        .replace("万", "")
        .replace("+", "")
        .replace(",", "")
        .strip()
    )
    try:
        if "k" in text.lower():
            return float(text.lower().replace("k", "")) / 10
        return float(text)
    except Exception:
        return 0.0


def parse_page_w_value(text):
    raw_text = str(text or "").strip().lower()
    if not raw_text:
        return 0.0
    if any(unit in raw_text for unit in ["w", "万", "k"]):
        return parse_w_value(raw_text)
    clean_text = (
        raw_text.replace(",", "")
        .replace("+", "")
        .replace("粉丝", "")
        .replace("获赞与收藏", "")
        .strip()
    )
    try:
        num = float(clean_text)
    except Exception:
        return 0.0
    if "." in clean_text and num < 100:
        return num
    return round(num / 10000, 4)


def get_level(fans_w):
    if fans_w < 1:
        return "KOC"
    if 1 <= fans_w < 10:
        return "尾部"
    if 10 <= fans_w < 30:
        return "腰部"
    if 30 <= fans_w < 50:
        return "肩部"
    if fans_w >= 50:
        return "头部"
    return "未知"


def get_first_visible(locator, timeout=3000):
    try:
        count = locator.count()
        for i in range(count):
            item = locator.nth(i)
            try:
                if item.is_visible(timeout=timeout):
                    return item
            except Exception:
                pass
    except Exception:
        pass
    return None


def wait_for_login_confirmation():
    prompt = "\n>>> 登录完成后，按 Enter 键开始采集..."
    if REQUIRE_ENTER_CONFIRM:
        try:
            if sys.stdin and sys.stdin.isatty():
                input(prompt)
                return
        except EOFError:
            pass
    print(prompt)
    print(f"等待 {LOGIN_WAIT_SECONDS} 秒后自动继续，并复用已有登录态...")
    time.sleep(LOGIN_WAIT_SECONDS)


def extract_user_id_from_pgy_url(url):
    match = re.search(r"/blogger-detail/([^/?]+)", str(url))
    return match.group(1) if match else None


def fetch_json(page, url):
    if DEBUG_VERBOSE:
        print(f"    [DEBUG] 请求接口: {url}")
    result = page.evaluate(
        """
        async (requestUrl) => {
            const response = await fetch(requestUrl, {
                method: 'GET',
                credentials: 'include'
            });
            const text = await response.text();
            return { ok: response.ok, status: response.status, text };
        }
        """,
        url,
    )
    if not result.get("ok"):
        raise Exception(f"接口请求失败: {url} (HTTP {result.get('status')})")
    try:
        payload = page.evaluate("text => JSON.parse(text)", result.get("text", ""))
    except Exception as e:
        raise Exception(f"接口返回非 JSON: {url} ({e})")
    if isinstance(payload, dict):
        if payload.get("code") not in (None, 0):
            raise Exception(
                f"接口返回异常: {url} code={payload.get('code')} msg={payload.get('msg')}"
            )
        if payload.get("success") is False:
            raise Exception(f"接口返回失败: {url} msg={payload.get('msg')}")
        if "data" in payload:
            return payload.get("data")
    return payload


def get_cached_json(api_cache, url):
    payload = api_cache.get(url)
    if isinstance(payload, dict):
        if payload.get("code") not in (None, 0):
            raise Exception(
                f"接口返回异常: {url} code={payload.get('code')} msg={payload.get('msg')}"
            )
        if payload.get("success") is False:
            raise Exception(f"接口返回失败: {url} msg={payload.get('msg')}")
        if "data" in payload:
            return payload.get("data")
    return payload


def get_json_with_cache(page, api_cache, url):
    cached = api_cache.get(url)
    if cached is not None:
        if DEBUG_VERBOSE:
            print(f"    [DEBUG] 使用页面缓存响应: {url}")
        return get_cached_json(api_cache, url)
    return fetch_json(page, url)


def find_cached_json_by_keyword(api_cache, keyword):
    for url, payload in api_cache.items():
        if keyword in url:
            return get_cached_json({url: payload}, url)
    return None


def wait_for_cached_json_by_keyword(api_cache, keyword, timeout_ms=5000):
    deadline = time.time() + max(timeout_ms, 0) / 1000
    while time.time() < deadline:
        data = find_cached_json_by_keyword(api_cache, keyword)
        if data is not None:
            return data
        time.sleep(0.2)
    return None


def to_ratio_decimal(value):
    try:
        if value in (None, ""):
            return ""
        num = float(value)
        if num > 1:
            num /= 100
        return round(num, 6)
    except Exception:
        return ""


def find_percent(items, targets, name_keys=("group", "name", "desc")):
    if isinstance(targets, str):
        targets = [targets]
    target_set = {str(x).strip().lower() for x in targets}
    for item in items or []:
        for key in name_keys:
            value = item.get(key)
            if value is not None and str(value).strip().lower() in target_set:
                return item.get("percent")
    return None


def get_page_data_value(page, label):
    return (
        page.evaluate(
            """
        (label) => {
            const labels = Array.from(document.querySelectorAll('.blogger-data__label'));
            const labelEl = labels.find(el => (el.textContent || '').trim() === label);
            if (!labelEl) return '';
            const item = labelEl.closest('.blogger-data__item');
            const valueEl = item && item.querySelector('.blogger-data__value');
            return valueEl ? (valueEl.textContent || '').trim() : '';
        }
        """,
            label,
        )
        or ""
    )


def get_page_price_value(page, label):
    text = (
        page.evaluate(
            """
        (label) => {
            const labels = Array.from(document.querySelectorAll('.price-box span'));
            const labelEl = labels.find(el => (el.textContent || '').trim() === label);
            if (!labelEl) return '';
            const box = labelEl.closest('.price-box');
            const priceEl = box && box.querySelector('.active-price');
            return priceEl ? (priceEl.textContent || '').trim() : '';
        }
        """,
            label,
        )
        or ""
    )
    price_text = re.sub(r"[^\d.]", "", text)
    return float(price_text) if price_text else 0.0


def infer_cooperation_form_from_notes(notes_data, blogger_data):
    notes = (notes_data or {}).get("list") or []
    if notes:
        video_count = sum(1 for note in notes if note.get("isVideo"))
        image_count = sum(1 for note in notes if not note.get("isVideo"))
        if DEBUG_VERBOSE:
            print(f"    [DEBUG] 笔记案例统计: video={video_count}, image={image_count}")
        if video_count > image_count:
            return "视频"
        if image_count > video_count:
            return "图文"
    if blogger_data.get("videoState") == 1 and blogger_data.get("pictureState") != 1:
        return "视频"
    if blogger_data.get("pictureState") == 1 and blogger_data.get("videoState") != 1:
        return "图文"
    if blogger_data.get("videoState") == 1:
        return "视频"
    if blogger_data.get("pictureState") == 1:
        return "图文"
    return ""


def _to_number(value):
    try:
        if value in (None, ""):
            return 0
        return float(value)
    except Exception:
        return 0


def get_note_read_count(note):
    for key in (
        "readNum",
        "readCount",
        "readingNum",
        "read",
        "exposureNum",
        "impressionNum",
    ):
        value = _to_number((note or {}).get(key))
        if value > 0:
            return value
    return 0


def get_note_interaction_count(note):
    note = note or {}
    return (
        _to_number(note.get("likeNum"))
        + _to_number(note.get("collectNum"))
        + _to_number(note.get("commentNum"))
        + _to_number(note.get("shareNum"))
    )


def get_notes_median_fallback(notes_data, top_n=4):
    notes = (notes_data or {}).get("list") or []
    notes = notes[:top_n]
    if not notes:
        return 0, 0
    read_values = [get_note_read_count(note) for note in notes]
    interact_values = [get_note_interaction_count(note) for note in notes]
    read_values = [value for value in read_values if value > 0]
    interact_values = [value for value in interact_values if value > 0]
    read_median = statistics.median_low(read_values) if read_values else 0
    interact_median = statistics.median_low(interact_values) if interact_values else 0
    return read_median, interact_median


def is_xiaohongshu_host(parsed_url):
    host = (parsed_url.hostname or "").lower()
    return host == "xiaohongshu.com" or host.endswith(".xiaohongshu.com")


def normalize_xhs_homepage_url(raw_url):
    """把登录跳转、来源参数等还原成稳定的小红书主页地址。"""
    url = str(raw_url or "").strip()
    if not url or url == "about:blank":
        return ""

    if url.startswith("//"):
        url = f"https:{url}"
    elif url.startswith("/"):
        url = f"https://www.xiaohongshu.com{url}"

    for _ in range(3):
        parsed = urlsplit(url)
        if not (is_xiaohongshu_host(parsed) and parsed.path.rstrip("/") == "/login"):
            break

        redirect_values = parse_qs(parsed.query).get("redirectPath") or []
        if not redirect_values:
            break

        redirected_url = unquote(redirect_values[0] or "").strip()
        if not redirected_url or redirected_url == url:
            break

        if redirected_url.startswith("//"):
            redirected_url = f"https:{redirected_url}"
        elif redirected_url.startswith("/"):
            redirected_url = f"https://www.xiaohongshu.com{redirected_url}"

        url = redirected_url

    parsed = urlsplit(url)
    if is_xiaohongshu_host(parsed) and parsed.path.startswith("/user/profile/"):
        profile_id = parsed.path.split("/user/profile/", 1)[1].split("/", 1)[0]
        if profile_id:
            return urlunsplit(
                (
                    "https",
                    "www.xiaohongshu.com",
                    f"/user/profile/{profile_id}",
                    "",
                    "",
                )
            )

    return url


def get_homepage_url_from_profile_click(context, page, red_id):
    if not red_id:
        return ""
    candidates = [
        page.locator("a").filter(has_text=red_id).first,
        page.get_by_text(red_id, exact=True),
    ]
    for locator in candidates:
        try:
            if not locator.count():
                continue
            clickable = get_first_visible(locator)
            if not clickable:
                clickable = locator.first
            try:
                href = clickable.evaluate(
                    """
                    (el) => {
                        const anchor = el.closest('a') || el;
                        return anchor && anchor.href ? anchor.href : '';
                    }
                    """
                )
                if href and href != "about:blank":
                    return normalize_xhs_homepage_url(href) or href
            except Exception:
                pass
            with context.expect_page(timeout=3000) as popup_info:
                clickable.click()
            popup = popup_info.value
            popup.wait_for_load_state("domcontentloaded", timeout=10000)
            url = ""
            deadline = time.time() + HOMEPAGE_POPUP_WAIT_MS / 1000
            while time.time() < deadline:
                url = popup.url or ""
                if url and url != "about:blank":
                    break
                popup.wait_for_timeout(200)
            popup.close()
            if url and url != "about:blank":
                return normalize_xhs_homepage_url(url) or url
        except Exception:
            try:
                href = locator.first.evaluate(
                    """
                    (el) => {
                        const anchor = el.closest('a') || el;
                        return anchor && anchor.href ? anchor.href : '';
                    }
                    """
                )
                if href:
                    return normalize_xhs_homepage_url(href) or href
            except Exception:
                pass
    return ""


def get_kol_type_from_page(page):
    """从页面 .blogger-tag-list 提取第一个标签作为 KOL 类型"""
    return (
        page.evaluate(
            """
            () => {
                const tagList = document.querySelector('.blogger-tag-list');
                if (!tagList) return '';
                const firstTag = tagList.querySelector('span, li, a, div');
                return firstTag ? (firstTag.textContent || '').trim() : '';
            }
            """
        )
        or ""
    )


def goto_with_retry(page, url, wait_until="domcontentloaded", timeout=30000, retries=2):
    last_error = None
    for attempt in range(retries + 1):
        try:
            page.goto(url, wait_until=wait_until, timeout=timeout)
            return True
        except Exception as e:
            last_error = e
            message = str(e)
            if "interrupted by another navigation" not in message or attempt >= retries:
                raise
            if DEBUG_VERBOSE:
                print(f"    [DEBUG] 页面跳转被打断，准备重试: {url}")
            page.wait_for_timeout(1500)
    if last_error:
        raise last_error
    return False


def matches_debug_row(target_row_value, actual_row):
    if not target_row_value:
        return True
    for part in str(target_row_value).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start_text, end_text = part.split("-", 1)
                start = int(start_text.strip())
                end = int(end_text.strip())
                if end < start:
                    start, end = end, start
                if start <= actual_row <= end:
                    return True
            except Exception:
                continue
        else:
            try:
                if actual_row == int(part):
                    return True
            except Exception:
                continue
    return False


def should_process_row(row_idx, kol_name, processed_count):
    if DEBUG_TARGET_ROW:
        try:
            if not matches_debug_row(DEBUG_TARGET_ROW, row_idx):
                return False
        except Exception:
            pass
    if DEBUG_TARGET_NAME:
        if DEBUG_TARGET_NAME.lower() not in str(kol_name).lower():
            return False
    if DEBUG_MAX_ROWS:
        try:
            if processed_count >= int(DEBUG_MAX_ROWS):
                return False
        except Exception:
            pass
    return True


def build_notes_detail_keyword(user_id, page_number=1):
    return (
        "/api/solar/kol/data_v2/notes_detail"
        f"?advertiseSwitch=1&orderType=1&pageNumber={page_number}&pageSize=8"
        f"&userId={user_id}&noteType=4&isThirdPlatform=0"
    )


def open_brand_note_tabs(page):
    data_tab = get_first_visible(page.get_by_text("数据概览", exact=True))
    if data_tab:
        data_tab.click()
        if TAB_WAIT_MS:
            page.wait_for_timeout(TAB_WAIT_MS)
    coop_note_tab = get_first_visible(page.get_by_text("合作笔记", exact=True))
    if coop_note_tab:
        coop_note_tab.click()
        if TAB_WAIT_MS:
            page.wait_for_timeout(TAB_WAIT_MS)


def download_image(url, save_path):
    if not url:
        return False
    try:
        ensure_dir(os.path.dirname(save_path))
        urllib.request.urlretrieve(url, save_path)
        return True
    except Exception:
        return False


def pick_kitchen_note(notes):
    for note in (notes or [])[:16]:
        text = f"{note.get('title') or ''} {note.get('brandName') or ''}"
        if any(keyword in text for keyword in KITCHEN_KEYWORDS):
            return note
    return None


# ============== 蒸烤KOL 专用配置 ==============

STEAM_COMPLETED_FIELDS = [
    "主页链接",
    "厨房场景图",
    "KOL类型",
    "粉丝量（W）",
    "平台价格",
    "女粉占比",
    "18-24年龄占比",
    "25-34年龄占比",
    "35-44年龄占比",
    "近30天阅读中位数",
    "近30天互动中位",
]


def run_extraction():
    if not os.path.exists(EXCEL_PATH):
        print(f"错误: 找不到文件 {EXCEL_PATH}")
        return

    ensure_dir(SCREENSHOT_DIR)
    wb = load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    headers = header_map(ws)
    prepare_sheet_layout(ws, headers)
    wb._codex_image_jobs = {ws.title: []}

    pending_rows = []
    for row_idx in range(2, ws.max_row + 1):
        kol_name = (
            ws.cell(row=row_idx, column=headers["KOL名称"]).value
            if headers.get("KOL名称")
            else ""
        )
        pgy_url = (
            ws.cell(row=row_idx, column=headers["蒲公英链接"]).value
            if headers.get("蒲公英链接")
            else ""
        )
        if not pgy_url or not str(pgy_url).startswith("http"):
            continue
        if not should_process_row(row_idx, kol_name, len(pending_rows)):
            continue
        row_snapshot = {
            header: ws.cell(row=row_idx, column=col).value
            for header, col in headers.items()
            if header
        }
        if not FORCE_REWRITE and has_all_values(row_snapshot, STEAM_COMPLETED_FIELDS):
            continue
        pending_rows.append(row_idx)

    if not pending_rows:
        print("当前选中行均已有结果，跳过浏览器采集。")
        last_saved_path = save_progress(wb, OUTPUT_PATH, 0)
        cleanup_dir(SCREENSHOT_DIR)
        print(f"\n[完成] 所有任务处理完毕！结果已保存到:\n  {last_saved_path}")
        return

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR, headless=False, slow_mo=SLOW_MO_MS
        )
        page = context.new_page()
        api_cache = {}

        def on_response(response):
            try:
                url = response.url
                content_type = response.headers.get("content-type", "")
                if (
                    "application/json" not in content_type
                    and "text/json" not in content_type
                ):
                    return
                if "pgy.xiaohongshu.com/api/" not in url:
                    return
                api_cache[url] = response.json()
            except Exception:
                pass

        page.on("response", on_response)

        print("\n" + "=" * 50)
        print("【步骤 1: 请手动登录蒲公英】")
        print("浏览器已打开，请在浏览器中完成登录：")
        print("  > 网址: https://pgy.xiaohongshu.com")
        print("  > 选择【我是代理商】进行账号密码登录")
        print("  > 登录完成后，回到此窗口按 Enter 键继续")
        print("=" * 50)

        try:
            page.goto(
                "https://pgy.xiaohongshu.com",
                wait_until="domcontentloaded",
                timeout=15000,
            )
        except Exception:
            pass

        wait_for_login_confirmation()

        print("\n【步骤 2: 开始执行采集】")
        if DEBUG_TARGET_NAME or DEBUG_TARGET_ROW or DEBUG_MAX_ROWS:
            print("【调试模式】已启用单条/限量处理")
            if DEBUG_TARGET_NAME:
                print(f"  - 按名称筛选: {DEBUG_TARGET_NAME}")
            if DEBUG_TARGET_ROW:
                print(f"  - 指定 Excel 行号: {DEBUG_TARGET_ROW}")
            if DEBUG_MAX_ROWS:
                print(f"  - 最多处理条数: {DEBUG_MAX_ROWS}")

        processed_count = 0
        last_saved_path = OUTPUT_PATH
        fallback_rows = []

        for row_idx in pending_rows:
            kol_name = (
                ws.cell(row=row_idx, column=headers["KOL名称"]).value
                if headers.get("KOL名称")
                else f"Row {row_idx}"
            )
            print(f"\n>>> [{row_idx}/{ws.max_row}] 正在处理: {kol_name}")

            pgy_url = (
                ws.cell(row=row_idx, column=headers["蒲公英链接"]).value
                if headers.get("蒲公英链接")
                else ""
            )
            if not pgy_url or not str(pgy_url).startswith("http"):
                continue

            try:
                goto_with_retry(
                    page,
                    pgy_url,
                    wait_until="domcontentloaded",
                    timeout=30000,
                    retries=2,
                )
                print("  - 蒲公英页已打开，等待关键信息加载...")
                wait_for_profile_page_ready(page)

                user_id = extract_user_id_from_pgy_url(pgy_url)
                if not user_id:
                    raise Exception("无法从蒲公英链接中解析 userId")
                if DEBUG_VERBOSE:
                    print(f"    [DEBUG] userId: {user_id}")

                blogger_api = f"https://pgy.xiaohongshu.com/api/solar/cooperator/user/blogger/{user_id}"
                fans_profile_api = f"https://pgy.xiaohongshu.com/api/solar/kol/data/{user_id}/fans_profile"
                summary_api = f"https://pgy.xiaohongshu.com/api/pgy/kol/data/data_summary?userId={user_id}&business=1"
                notes_detail_keyword = build_notes_detail_keyword(user_id, 1)

                blogger_data = get_json_with_cache(page, api_cache, blogger_api) or {}
                fans_profile = (
                    get_json_with_cache(page, api_cache, fans_profile_api) or {}
                )
                notes_detail = (
                    find_cached_json_by_keyword(api_cache, notes_detail_keyword) or {}
                )
                if not notes_detail:
                    open_brand_note_tabs(page)
                    notes_detail = (
                        wait_for_cached_json_by_keyword(
                            api_cache, notes_detail_keyword, timeout_ms=6000
                        )
                        or {}
                    )

                try:
                    summary_data = (
                        get_json_with_cache(page, api_cache, summary_api) or {}
                    )
                except Exception as summary_err:
                    summary_data = {}
                    print(f"  - 数据概览接口获取失败，已跳过中位数: {summary_err}")

                # 小红书号
                red_id = blogger_data.get("redId", "")
                if red_id:
                    if headers.get("ID"):
                        set_cell(ws, headers, row_idx, "ID", red_id)
                    print(f"  - 抓取小红书号: {red_id}")

                    homepage_url = get_homepage_url_from_profile_click(
                        context, page, red_id
                    )
                    if homepage_url:
                        set_cell(ws, headers, row_idx, "主页链接", homepage_url)
                        print(f"  - 抓取主页链接: {homepage_url}")

                # 基础统计
                fans_text = get_page_data_value(page, "粉丝数")
                likes_text = get_page_data_value(page, "获赞与收藏")
                fans_w = parse_page_w_value(fans_text)
                if headers.get("粉丝量（W）"):
                    set_cell(ws, headers, row_idx, "粉丝量（W）", fans_w)
                if headers.get("赞藏量（W）"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "赞藏量（W）",
                        parse_page_w_value(likes_text),
                    )

                # KOL类型
                kol_type = get_kol_type_from_page(page)
                if kol_type and headers.get("KOL类型"):
                    set_cell(ws, headers, row_idx, "KOL类型", kol_type)
                    print(f"  - 抓取KOL类型: {kol_type}")

                # 平台价格与合作形式
                cooperation_form = infer_cooperation_form_from_notes(
                    notes_detail, blogger_data
                )
                if cooperation_form and headers.get("合作形式"):
                    set_cell(ws, headers, row_idx, "合作形式", cooperation_form)
                if cooperation_form == "图文" and headers.get("平台价格"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "平台价格",
                        get_page_price_value(page, "图文笔记一口价"),
                    )
                elif cooperation_form == "视频" and headers.get("平台价格"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "平台价格",
                        get_page_price_value(page, "视频笔记一口价"),
                    )

                # 粉丝画像
                gender_data = fans_profile.get("gender") or {}
                age_data = fans_profile.get("ages") or []

                if headers.get("女粉占比"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "女粉占比",
                        to_ratio_decimal(gender_data.get("female")),
                    )
                if headers.get("18-24年龄占比"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "18-24年龄占比",
                        to_ratio_decimal(find_percent(age_data, "18-24")),
                    )
                if headers.get("25-34年龄占比"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "25-34年龄占比",
                        to_ratio_decimal(find_percent(age_data, "25-34")),
                    )
                if headers.get("35-44年龄占比"):
                    set_cell(
                        ws,
                        headers,
                        row_idx,
                        "35-44年龄占比",
                        to_ratio_decimal(find_percent(age_data, "35-44")),
                    )

                # 数据概览中位数
                notes_list = (notes_detail or {}).get("list", [])
                fallback_read, fallback_interact = get_notes_median_fallback(
                    notes_detail, top_n=4
                )

                if headers.get("近30天阅读中位数"):
                    read_value = summary_data.get("mValidRawReadFeedNum", 0) or 0
                    if not read_value:
                        read_value = fallback_read
                        fallback_rows.append(
                            {
                                "row": row_idx,
                                "kol": kol_name,
                                "field": "阅读",
                                "fallback": fallback_read,
                            }
                        )
                    set_cell(ws, headers, row_idx, "近30天阅读中位数", read_value)

                if headers.get("近30天互动中位"):
                    interact_value = summary_data.get("mEngagementNum", 0) or 0
                    if not interact_value:
                        interact_value = fallback_interact
                        fallback_rows.append(
                            {
                                "row": row_idx,
                                "kol": kol_name,
                                "field": "互动",
                                "fallback": fallback_interact,
                            }
                        )
                    set_cell(ws, headers, row_idx, "近30天互动中位", interact_value)

                # 厨房场景图
                if headers.get("厨房场景图"):
                    kitchen_note = pick_kitchen_note(notes_list)
                    if kitchen_note and kitchen_note.get("imgUrl"):
                        kitchen_path = os.path.join(
                            SCREENSHOT_DIR,
                            f"row_{row_idx}_{sanitize_filename(user_id)}",
                            "kitchen.png",
                        )
                        if download_image(kitchen_note.get("imgUrl"), kitchen_path):
                            queue_image_for_cell(
                                ws,
                                headers,
                                wb._codex_image_jobs[ws.title],
                                row_idx,
                                "厨房场景图",
                                kitchen_path,
                                170,
                                120,
                            )
                            print(f"  - 抓取厨房场景图: {kitchen_path}")
                        else:
                            set_cell(
                                ws,
                                headers,
                                row_idx,
                                "厨房场景图",
                                "整体厨房图建联博主后给到",
                            )
                            print(f"  - 厨房场景图下载失败")
                    else:
                        set_cell(
                            ws,
                            headers,
                            row_idx,
                            "厨房场景图",
                            "整体厨房图建联博主后给到",
                        )
                        print(f"  - 未找到厨房相关笔记")

            except Exception as e:
                print(f"  - 蒲公英数据抓取失败: {e}")

            print(f"  - 第 {row_idx} 行处理完成")
            processed_count += 1
            if processed_count % SAVE_EVERY_ROWS == 0:
                last_saved_path = save_progress(wb, OUTPUT_PATH, processed_count)

        last_saved_path = save_progress(wb, OUTPUT_PATH, processed_count)
        context.close()
        cleanup_dir(SCREENSHOT_DIR)

    if fallback_rows:
        unique_rows = {}
        for item in fallback_rows:
            row = item["row"]
            if row not in unique_rows:
                unique_rows[row] = item
        print(
            f"\n⚠️  以下 {len(unique_rows)} 行走了 fallback（接口返回0，重试后仍失败）："
        )
        print(f"{'Excel行号':<10} {'KOL名称':<20} {'字段':<8} {'fallback值':<12}")
        print("-" * 55)
        for row, item in unique_rows.items():
            print(
                f"{item['row']:<10} {item['kol'][:18]:<20} {item['field']:<8} {item['fallback']:<12.0f}"
            )

    print(f"\n[完成] 所有任务处理完毕！结果已保存到:\n  {last_saved_path}")


if __name__ == "__main__":
    run_extraction()
