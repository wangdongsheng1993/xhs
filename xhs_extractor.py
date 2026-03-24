import pandas as pd
import re
import time
import os
import sys
import statistics
from playwright.sync_api import sync_playwright

# --- 配置区 ---
EXCEL_PATH = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表 副本.xlsx"
OUTPUT_PATH = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
SHEET_NAME = "小红书品牌-KOL"
# 浏览器数据目录，用于保存登录状态
USER_DATA_DIR = os.path.join(os.getcwd(), "browser_session")
# 每个页面打开后等待的秒数
PAGE_WAIT_SECONDS = 10
DEBUG_TARGET_NAME = os.getenv("XHS_DEBUG_NAME", "").strip()
DEBUG_TARGET_ROW = os.getenv("XHS_DEBUG_ROW", "").strip()
DEBUG_MAX_ROWS = os.getenv("XHS_DEBUG_MAX_ROWS", "").strip()
DEBUG_VERBOSE = os.getenv("XHS_DEBUG_VERBOSE", "0").strip() == "1"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

def parse_w_value(text):
    """处理带 'w' 或 '万' 的数值字符串"""
    if not text or text == '-': return 0.0
    text = text.replace('w', '').replace('万', '').replace('+', '').replace(',', '').strip()
    try:
        if 'k' in text.lower():
            return float(text.lower().replace('k', '')) / 10
        return float(text)
    except:
        return 0.0

def get_level(fans_w):
    """根据粉丝量（w）计算量级"""
    if fans_w < 1: return "KOC"
    if 1 <= fans_w < 10: return "尾部"
    if 10 <= fans_w < 30: return "腰部"
    if 30 <= fans_w < 50: return "肩部"
    if fans_w >= 50: return "头部"
    return "未知"

def get_first_visible(locator, timeout=3000):
    """从多个候选元素中返回第一个可见的 locator，避免 strict mode 冲突"""
    try:
        count = locator.count()
        for i in range(count):
            item = locator.nth(i)
            try:
                if item.is_visible(timeout=timeout):
                    return item
            except:
                pass
    except:
        pass
    return None

def wait_for_login_confirmation():
    """兼容交互式终端和无 stdin 的运行环境"""
    prompt = "\n>>> 登录完成后，按 Enter 键开始采集..."
    try:
        if sys.stdin and sys.stdin.isatty():
            input(prompt)
            return
    except EOFError:
        pass

    print(prompt)
    print("检测到当前环境无法交互输入，等待 10 秒后继续，并复用已有登录态...")
    time.sleep(10)

def extract_user_id_from_pgy_url(url):
    match = re.search(r'/blogger-detail/([^/?]+)', str(url))
    return match.group(1) if match else None

def fetch_json(page, url):
    """通过页面上下文发起带登录态的请求"""
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
            return {
                ok: response.ok,
                status: response.status,
                text
            };
        }
        """,
        url
    )

    if not result.get("ok"):
        raise Exception(f"接口请求失败: {url} (HTTP {result.get('status')})")

    try:
        payload = page.evaluate("text => JSON.parse(text)", result.get("text", ""))
    except Exception as e:
        raise Exception(f"接口返回非 JSON: {url} ({e})")

    if isinstance(payload, dict):
        if payload.get("code") not in (None, 0):
            raise Exception(f"接口返回异常: {url} code={payload.get('code')} msg={payload.get('msg')}")
        if payload.get("success") is False:
            raise Exception(f"接口返回失败: {url} msg={payload.get('msg')}")
        if "data" in payload:
            if DEBUG_VERBOSE:
                if isinstance(payload.get("data"), dict):
                    print(f"    [DEBUG] 接口成功，data keys: {list(payload.get('data').keys())[:10]}")
                else:
                    print(f"    [DEBUG] 接口成功，data type: {type(payload.get('data')).__name__}")
            return payload.get("data")

    return payload

def get_cached_json(api_cache, url):
    payload = api_cache.get(url)
    if isinstance(payload, dict):
        if payload.get("code") not in (None, 0):
            raise Exception(f"接口返回异常: {url} code={payload.get('code')} msg={payload.get('msg')}")
        if payload.get("success") is False:
            raise Exception(f"接口返回失败: {url} msg={payload.get('msg')}")
        if "data" in payload:
            return payload.get("data")
    return payload

def get_json_with_cache(page, api_cache, url):
    """优先读取页面真实请求的响应，拿不到再用 fetch 补一次"""
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

def format_percent(value):
    """把 0-1 的比例转为四舍五入后的百分比字符串"""
    try:
        if value is None or value == "":
            return ""
        num = float(value)
        if num <= 1:
            num *= 100
        return f"{int(num + 0.5)}%"
    except:
        return ""

def find_percent(items, targets, name_keys=("group", "name", "desc")):
    """从接口列表中按名称匹配占比"""
    if isinstance(targets, str):
        targets = [targets]

    target_set = {str(x).strip().lower() for x in targets}
    for item in items or []:
        for key in name_keys:
            value = item.get(key)
            if value is not None and str(value).strip().lower() in target_set:
                return item.get("percent")
    return None

def get_cooperation_form(blogger_data):
    forms = []
    if blogger_data.get("pictureState") == 1:
        forms.append("图文")
    if blogger_data.get("videoState") == 1:
        forms.append("视频")
    return "+".join(forms)

def get_page_data_value(page, label):
    """从左侧卡片读取粉丝数、获赞与收藏的页面显示值"""
    return page.evaluate(
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
        label
    ) or ""

def get_page_price_value(page, label):
    """从合作报价卡片读取指定报价"""
    text = page.evaluate(
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
        label
    ) or ""
    price_text = re.sub(r"[^\d.]", "", text)
    return float(price_text) if price_text else 0.0

def infer_cooperation_form_from_notes(notes_data, blogger_data):
    """根据笔记案例中视频/图文数量判断合作形式"""
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
    """把接口里的各种数值安全转成数字，失败时返回 0。"""
    try:
        if value in (None, ""):
            return 0
        return float(value)
    except:
        return 0


def get_note_read_count(note):
    """尽量从笔记数据里取出阅读量，兼容不同字段名。"""
    for key in ("readNum", "readCount", "readingNum", "read", "exposureNum", "impressionNum"):
        value = _to_number((note or {}).get(key))
        if value > 0:
            return value
    return 0


def get_note_interaction_count(note):
    """把点赞、收藏、评论、转发加总成互动量。"""
    note = note or {}
    return (
        _to_number(note.get("likeNum"))
        + _to_number(note.get("collectNum"))
        + _to_number(note.get("commentNum"))
        + _to_number(note.get("shareNum"))
    )


def get_notes_median_fallback(notes_data, top_n=4):
    """从前几条合作笔记里计算阅读和互动中位数，偶数条时取靠前的中间值。"""
    notes = (notes_data or {}).get("list") or []
    notes = notes[:top_n]
    if not notes:
        return 0, 0

    read_values = [get_note_read_count(note) for note in notes]
    interact_values = [get_note_interaction_count(note) for note in notes]

    read_values = [value for value in read_values if value > 0]
    interact_values = [value for value in interact_values if value > 0]

    # 这里不用标准 median，是因为 4 条数据时它会返回中间两项的平均值。
    # 当前业务希望取排序后真正存在的“中间值”，所以偶数条时取靠前的那个中间值。
    read_median = statistics.median_low(read_values) if read_values else 0
    interact_median = statistics.median_low(interact_values) if interact_values else 0
    return read_median, interact_median


def get_homepage_url_from_profile_click(context, page, red_id):
    """点击蒲公英页里的小红书号，读取新打开主页的真实地址。"""
    if not red_id:
        return ""

    # 先准备几种常见定位方式，尽量兼容页面结构的小变化。
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

            with context.expect_page(timeout=5000) as popup_info:
                clickable.click()

            popup = popup_info.value
            popup.wait_for_load_state("domcontentloaded", timeout=10000)
            popup.wait_for_timeout(2000)
            url = popup.url or ""
            popup.close()
            if url and url != "about:blank":
                return url
        except:
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
                    return href
            except:
                pass

    return ""


def goto_with_retry(page, url, wait_until="domcontentloaded", timeout=30000, retries=2):
    """页面跳转偶尔会被站内二次导航打断，这里做一次轻量重试。"""
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

def should_process_row(index, row, processed_count):
    """调试模式下按名称、行号或最大条数筛选"""
    if DEBUG_TARGET_ROW:
        try:
            target_row = int(DEBUG_TARGET_ROW)
            if index + 1 != target_row:
                return False
        except:
            pass

    if DEBUG_TARGET_NAME:
        kol_name = str(row.get('KOL名称', '')).strip()
        if DEBUG_TARGET_NAME.lower() not in kol_name.lower():
            return False

    if DEBUG_MAX_ROWS:
        try:
            if processed_count >= int(DEBUG_MAX_ROWS):
                return False
        except:
            pass

    return True

def run_extraction():
    # 1. 检查是否存在 Excel
    if not os.path.exists(EXCEL_PATH):
        print(f"错误: 找不到文件 {EXCEL_PATH}")
        return

    # 2. 读取 Excel
    print(f"正在读取 Excel: {EXCEL_PATH} ...")
    try:
        # 尝试读取，如果报错可能是文件被占用
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
        # 将可能写入数据的列转为 object 类型，避免 float64 类型冲突
        for col in ['平台价格', '粉丝量（w）', '赞藏量（w）', '量级', 'ID',
                    '女粉占比', '18-24年龄占比', '25-34年龄占比（不超50%）',
                    '35-44年龄占比（前2）', '苹果用户占比', '华为用户占比',
                    '合作形式']:
            if col in df.columns:
                df[col] = df[col].astype(object)
    except Exception as e:
        print(f"读取失败 (请检查文件是否被 Excel 打开): {e}")
        return
    
    # 3. 启动浏览器
    with sync_playwright() as p:
        # 使用持久化上下文，保存登录 Cookies
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            slow_mo=500
        )
        page = context.new_page()
        api_cache = {}

        def on_response(response):
            try:
                url = response.url
                content_type = response.headers.get("content-type", "")
                if "application/json" not in content_type and "text/json" not in content_type:
                    return
                if "pgy.xiaohongshu.com/api/" not in url:
                    return
                api_cache[url] = response.json()
            except:
                pass

        page.on("response", on_response)

        # --- 登录等待 ---
        print("\n" + "="*50)
        print("【步骤 1: 请手动登录蒲公英】")
        print("浏览器已打开，请在浏览器中完成登录：")
        print("  > 网址: https://pgy.xiaohongshu.com")
        print("  > 选择【我是代理商】进行账号密码登录")
        print("  > 登录完成后，回到此窗口按 Enter 键继续")
        print("="*50)

        # 仅打开首页，不做任何等待判断
        try:
            page.goto("https://pgy.xiaohongshu.com", wait_until="domcontentloaded", timeout=15000)
        except:
            pass  # 忽略超时，让用户自行操作

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
        for index, row in df.iterrows():
            # 跳过空行（主页链接和蒲公英链接都为空）
            if pd.isna(row.get('主页链接')) and pd.isna(row.get('蒲公英链接')):
                continue

            if not should_process_row(index, row, processed_count):
                continue
            
            kol_name = row.get('KOL名称', f"Row {index+1}")
            print(f"\n>>> [{index+1}/{len(df)}] 正在处理: {kol_name}")

            # --- 步骤 1: 通过蒲公英链接抓取全部数据 ---
            pgy_url = row.get('蒲公英链接')
            if pd.notna(pgy_url) and str(pgy_url).startswith("http"):
                try:
                    goto_with_retry(page, pgy_url, wait_until="domcontentloaded", timeout=30000, retries=2)
                    print(f"  - 蒲公英页已打开，等待 {PAGE_WAIT_SECONDS} 秒加载...")
                    page.wait_for_timeout(PAGE_WAIT_SECONDS * 1000)

                    user_id = extract_user_id_from_pgy_url(pgy_url)
                    if not user_id:
                        raise Exception("无法从蒲公英链接中解析 userId")
                    if DEBUG_VERBOSE:
                        print(f"    [DEBUG] userId: {user_id}")

                    data_tab = get_first_visible(page.get_by_text("数据概览", exact=True))
                    if data_tab:
                        data_tab.click()
                        page.wait_for_timeout(1200)
                    coop_note_tab = get_first_visible(page.get_by_text("合作笔记", exact=True))
                    if coop_note_tab:
                        coop_note_tab.click()
                        page.wait_for_timeout(2000)

                    blogger_api = f"https://pgy.xiaohongshu.com/api/solar/cooperator/user/blogger/{user_id}"
                    fans_profile_api = f"https://pgy.xiaohongshu.com/api/solar/kol/data/{user_id}/fans_profile"
                    summary_api = f"https://pgy.xiaohongshu.com/api/pgy/kol/data/data_summary?userId={user_id}&business=1"

                    blogger_data = get_json_with_cache(page, api_cache, blogger_api) or {}
                    fans_profile = get_json_with_cache(page, api_cache, fans_profile_api) or {}
                    notes_detail = find_cached_json_by_keyword(
                        api_cache,
                        f"/api/solar/kol/data_v2/notes_detail?advertiseSwitch=1&orderType=1&pageNumber=1&pageSize=8&userId={user_id}&noteType=4&isThirdPlatform=0"
                    ) or {}

                    try:
                        summary_data = get_json_with_cache(page, api_cache, summary_api) or {}
                    except Exception as summary_err:
                        summary_data = {}
                        print(f"  - 数据概览接口获取失败，已跳过中位数: {summary_err}")

                    # 小红书号
                    red_id = blogger_data.get("redId", "")
                    if red_id:
                        df.at[index, 'ID'] = red_id
                        print(f"  - 抓取小红书号: {red_id}")

                        homepage_url = get_homepage_url_from_profile_click(context, page, red_id)
                        if homepage_url:
                            df.at[index, '主页链接'] = homepage_url
                            print(f"  - 抓取主页链接: {homepage_url}")

                    # 基础统计：直接取页面显示值
                    fans_text = get_page_data_value(page, "粉丝数")
                    likes_text = get_page_data_value(page, "获赞与收藏")
                    fans_w = parse_w_value(fans_text)
                    df.at[index, '粉丝量（w）'] = fans_w
                    df.at[index, '量级'] = get_level(fans_w)
                    df.at[index, '赞藏量（w）'] = parse_w_value(likes_text)

                    # 平台价格与合作形式
                    cooperation_form = infer_cooperation_form_from_notes(notes_detail, blogger_data)
                    if cooperation_form:
                        df.at[index, '合作形式'] = cooperation_form
                    if cooperation_form == "图文":
                        df.at[index, '平台价格'] = get_page_price_value(page, "图文笔记一口价")
                    elif cooperation_form == "视频":
                        df.at[index, '平台价格'] = get_page_price_value(page, "视频笔记一口价")

                    # 粉丝画像
                    gender_data = fans_profile.get("gender") or {}
                    age_data = fans_profile.get("ages") or []
                    device_data = fans_profile.get("devices") or []

                    df.at[index, '女粉占比'] = format_percent(gender_data.get("female"))
                    df.at[index, '18-24年龄占比'] = format_percent(find_percent(age_data, "18-24"))
                    df.at[index, '25-34年龄占比（不超50%）'] = format_percent(find_percent(age_data, "25-34"))
                    df.at[index, '35-44年龄占比（前2）'] = format_percent(find_percent(age_data, "35-44"))
                    df.at[index, '苹果用户占比'] = format_percent(find_percent(device_data, ["apple inc.", "apple", "苹果"]))
                    df.at[index, '华为用户占比'] = format_percent(find_percent(device_data, ["huawei", "华为"]))

                    # 数据概览中位数
                    col_read = "近30天预估阅读量\n(近30天阅读中位数）"
                    col_interact = "近30天互动量\n（近30天互动中位）"
                    fallback_read, fallback_interact = get_notes_median_fallback(notes_detail, top_n=4)
                    if col_read in df.columns:
                        read_value = summary_data.get("mValidRawReadFeedNum", 0) or 0
                        if not read_value:
                            read_value = fallback_read
                        df.at[index, col_read] = read_value
                    if col_interact in df.columns:
                        interact_value = summary_data.get("mEngagementNum", 0) or 0
                        if not interact_value:
                            interact_value = fallback_interact
                        df.at[index, col_interact] = interact_value

                except Exception as e:
                    print(f"  - 蒲公英数据抓取失败: {e}")

            # 每行处理完即时保存到新文件，防止原文件被占用
            try:
                df.to_excel(OUTPUT_PATH, sheet_name=SHEET_NAME, index=False)
                print(f"  - 已保存进度到: {OUTPUT_PATH}")
            except Exception as save_err:
                print(f"警告: 保存失败: {save_err}")

            processed_count += 1
                
    print(f"\n[完成] 所有任务处理完毕！结果已保存到:\n  {OUTPUT_PATH}")

if __name__ == "__main__":
    run_extraction()
