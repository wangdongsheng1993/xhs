import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

from playwright.sync_api import sync_playwright

from xhs_extractor import (
    DEBUG_MAX_ROWS,
    DEBUG_TARGET_NAME,
    DEBUG_TARGET_ROW,
    DEBUG_VERBOSE,
    PAGE_WAIT_SECONDS,
    USER_DATA_DIR,
    extract_user_id_from_pgy_url,
    find_cached_json_by_keyword,
    find_percent,
    format_percent,
    get_first_visible,
    get_homepage_url_from_profile_click,
    get_json_with_cache,
    get_level,
    get_notes_median_fallback,
    get_page_data_value,
    get_page_price_value,
    goto_with_retry,
    infer_cooperation_form_from_notes,
    parse_w_value,
    wait_for_login_confirmation,
)


# 飞书应用 ID 建议优先从环境变量读取。
# 如果环境变量未提供，这里回退到当前会话里已确认可用的应用 ID。
FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "cli_a94b97e076b91bc6").strip()

# 飞书应用密钥涉及敏感信息，优先从环境变量读取。
# 如果你不想写到脚本里，运行前设置 FEISHU_APP_SECRET 即可。
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "ZFC9t6OeZrPNuV559jmdoc4lvX7K7LON").strip()

# 这里填写 wiki 链接里的 token。
FEISHU_WIKI_TOKEN = os.getenv("FEISHU_WIKI_TOKEN", "GHo0wWtKRiU2d6k13WRceJponuv").strip()

# 这里填写链接里的 sheet 参数，也就是工作表 ID。
FEISHU_SHEET_ID = os.getenv("FEISHU_SHEET_ID", "Z5aXsY").strip()

# 读取范围给得稍宽一些，方便覆盖常见表格长度。
FEISHU_READ_RANGE = os.getenv("FEISHU_READ_RANGE", "A1:AZ500").strip()

# 这里指定飞书里要处理的 sheet 名称，仅用于日志提示和结果校验。
SHEET_NAME = "小红书品牌-KOL"

# 飞书开放平台统一网关。
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def should_process_feishu_row(sheet_row, row, processed_count):
    """按调试参数筛选当前飞书表格行。"""
    if DEBUG_TARGET_ROW:
        try:
            if sheet_row != int(DEBUG_TARGET_ROW):
                return False
        except Exception:
            pass

    if DEBUG_TARGET_NAME:
        kol_name = str(row.get("KOL名称", "")).strip()
        if DEBUG_TARGET_NAME.lower() not in kol_name.lower():
            return False

    if DEBUG_MAX_ROWS:
        try:
            if processed_count >= int(DEBUG_MAX_ROWS):
                return False
        except Exception:
            pass

    return True


def col_to_letter(col_index):
    """把 1-based 列号转换成 Excel/Sheets 列字母。"""
    result = []
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def normalize_cell_value(value):
    """把飞书返回的单元格值统一转成更适合脚本处理的文本。"""
    if value is None:
        return ""
    if isinstance(value, list):
        parts = []
        for item in value:
            if isinstance(item, dict):
                if item.get("link"):
                    parts.append(str(item.get("link")))
                elif item.get("text"):
                    parts.append(str(item.get("text")))
                else:
                    parts.append(str(item))
            else:
                parts.append(str(item))
        return "".join(parts)
    return str(value)


def to_percent_decimal(value):
    """把占比值统一转换成飞书表格可计算的小数形式。"""
    if value in (None, ""):
        return ""

    if isinstance(value, str):
        value = value.strip()
        if not value:
            return ""
        if value.endswith("%"):
            value = value[:-1]

    try:
        num = float(value)
    except Exception:
        return ""

    if num > 1:
        num /= 100

    return round(num, 6)


def feishu_request(method, path, token="", params=None, body=None):
    """统一封装飞书 API 请求，并在失败时抛出清晰错误。"""
    url = f"{FEISHU_BASE_URL}{path}"
    if params:
        url = f"{url}?{urllib.parse.urlencode(params, doseq=True)}"

    data = None
    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")

    request = urllib.request.Request(url, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        error_text = e.read().decode("utf-8", errors="replace")
        raise Exception(f"飞书接口失败: {method} {url} -> HTTP {e.code}: {error_text}")

    if payload.get("code") not in (0, None):
        raise Exception(f"飞书接口返回异常: {method} {url} -> {payload}")
    return payload.get("data", payload)


def get_tenant_access_token():
    """通过 app_id 和 app_secret 获取 tenant_access_token。"""
    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        raise Exception("缺少 FEISHU_APP_ID 或 FEISHU_APP_SECRET，无法访问飞书在线表格。")

    data = feishu_request(
        "POST",
        "/auth/v3/tenant_access_token/internal",
        body={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
    )
    return data.get("tenant_access_token", "")


def resolve_spreadsheet_token(tenant_access_token):
    """通过 wiki token 解析底层表格 token。"""
    data = feishu_request(
        "GET",
        "/wiki/v2/spaces/get_node",
        token=tenant_access_token,
        params={"token": FEISHU_WIKI_TOKEN},
    )
    node = data.get("node") or {}
    spreadsheet_token = node.get("obj_token", "")
    obj_type = node.get("obj_type", "")
    if obj_type != "sheet":
        raise Exception(f"当前 wiki 节点不是飞书表格，obj_type={obj_type}")
    if not spreadsheet_token:
        raise Exception("无法从 wiki 节点解析 spreadsheet_token")
    return spreadsheet_token


def read_sheet_rows(tenant_access_token, spreadsheet_token):
    """读取飞书表格指定范围，并转成按表头索引的行对象。"""
    target_range = f"{FEISHU_SHEET_ID}!{FEISHU_READ_RANGE}"
    data = feishu_request(
        "GET",
        f"/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_get",
        token=tenant_access_token,
        params={"ranges": target_range},
    )

    value_ranges = data.get("valueRanges") or []
    if not value_ranges:
        raise Exception("飞书表格读取成功，但没有返回任何 valueRanges。")

    values = value_ranges[0].get("values") or []
    if not values:
        raise Exception("飞书表格读取成功，但目标范围内没有数据。")

    headers = [normalize_cell_value(cell).strip() for cell in values[0]]
    rows = []
    for idx, raw_row in enumerate(values[1:], start=2):
        row_map = {}
        for col_idx, header in enumerate(headers):
            cell_value = raw_row[col_idx] if col_idx < len(raw_row) else ""
            row_map[header] = normalize_cell_value(cell_value).strip()
        rows.append({"sheet_row": idx, "data": row_map})

    return headers, rows


def build_update_payload(headers, sheet_row, updates):
    """把字段更新转换成飞书 values_batch_update 所需的 valueRanges。"""
    header_to_col = {header: idx + 1 for idx, header in enumerate(headers) if header}
    value_ranges = []

    for header, value in updates.items():
        col_index = header_to_col.get(header)
        if not col_index:
            continue
        col_letter = col_to_letter(col_index)
        value_ranges.append(
            {
                "range": f"{FEISHU_SHEET_ID}!{col_letter}{sheet_row}:{col_letter}{sheet_row}",
                "values": [[value]],
            }
        )

    return value_ranges


def write_row_updates(tenant_access_token, spreadsheet_token, headers, sheet_row, updates):
    """把单行更新写回飞书表格。"""
    value_ranges = build_update_payload(headers, sheet_row, updates)
    if not value_ranges:
        return

    feishu_request(
        "POST",
        f"/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update",
        token=tenant_access_token,
        body={"valueRanges": value_ranges},
    )


def run_extraction():
    """飞书版主流程：在线读表、抓取蒲公英、在线写回结果。"""
    tenant_access_token = get_tenant_access_token()
    spreadsheet_token = resolve_spreadsheet_token(tenant_access_token)
    headers, rows = read_sheet_rows(tenant_access_token, spreadsheet_token)

    if DEBUG_TARGET_NAME or DEBUG_TARGET_ROW or DEBUG_MAX_ROWS:
        print("【调试模式】已启用单条/限量处理")
        if DEBUG_TARGET_NAME:
            print(f"  - 按名称筛选: {DEBUG_TARGET_NAME}")
        if DEBUG_TARGET_ROW:
            print(f"  - 指定飞书行号: {DEBUG_TARGET_ROW}")
        if DEBUG_MAX_ROWS:
            print(f"  - 最多处理条数: {DEBUG_MAX_ROWS}")

    print(f"已连接飞书在线表格，sheet={SHEET_NAME}，共读取 {len(rows)} 条候选数据。")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            slow_mo=500,
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
            page.goto("https://pgy.xiaohongshu.com", wait_until="domcontentloaded", timeout=15000)
        except Exception:
            pass

        wait_for_login_confirmation()

        print("\n【步骤 2: 开始执行采集】")
        processed_count = 0

        for item in rows:
            sheet_row = item["sheet_row"]
            row = item["data"]

            if not row.get("主页链接") and not row.get("蒲公英链接"):
                continue

            if not should_process_feishu_row(sheet_row, row, processed_count):
                continue

            kol_name = row.get("KOL名称") or f"Row {sheet_row}"
            pgy_url = row.get("蒲公英链接")
            print(f"\n>>> [{sheet_row}] 正在处理: {kol_name}")

            if not pgy_url or not str(pgy_url).startswith("http"):
                print("  - 当前行没有有效的蒲公英链接，已跳过。")
                continue

            api_cache.clear()

            try:
                goto_with_retry(page, str(pgy_url), wait_until="domcontentloaded", timeout=30000, retries=2)
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
                    f"/api/solar/kol/data_v2/notes_detail?advertiseSwitch=1&orderType=1&pageNumber=1&pageSize=8&userId={user_id}&noteType=4&isThirdPlatform=0",
                ) or {}

                try:
                    summary_data = get_json_with_cache(page, api_cache, summary_api) or {}
                except Exception as summary_err:
                    summary_data = {}
                    print(f"  - 数据概览接口获取失败，已跳过接口中位数: {summary_err}")

                updates = {}

                red_id = blogger_data.get("redId", "")
                if red_id:
                    updates["ID"] = red_id
                    print(f"  - 抓取小红书号: {red_id}")

                    homepage_url = get_homepage_url_from_profile_click(context, page, red_id)
                    if homepage_url:
                        updates["主页链接"] = homepage_url
                        print(f"  - 抓取主页链接: {homepage_url}")

                fans_text = get_page_data_value(page, "粉丝数")
                likes_text = get_page_data_value(page, "获赞与收藏")
                fans_w = parse_w_value(fans_text)
                updates["粉丝量（w）"] = fans_w
                updates["量级"] = get_level(fans_w)
                updates["赞藏量（w）"] = parse_w_value(likes_text)

                cooperation_form = infer_cooperation_form_from_notes(notes_detail, blogger_data)
                if cooperation_form:
                    updates["合作形式"] = cooperation_form
                if cooperation_form == "图文":
                    updates["平台价格"] = get_page_price_value(page, "图文笔记一口价")
                elif cooperation_form == "视频":
                    updates["平台价格"] = get_page_price_value(page, "视频笔记一口价")

                gender_data = fans_profile.get("gender") or {}
                age_data = fans_profile.get("ages") or []
                device_data = fans_profile.get("devices") or []
                updates["女粉占比"] = to_percent_decimal(gender_data.get("female"))
                updates["18-24年龄占比"] = to_percent_decimal(find_percent(age_data, "18-24"))
                updates["25-34年龄占比（不超50%）"] = to_percent_decimal(find_percent(age_data, "25-34"))
                updates["35-44年龄占比（前2）"] = to_percent_decimal(find_percent(age_data, "35-44"))
                updates["苹果用户占比"] = to_percent_decimal(find_percent(device_data, ["apple inc.", "apple", "苹果"]))
                updates["华为用户占比"] = to_percent_decimal(find_percent(device_data, ["huawei", "华为"]))

                fallback_read, fallback_interact = get_notes_median_fallback(notes_detail, top_n=4)
                read_value = summary_data.get("mValidRawReadFeedNum", 0) or 0
                interact_value = summary_data.get("mEngagementNum", 0) or 0
                if not read_value:
                    read_value = fallback_read
                if not interact_value:
                    interact_value = fallback_interact
                updates["近30天预估阅读量\n(近30天阅读中位数）"] = read_value
                updates["近30天互动量\n（近30天互动中位）"] = interact_value

                write_row_updates(tenant_access_token, spreadsheet_token, headers, sheet_row, updates)
                print(f"  - 已写回飞书在线表格: 第 {sheet_row} 行")
            except Exception as e:
                print(f"  - 飞书在线行处理失败: {e}")

            processed_count += 1

        context.close()

    print("\n[完成] 飞书在线表格处理结束。")


if __name__ == "__main__":
    run_extraction()
