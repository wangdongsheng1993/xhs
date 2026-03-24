import os
import sys

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
    get_first_visible,
    get_homepage_url_from_profile_click,
    get_json_with_cache,
    get_level,
    get_notes_median_fallback,
    get_page_data_value,
    get_page_price_value,
    goto_with_retry,
    parse_w_value,
    wait_for_login_confirmation,
)
from xhs_ecommerce_extractor import (
    REFERENCE_CASE_KEYWORDS,
    build_note_link_from_note_id,
    build_user_interest_text,
    click_tab,
    collect_recent_notes,
    copy_note_link_from_case_detail,
    count_recent_30_notes,
    count_recent_90_hot_notes,
    infer_kol_content_type,
    pick_kitchen_note,
    pick_reference_case_note,
)
from xhs_extractor_feishu import (
    FEISHU_APP_ID,
    FEISHU_APP_SECRET,
    FEISHU_WIKI_TOKEN,
    feishu_request,
    get_tenant_access_token,
    normalize_cell_value,
    resolve_spreadsheet_token,
    should_process_feishu_row,
    to_percent_decimal,
)


# 电商表对应的 sheet 参数来自用户给出的飞书链接。
FEISHU_SHEET_ID = os.getenv("FEISHU_SHEET_ID", "MOhHCP").strip()

# 读取范围给得稍宽，覆盖当前线上电商表的列数与行数。
FEISHU_READ_RANGE = os.getenv("FEISHU_READ_RANGE", "A1:BC500").strip()

# 仅用于日志提示。
SHEET_NAME = "小红书电商-KOL"


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def read_ecommerce_sheet_rows(tenant_access_token, spreadsheet_token):
    """读取电商飞书表格指定范围，并转成按表头索引的行对象。"""
    target_range = f"{FEISHU_SHEET_ID}!{FEISHU_READ_RANGE}"
    data = feishu_request(
        "GET",
        f"/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_get",
        token=tenant_access_token,
        params={"ranges": target_range},
    )

    value_ranges = data.get("valueRanges") or []
    if not value_ranges:
        raise Exception("飞书电商表读取成功，但没有返回任何 valueRanges。")

    values = value_ranges[0].get("values") or []
    if not values:
        raise Exception("飞书电商表读取成功，但目标范围内没有数据。")

    headers = [normalize_cell_value(cell).strip() for cell in values[0]]
    rows = []
    for idx, raw_row in enumerate(values[1:], start=2):
        row_map = {}
        for col_idx, header in enumerate(headers):
            cell_value = raw_row[col_idx] if col_idx < len(raw_row) else ""
            row_map[header] = normalize_cell_value(cell_value).strip()
        rows.append({"sheet_row": idx, "data": row_map})

    return headers, rows


def col_to_letter(col_index):
    """把 1-based 列号转换成 Excel/Sheets 列字母。"""
    result = []
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def build_ecommerce_update_payload(headers, sheet_row, updates):
    """把字段更新转换成电商 sheet 专用的 valueRanges。"""
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


def write_ecommerce_row_updates(tenant_access_token, spreadsheet_token, headers, sheet_row, updates):
    """把单行更新写回飞书电商表格。"""
    value_ranges = build_ecommerce_update_payload(headers, sheet_row, updates)
    if not value_ranges:
        return

    feishu_request(
        "POST",
        f"/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update",
        token=tenant_access_token,
        body={"valueRanges": value_ranges},
    )


def run_extraction():
    """飞书电商表主流程：在线读表、抓取蒲公英、在线写回数值和链接字段。"""
    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        raise Exception("缺少飞书应用凭证，无法访问线上电商表。")

    tenant_access_token = get_tenant_access_token()
    spreadsheet_token = resolve_spreadsheet_token(tenant_access_token)
    headers, rows = read_ecommerce_sheet_rows(tenant_access_token, spreadsheet_token)

    if DEBUG_TARGET_NAME or DEBUG_TARGET_ROW or DEBUG_MAX_ROWS:
        print("【调试模式】已启用单条/限量处理")
        if DEBUG_TARGET_NAME:
            print(f"  - 按名称筛选: {DEBUG_TARGET_NAME}")
        if DEBUG_TARGET_ROW:
            print(f"  - 指定飞书行号: {DEBUG_TARGET_ROW}")
        if DEBUG_MAX_ROWS:
            print(f"  - 最多处理条数: {DEBUG_MAX_ROWS}")

    print(f"已连接飞书在线表格，sheet={SHEET_NAME}，共读取 {len(rows)} 条候选数据。")
    print("说明：飞书在线版当前只写回文字、数字和链接字段，截图类列暂不做在线插图。")

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

                blogger_api = f"https://pgy.xiaohongshu.com/api/solar/cooperator/user/blogger/{user_id}"
                fans_profile_api = f"https://pgy.xiaohongshu.com/api/solar/kol/data/{user_id}/fans_profile"
                summary_api = f"https://pgy.xiaohongshu.com/api/pgy/kol/data/data_summary?userId={user_id}&business=1"

                click_tab(page, "数据概览")
                click_tab(page, "合作笔记")
                click_tab(page, "笔记数据")

                blogger_data = get_json_with_cache(page, api_cache, blogger_api) or {}
                fans_profile = get_json_with_cache(page, api_cache, fans_profile_api) or {}
                coop_notes_detail = find_cached_json_by_keyword(
                    api_cache,
                    f"/api/solar/kol/data_v2/notes_detail?advertiseSwitch=1&orderType=1&pageNumber=1&pageSize=8&userId={user_id}&noteType=4&isThirdPlatform=0",
                ) or {}

                try:
                    summary_data = get_json_with_cache(page, api_cache, summary_api) or {}
                except Exception as summary_err:
                    summary_data = {}
                    print(f"  - 数据概览接口获取失败，已跳过中位数: {summary_err}")

                notes = collect_recent_notes(page, api_cache, user_id)
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

                kol_type = infer_kol_content_type(notes, blogger_data)
                updates["KOL类型（图文/视频）"] = kol_type
                if kol_type == "图文":
                    updates["平台价格"] = get_page_price_value(page, "图文笔记一口价")
                elif kol_type == "视频":
                    updates["平台价格"] = get_page_price_value(page, "视频笔记一口价")

                device_data = fans_profile.get("devices") or []
                updates["苹果用户占比"] = to_percent_decimal(find_percent(device_data, ["apple inc.", "apple", "苹果"]))
                updates["华为用户占比"] = to_percent_decimal(find_percent(device_data, ["huawei", "华为"]))

                col_read = "近30天预估阅读量\n(近30天阅读中位数）"
                col_interact = "近30天互动量\n（近30天互动中位）"
                fallback_read, fallback_interact = get_notes_median_fallback(coop_notes_detail, top_n=4)
                read_value = summary_data.get("mValidRawReadFeedNum", 0) or 0
                interact_value = summary_data.get("mEngagementNum", 0) or 0
                if not read_value:
                    read_value = fallback_read
                if not interact_value:
                    interact_value = fallback_interact
                updates[col_read] = read_value
                updates[col_interact] = interact_value

                note_count_30 = count_recent_30_notes(notes)
                hot_count_90 = count_recent_90_hot_notes(notes)
                if not note_count_30 and summary_data.get("noteNumber") is not None:
                    note_count_30 = summary_data.get("noteNumber") or 0
                updates["近30天笔记数量"] = note_count_30
                updates["近90天爆文篇数"] = hot_count_90

                updates["用户兴趣"] = build_user_interest_text(fans_profile)

                reference_note = pick_reference_case_note((coop_notes_detail or {}).get("list") or [])
                if reference_note:
                    reference_link = copy_note_link_from_case_detail(
                        page,
                        (coop_notes_detail or {}).get("list") or [],
                        reference_note,
                    )
                    if not reference_link:
                        reference_link = build_note_link_from_note_id(reference_note.get("noteId"))
                    if reference_link:
                        updates["参考案例"] = reference_link

                kitchen_note = pick_kitchen_note(notes)
                if kitchen_note and kitchen_note.get("imgUrl"):
                    updates["达人厨房图"] = "厨房图已识别，飞书在线版暂不支持自动插图"
                else:
                    updates["达人厨房图"] = "整体厨房图建联博主后给到"

                # 飞书在线版暂不向截图列写图片。
                # 这里先写说明文字，避免整列空白让人误以为脚本没处理。
                if "性别占比（截图）" in headers:
                    updates["性别占比（截图）"] = "飞书在线版暂不支持自动插图"
                if "粉丝年龄占比（截图）" in headers:
                    updates["粉丝年龄占比（截图）"] = "飞书在线版暂不支持自动插图"
                if "粉丝地域（截图）" in headers:
                    updates["粉丝地域（截图）"] = "飞书在线版暂不支持自动插图"

                write_ecommerce_row_updates(tenant_access_token, spreadsheet_token, headers, sheet_row, updates)
                print(f"  - 已写回飞书在线表格: 第 {sheet_row} 行")
            except Exception as e:
                print(f"  - 飞书在线电商行处理失败: {e}")

            processed_count += 1

        context.close()

    print("\n[完成] 飞书在线电商表处理结束。")


if __name__ == "__main__":
    run_extraction()
