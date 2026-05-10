import argparse
import csv
import os
import random
import sys
import time

from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
DEFAULT_USER_DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "browser_session")

from fix_xhs_note_urls import (
    check_verification_popup,
    extract_note_id,
    is_blank,
    is_valid_note_link,
    locate_note_url,
    normalize_homepage_url,
    normalize_note_url,
    normalize_text,
    random_delay,
)
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

NOTE_URL_HEADER = "笔记官方地址"
TITLE_HEADER = "笔记标题"
HOMEPAGE_HEADER = "主页链接"


def read_failed_rows(failed_csv_path):
    failed = {}
    with open(failed_csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            row_num = row.get("行号", "")
            if row_num:
                failed[int(row_num)] = {
                    "title": str(row.get("笔记标题", "") or "").strip(),
                    "homepage_url": str(row.get("主页链接", "") or "").strip(),
                    "old_url": str(row.get("原笔记地址", "") or "").strip(),
                    "reason": str(row.get("失败原因", "") or "").strip(),
                }
    return failed


def retry_failed(
    result_csv_path,
    failed_csv_path,
    max_scrolls=15,
    per_item_timeout=20,
    detail_wait_ms=8000,
    login_wait=20,
    user_data_dir=DEFAULT_USER_DATA_DIR,
    skip_no_title=True,
    target_date=None,
):
    failed_rows = read_failed_rows(failed_csv_path)
    if not failed_rows:
        print("没有失败数据需要重试。", flush=True)
        return

    print(f"读取到 {len(failed_rows)} 条失败数据，准备重试...", flush=True)

    updated = 0
    still_failed = 0
    failed_items = []
    skipped_empty = 0
    skipped_no_title = 0
    timeout_count = 0
    error_count = 0
    not_found_count = 0
    no_note_id_count = 0

    print(f"浏览器会话目录: {user_data_dir}", flush=True)
    target_date_str = target_date.strftime("%Y-%m-%d") if target_date else "未设置"
    print(f"参数: max_scrolls={max_scrolls}, per_item_timeout={per_item_timeout}, detail_wait_ms={detail_wait_ms}, login_wait={login_wait}, skip_no_title={skip_no_title}, target_date={target_date_str}", flush=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            slow_mo=80,
            viewport={"width": 1440, "height": 1000},
        )
        page = context.pages[0] if context.pages else context.new_page()

        if login_wait:
            first_homepage = ""
            for row_num in sorted(failed_rows.keys()):
                first_homepage = normalize_homepage_url(failed_rows[row_num].get("homepage_url", ""))
                if first_homepage:
                    break
            login_url = first_homepage or "https://www.xiaohongshu.com/explore"
            print(f"登录/页面准备: {login_url}", flush=True)
            print(f"等待 {login_wait} 秒...", flush=True)
            page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(login_wait * 1000)

        try:
            for index, row_num in enumerate(sorted(failed_rows.keys()), start=1):
                info = failed_rows[row_num]
                title = info["title"]
                homepage_url = info["homepage_url"]
                old_url = info["old_url"]

                print(f"\n[{index}/{len(failed_rows)}] 第 {row_num} 行: {title}", flush=True)

                if index > 1:
                    random_delay(800, 2000)
                
                batch_size = args.batch_size
                batch_interval = args.batch_interval
                if batch_size > 0 and index > 1 and (index - 1) % batch_size == 0:
                    print(f"  - [批次休息] 已处理 {index - 1} 条，休息 {batch_interval} 秒避免风控...", flush=True)
                    time.sleep(batch_interval)
                
                check_verification_popup(page)

                if skip_no_title and ("笔记暂未设置标题" in title or not title):
                    skipped_no_title += 1
                    still_failed += 1
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": "笔记无标题(不可恢复)"})
                    print(f"  - 跳过：笔记无标题，无法匹配 (skip_no_title={skip_no_title})", flush=True)
                    continue

                if not title or not homepage_url:
                    still_failed += 1
                    skipped_empty += 1
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": "标题或主页链接为空"})
                    print(f"  - 跳过：标题或主页链接为空 (title={'有' if title else '空'}, homepage_url={'有' if homepage_url else '空'})", flush=True)
                    continue

                try:
                    note_url = locate_note_url(
                        page,
                        homepage_url,
                        title,
                        max_scrolls=max_scrolls,
                        scroll_pixels=900,
                        detail_wait_ms=detail_wait_ms,
                        require_xsec_token=True,
                        per_item_timeout_sec=per_item_timeout,
                        target_date=target_date,
                    )
                except PlaywrightTimeoutError as exc:
                    note_url = ""
                    timeout_count += 1
                    still_failed += 1
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": f"查找超时: {exc}"})
                    print(f"  - 查找超时: {exc}", flush=True)
                    continue
                except Exception as exc:
                    note_url = ""
                    error_count += 1
                    still_failed += 1
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": f"查找失败: {exc}"})
                    print(f"  - 查找失败: {type(exc).__name__}: {exc}", flush=True)
                    continue

                if not note_url:
                    not_found_count += 1
                    still_failed += 1
                    old_url_has_note_id = bool(extract_note_id(old_url))
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": f"未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}, 原地址={old_url[:80]}, 上次失败原因={info.get('reason', 'N/A')})"})
                    print(f"  - 未找到匹配笔记 (原地址有笔记ID={old_url_has_note_id}, 上次原因={info.get('reason', 'N/A')})", flush=True)
                    continue

                new_note_id = extract_note_id(note_url)
                if not new_note_id:
                    no_note_id_count += 1
                    still_failed += 1
                    failed_items.append({"row": row_num, "title": title, "homepage_url": homepage_url, "old_url": old_url, "reason": f"地址没有笔记ID: {note_url}"})
                    print(f"  - 地址没有笔记ID: {note_url}", flush=True)
                    continue

                updated += 1
                print(f"  - 已找到: {note_url}", flush=True)

                with open(result_csv_path, "a", encoding="utf-8-sig", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow([row_num, note_url])

        finally:
            context.close()

    print(f"\n重试完成。更新: {updated}，仍然失败: {still_failed}", flush=True)
    print(f"失败分类: 跳过(空数据)={skipped_empty}, 跳过(无标题)={skipped_no_title}, 超时={timeout_count}, 异常={error_count}, 未找到={not_found_count}, 无笔记ID={no_note_id_count}", flush=True)

    if failed_items:
        from fix_xhs_note_urls import export_failed_items
        failed_path = export_failed_items(failed_items, result_csv_path)
        if failed_path:
            print(f"仍然失败的数据已导出: {failed_path}", flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="重试失败数据")
    parser.add_argument("--result-csv", required=True, help="结果CSV路径")
    parser.add_argument("--failed-csv", required=True, help="失败数据CSV路径")
    parser.add_argument("--max-scrolls", type=int, default=15)
    parser.add_argument("--per-item-timeout", type=int, default=20)
    parser.add_argument("--detail-wait-ms", type=int, default=8000)
    parser.add_argument("--login-wait", type=int, default=20)
    parser.add_argument("--user-data-dir", default=DEFAULT_USER_DATA_DIR, help="浏览器会话目录")
    parser.add_argument("--parallel-index", type=int, default=-1, help="并行任务索引（0开始），-1表示非并行模式")
    parser.add_argument("--parallel-session", default="", help="并行会话ID，用于区分浏览器目录")
    parser.add_argument("--skip-no-title", action="store_true", default=True, help="跳过无标题笔记（默认开启）")
    parser.add_argument("--no-skip-no-title", action="store_true", help="不跳过无标题笔记")
    parser.add_argument("--target-date", default="", help="目标日期（格式：YYYY-MM-DD 或 MM-DD），用于提前停止滚动")
    parser.add_argument("--batch-size", type=int, default=30, help="每处理多少条后休息一次，0表示不休息，默认30")
    parser.add_argument("--batch-interval", type=int, default=30, help="批次休息秒数，默认30")
    parser.add_argument("--sessions", default="", help="多个session目录，用逗号分隔，如 session1,session2,session3")
    parser.add_argument("--session-mode", choices=["rotate", "bind"], default="rotate", help="账号模式: rotate=所有任务共享所有账号, bind=每个任务使用分配到的账号组")
    args = parser.parse_args()

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

    user_data_dir = args.user_data_dir
    if args.parallel_session:
        user_data_dir = os.path.join(args.user_data_dir, f"parallel_{args.parallel_session}_{args.parallel_index}")

    with open(args.result_csv, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["__row_num__", NOTE_URL_HEADER])

    retry_failed(
        args.result_csv,
        args.failed_csv,
        max_scrolls=args.max_scrolls,
        per_item_timeout=args.per_item_timeout,
        detail_wait_ms=args.detail_wait_ms,
        login_wait=args.login_wait,
        user_data_dir=user_data_dir,
        skip_no_title=skip_no_title,
        target_date=target_date,
    )
