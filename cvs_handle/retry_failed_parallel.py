import argparse
import atexit
import csv
import os
import shutil
import signal
import subprocess
import sys
import time
from datetime import datetime
from playwright.sync_api import sync_playwright


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RETRY_SCRIPT_PATH = os.path.join(BASE_DIR, "retry_failed.py")
DEFAULT_USER_DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "browser_session")

_processes = []


def _cleanup_processes():
    global _processes
    for p in _processes:
        if p.poll() is None:
            try:
                p.terminate()
                p.wait(timeout=3)
            except Exception:
                try:
                    p.kill()
                except Exception:
                    pass
    _processes = []


def _signal_handler(signum, frame):
    print("\n收到终止信号，正在停止所有进程...", flush=True)
    _cleanup_processes()
    sys.exit(130)


def read_failed_rows(failed_csv_path):
    failed = []
    for enc in ("utf-8-sig", "gbk"):
        try:
            with open(failed_csv_path, "r", encoding=enc) as f:
                reader = csv.DictReader(f)
                for row in reader:
                    row_num = row.get("行号", "")
                    if row_num:
                        failed.append({
                            "row": int(row_num),
                            "title": str(row.get("笔记标题", "") or "").strip(),
                            "homepage_url": str(row.get("主页链接", "") or "").strip(),
                            "old_url": str(row.get("原笔记地址", "") or "").strip(),
                            "reason": str(row.get("失败原因", "") or "").strip(),
                        })
                return failed
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(
        "gbk", b"", 0, 1, f"Could not decode {failed_csv_path} with any known encoding"
    )


def first_homepage_from_failed_rows(failed_rows):
    for item in failed_rows:
        homepage = str(item.get("homepage_url", "") or "").strip()
        if homepage:
            return homepage
    return "https://www.xiaohongshu.com/explore"


def is_login_required(page):
    try:
        return bool(page.evaluate(
            """
            () => {
                const body = document.body ? document.body.innerText || '' : '';
                const modal = document.querySelector('.login-modal, [class*="login-modal"], .login-container');
                return Boolean(
                    modal
                    || body.includes('登录即可查看')
                    || (body.includes('手机号登录') && body.includes('获取验证码'))
                );
            }
            """
        ))
    except Exception:
        return False


def prepare_logged_in_session(user_data_dir, login_url, login_wait):
    wait_seconds = max(int(login_wait or 0), 1)
    print(f"登录状态检查: {login_url}", flush=True)
    print(f"浏览器会话目录: {user_data_dir}", flush=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            slow_mo=80,
            viewport={"width": 1440, "height": 1000},
        )
        page = context.pages[0] if context.pages else context.new_page()
        try:
            page.goto(login_url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(1500)
            if not is_login_required(page):
                print("登录状态可用，准备复制 session。", flush=True)
                return True

            print(f"检测到未登录，请在打开的浏览器里完成登录，最多等待 {wait_seconds} 秒...", flush=True)
            deadline = time.time() + wait_seconds
            while time.time() < deadline:
                page.wait_for_timeout(1000)
                if not is_login_required(page):
                    print("登录完成，准备复制 session。", flush=True)
                    return True

            print("错误: 登录等待超时，未复制未登录 session。请登录后重试，或调大登录等待秒数。", flush=True)
            return False
        finally:
            context.close()


def _read_csv_rows(csv_path, encodings=("utf-8-sig", "gbk")):
    tried = []
    for enc in encodings:
        tried.append(enc)
        try:
            with open(csv_path, "r", encoding=enc, newline="") as f:
                rows = list(csv.reader(f))
            file_size = os.path.getsize(csv_path)
            print(f"  [编码检测] {os.path.basename(csv_path)}: 编码={enc}, 行数={len(rows)}, 文件大小={file_size} bytes (尝试顺序: {tried})", flush=True)
            return rows, enc
        except UnicodeDecodeError as e:
            print(f"  [编码检测] {os.path.basename(csv_path)}: 编码={enc} 解码失败 (byte 0x{e.object[e.start:e.start+1].hex() if e.object else '??'} at pos {e.start}), 尝试下一个...", flush=True)
            continue
    raise UnicodeDecodeError(
        "gbk", b"", 0, 1, f"Could not decode {csv_path} with encodings {tried}"
    )


def merge_retry_results(result_csv_path, temp_csv_paths):
    print(f"[合并] 读取结果CSV: {result_csv_path}", flush=True)
    original_rows, detected_enc = _read_csv_rows(result_csv_path)
    print(f"[合并] 结果CSV: 编码={detected_enc}, 总行数={len(original_rows)} (含表头)", flush=True)

    header_row = original_rows[0]
    print(f"[合并] 表头列: {header_row}", flush=True)
    try:
        url_col_idx = header_row.index("笔记官方地址")
    except ValueError:
        print(f"错误: 找不到列 '笔记官方地址', 当前表头: {header_row}", flush=True)
        return False
    print(f"[合并] '笔记官方地址' 列索引={url_col_idx}", flush=True)

    updated = 0
    skipped_no_url = 0
    skipped_no_token = 0
    skipped_row_out_of_range = 0
    temp_files_found = 0

    for temp_csv_path in temp_csv_paths:
        if not os.path.exists(temp_csv_path):
            print(f"[合并] 临时文件不存在，跳过: {os.path.basename(temp_csv_path)}", flush=True)
            continue
        temp_files_found += 1
        try:
            temp_rows, temp_enc = _read_csv_rows(temp_csv_path)
            print(f"[合并] 临时文件 {os.path.basename(temp_csv_path)}: 编码={temp_enc}, 行数={len(temp_rows)}", flush=True)

            if len(temp_rows) < 2:
                print(f"[合并] 临时文件 {os.path.basename(temp_csv_path)}: 无数据行，跳过", flush=True)
                continue

            temp_header = temp_rows[0]
            print(f"[合并] 临时文件表头: {temp_header}", flush=True)
            try:
                temp_url_col_idx = temp_header.index("笔记官方地址")
                temp_row_col_idx = temp_header.index("__row_num__")
            except ValueError as e:
                print(f"[合并] 临时文件 {os.path.basename(temp_csv_path)}: 缺少必要列 ({e}), 表头={temp_header}", flush=True)
                continue

            temp_updated = 0
            for temp_row in temp_rows[1:]:
                row_num = int(temp_row[temp_row_col_idx])
                note_url = temp_row[temp_url_col_idx]
                if not note_url:
                    skipped_no_url += 1
                    continue
                if "xsec_token=" not in note_url:
                    skipped_no_token += 1
                    print(f"[合并] 行{row_num}: URL缺少xsec_token, url={note_url[:80]}...", flush=True)
                    continue
                if row_num - 1 >= len(original_rows):
                    skipped_row_out_of_range += 1
                    print(f"[合并] 行{row_num}: 超出结果CSV行范围 (最大行号={len(original_rows)})", flush=True)
                    continue
                original_rows[row_num - 1][url_col_idx] = note_url
                temp_updated += 1
            print(f"[合并] 临时文件 {os.path.basename(temp_csv_path)}: 更新了 {temp_updated} 条", flush=True)
            updated += temp_updated
        except Exception as exc:
            print(f"警告: 合并临时文件失败 {os.path.basename(temp_csv_path)}: {exc}", flush=True)

    print(f"[合并] 统计: 找到临时文件={temp_files_found}/{len(temp_csv_paths)}, 成功更新={updated}, URL为空跳过={skipped_no_url}, 缺少token跳过={skipped_no_token}, 行号越界跳过={skipped_row_out_of_range}", flush=True)

    print(f"[合并] 写回结果CSV: {result_csv_path}, 编码={detected_enc}", flush=True)
    with open(result_csv_path, "w", encoding=detected_enc, newline="") as f:
        writer = csv.writer(f)
        writer.writerows(original_rows)

    print(f"合并完成，更新了 {updated} 条记录", flush=True)
    return True


def export_still_failed(failed_items, result_csv_path):
    if not failed_items:
        return ""
    folder = os.path.dirname(result_csv_path)
    name = os.path.splitext(os.path.basename(result_csv_path))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    failed_path = os.path.join(folder, f"{name}_仍然失败_{timestamp}.csv")
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


def collect_worker_failed_items(temp_csv_paths):
    import glob
    failed_items = []
    failed_paths = []
    print("[失败数据收集] 开始扫描 worker 失败数据文件...", flush=True)
    for temp_csv_path in temp_csv_paths:
        folder = os.path.dirname(temp_csv_path)
        name = os.path.splitext(os.path.basename(temp_csv_path))[0]
        pattern = os.path.join(folder, f"{name}_失败数据_*.csv")
        matched_files = glob.glob(pattern)
        if not matched_files:
            print(f"[失败数据收集] 未找到匹配文件: pattern={pattern}", flush=True)
            continue
        for failed_file in matched_files:
            failed_paths.append(failed_file)
            for enc in ("utf-8-sig", "gbk"):
                try:
                    with open(failed_file, "r", encoding=enc) as f:
                        reader = csv.DictReader(f)
                        count = 0
                        for row in reader:
                            failed_items.append({
                                "row": row.get("行号", ""),
                                "title": row.get("笔记标题", ""),
                                "homepage_url": row.get("主页链接", ""),
                                "old_url": row.get("原笔记地址", ""),
                                "reason": row.get("失败原因", ""),
                            })
                            count += 1
                    print(f"[失败数据收集] {os.path.basename(failed_file)}: 编码={enc}, 条数={count}", flush=True)
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as exc:
                    print(f"[失败数据收集] 读取失败 {os.path.basename(failed_file)}: {exc}", flush=True)
                    break
    print(f"[失败数据收集] 总计: {len(failed_items)} 条仍然失败, 涉及 {len(failed_paths)} 个文件", flush=True)
    return failed_items, failed_paths


def main():
    signal.signal(signal.SIGTERM, _signal_handler)
    signal.signal(signal.SIGINT, _signal_handler)
    atexit.register(_cleanup_processes)

    parser = argparse.ArgumentParser(description="并行重试失败数据")
    parser.add_argument("--result-csv", required=True, help="结果CSV路径")
    parser.add_argument("--failed-csv", required=True, help="失败数据CSV路径")
    parser.add_argument("--workers", type=int, default=5, help="并行进程数，默认 5")
    parser.add_argument("--max-scrolls", type=int, default=15, help="每个主页最多滚动次数")
    parser.add_argument("--per-item-timeout", type=int, default=20, help="每条数据超时秒数")
    parser.add_argument("--detail-wait-ms", type=int, default=8000, help="点击后等待毫秒数")
    parser.add_argument("--login-wait", type=int, default=20, help="登录等待秒数")
    parser.add_argument("--skip-no-title", action="store_true", default=True, help="跳过无标题笔记（默认开启）")
    parser.add_argument("--no-skip-no-title", action="store_true", help="不跳过无标题笔记")
    parser.add_argument("--target-date", default="", help="目标日期（格式：YYYY-MM-DD 或 MM-DD），用于提前停止滚动")
    parser.add_argument("--batch-size", type=int, default=30, help="每处理多少条后休息一次，0表示不休息，默认30")
    parser.add_argument("--batch-interval", type=int, default=30, help="批次休息秒数，默认30")
    parser.add_argument("--sessions", default="", help="多个session目录，用逗号分隔，如 session1,session2,session3")
    parser.add_argument("--session-mode", choices=["rotate", "bind"], default="rotate", help="账号模式: rotate=轮询(每条数据切换账号), bind=绑定(每个任务一个账号)")
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

    result_csv_path = os.path.abspath(args.result_csv)
    failed_csv_path = os.path.abspath(args.failed_csv)

    if not os.path.exists(result_csv_path):
        print(f"错误: 结果CSV不存在: {result_csv_path}", flush=True)
        return 1
    if not os.path.exists(failed_csv_path):
        print(f"错误: 失败数据CSV不存在: {failed_csv_path}", flush=True)
        return 1

    result_csv_size = os.path.getsize(result_csv_path)
    failed_csv_size = os.path.getsize(failed_csv_path)
    print(f"[初始化] 结果CSV: {result_csv_path} ({result_csv_size} bytes)", flush=True)
    print(f"[初始化] 失败数据CSV: {failed_csv_path} ({failed_csv_size} bytes)", flush=True)
    print(f"[初始化] 参数: workers={args.workers}, max_scrolls={args.max_scrolls}, per_item_timeout={args.per_item_timeout}, detail_wait_ms={args.detail_wait_ms}, login_wait={args.login_wait}, skip_no_title={skip_no_title}, target_date={target_date.strftime('%Y-%m-%d') if target_date else '未设置'}", flush=True)

    failed_rows = read_failed_rows(failed_csv_path)
    if not failed_rows:
        print("没有失败数据需要重试。", flush=True)
        return 0

    print(f"读取到 {len(failed_rows)} 条失败数据", flush=True)

    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.dirname(result_csv_path)
    name = os.path.splitext(os.path.basename(result_csv_path))[0]

    import glob as _glob
    for pattern in [
        os.path.join(folder, f"{name}_仍然失败_*.csv"),
        os.path.join(folder, f"{name}_retry_temp_*_失败数据_*.csv"),
        os.path.join(folder, f"{name}_retry_temp_{session_id}_*_失败数据_*.csv"),
    ]:
        for old_file in _glob.glob(pattern):
            try:
                os.remove(old_file)
            except Exception:
                pass

    base_user_data_dir = DEFAULT_USER_DATA_DIR
    login_url = first_homepage_from_failed_rows(failed_rows)
    if not prepare_logged_in_session(base_user_data_dir, login_url, args.login_wait):
        return 1

    print("正在复制已登录 session 到各并行进程...", flush=True)
    for i in range(args.workers):
        target_dir = os.path.join(base_user_data_dir, f"parallel_{session_id}_{i}")
        if not os.path.exists(target_dir):
            try:
                shutil.copytree(base_user_data_dir, target_dir, ignore=shutil.ignore_patterns("parallel_*"))
            except Exception as exc:
                print(f"警告: 复制会话目录失败 {target_dir}: {exc}", flush=True)
    print("登录 session 复制完成。", flush=True)

    chunk_size = (len(failed_rows) + args.workers - 1) // args.workers
    chunks = []
    for i in range(args.workers):
        start = i * chunk_size
        end = min(start + chunk_size, len(failed_rows))
        if start < len(failed_rows):
            chunks.append(failed_rows[start:end])

    temp_csv_paths = []

    for i, chunk in enumerate(chunks):
        temp_csv = os.path.join(folder, f"{name}_retry_temp_{session_id}_{i}.csv")
        temp_failed_csv = os.path.join(folder, f"{name}_retry_temp_{session_id}_{i}_失败数据_{session_id}.csv")
        temp_csv_paths.append(temp_csv)

        with open(temp_failed_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["行号", "笔记标题", "主页链接", "原笔记地址", "失败原因"])
            writer.writeheader()
            for item in chunk:
                writer.writerow({
                    "行号": item["row"],
                    "笔记标题": item["title"],
                    "主页链接": item["homepage_url"],
                    "原笔记地址": item["old_url"],
                    "失败原因": item["reason"],
                })

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    env["PYTHONUNBUFFERED"] = "1"

    processes = []

    for i, temp_csv in enumerate(temp_csv_paths):
        temp_failed_csv = os.path.join(folder, f"{name}_retry_temp_{session_id}_{i}_失败数据_{session_id}.csv")
        command = [
            sys.executable,
            "-u",
            RETRY_SCRIPT_PATH,
            "--result-csv", temp_csv,
            "--failed-csv", temp_failed_csv,
            "--max-scrolls", str(args.max_scrolls),
            "--per-item-timeout", str(args.per_item_timeout),
            "--detail-wait-ms", str(args.detail_wait_ms),
            "--login-wait", str(args.login_wait),
            "--user-data-dir", base_user_data_dir,
            "--parallel-index", str(i),
            "--parallel-session", session_id,
        ]

        if not skip_no_title:
            command.append("--no-skip-no-title")

        if target_date:
            command.extend(["--target-date", target_date.strftime("%Y-%m-%d")])
        
        command.extend(["--batch-size", str(args.batch_size)])
        command.extend(["--batch-interval", str(args.batch_interval)])
        
        if args.sessions:
            command.extend(["--sessions", args.sessions])
        if args.session_mode:
            command.extend(["--session-mode", args.session_mode])

        print(f"[任务 {i + 1}] 启动，处理 {len(chunks[i])} 条数据...", flush=True)

        try:
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0
            p = subprocess.Popen(
                command,
                cwd=BASE_DIR,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
            processes.append(p)
            _processes.append(p)
        except Exception as exc:
            print(f"[任务 {i + 1}] 启动失败: {exc}", flush=True)

        time.sleep(0.3)

    print(f"\n并行重试开始: {len(processes)} 个进程", flush=True)
    print(f"结果CSV: {result_csv_path}", flush=True)
    print(f"会话ID: {session_id}", flush=True)
    print("=" * 60, flush=True)

    try:
        while True:
            all_done = True
            for i, p in enumerate(processes):
                if p.poll() is None:
                    all_done = False
                    try:
                        line = p.stdout.readline()
                        if line:
                            print(f"[任务 {i + 1}] {line.rstrip()}", flush=True)
                    except Exception:
                        pass
                else:
                    remaining = p.stdout.read() if p.stdout else ""
                    if remaining:
                        for line in remaining.split("\n"):
                            if line.strip():
                                print(f"[任务 {i + 1}] {line}", flush=True)

            if all_done:
                break
            time.sleep(0.05)
    except KeyboardInterrupt:
        print("\n用户中断，正在停止所有进程...", flush=True)
        _cleanup_processes()
        return 130

    print("=" * 60, flush=True)
    print("所有任务完成，正在合并结果...", flush=True)

    try:
        merge_retry_results(result_csv_path, temp_csv_paths)

        still_failed_items, worker_failed_paths = collect_worker_failed_items(temp_csv_paths)
        if still_failed_items:
            still_failed_path = export_still_failed(still_failed_items, result_csv_path)
            if still_failed_path:
                print(f"仍然失败的数据已合并导出: {still_failed_path}", flush=True)
            print(f"仍然失败数据总数: {len(still_failed_items)}", flush=True)
    except Exception as exc:
        print(f"合并结果时出错: {exc}", flush=True)
    finally:
        print("[清理] 开始清理临时文件和session目录...", flush=True)
        cleaned_temp = 0
        for temp_path in temp_csv_paths:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                    cleaned_temp += 1
                except Exception as exc:
                    print(f"[清理] 删除临时文件失败 {os.path.basename(temp_path)}: {exc}", flush=True)

        cleaned_failed = 0
        for i in range(len(chunks)):
            temp_failed_csv = os.path.join(folder, f"{name}_retry_temp_{session_id}_{i}_失败数据_{session_id}.csv")
            if os.path.exists(temp_failed_csv):
                try:
                    os.remove(temp_failed_csv)
                    cleaned_failed += 1
                except Exception as exc:
                    print(f"[清理] 删除失败数据文件失败 {os.path.basename(temp_failed_csv)}: {exc}", flush=True)

        cleaned_sessions = 0
        for i in range(args.workers):
            temp_session_dir = os.path.join(base_user_data_dir, f"parallel_{session_id}_{i}")
            if os.path.exists(temp_session_dir):
                try:
                    shutil.rmtree(temp_session_dir)
                    cleaned_sessions += 1
                except Exception as exc:
                    print(f"[清理] 删除session目录失败 {temp_session_dir}: {exc}", flush=True)

        print(f"[清理] 完成: 临时结果文件={cleaned_temp}/{len(temp_csv_paths)}, 失败数据文件={cleaned_failed}/{len(chunks)}, session目录={cleaned_sessions}/{args.workers}", flush=True)

    print(f"\n重试完成，结果文件: {result_csv_path}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
