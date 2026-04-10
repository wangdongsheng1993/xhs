#!/usr/bin/env python3
"""
小红书KOL执行表同步脚本
从二核-KOL同步到确认执行-KOL-4月/5月，再从确认执行表同步到机器流转统计
"""

import json
import os
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.request

SPREADSHEET_TOKEN = "O0U1sPdLqhf3Ibt8RUcciIBPnDe"

SHEETS = {
    "二核-KOL": {"sheet_id": "xbVzrU", "header_row": 1},
    "确认执行-KOL-4月": {"sheet_id": "C829uw", "header_row": 2},
    "确认执行-KOL-5月": {"sheet_id": "790USh", "header_row": 2},
    "机器流转统计": {"sheet_id": "keg3mV", "header_row": 1},
}

TASK1_MAPPING = {
    "小红书昵称": "小红书昵称",
    "博主ID": "ID",
    "蒲公英链接": "蒲公英链接",
    "主页链接": "主页链接",
    "来源": "来源",
    "预计档期": "预计档期",
    "形式": "形式",
    "达人平台裸价": "达人平台裸价",
}

TASK3_MAPPING = {
    "博主": "小红书昵称",
    "燃气类型": "气源",
    "地址/收件人/联系电话": "产品邮寄地址",
    "是否已发布": "稿件进度",
    "发布日期": "档期",
}


def _configure_console_encoding():
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def normalize_text(value):
    if value is None:
        return ""
    if isinstance(value, dict):
        return value.get("text") or value.get("link") or ""
    text = str(value)
    text = text.replace("\ufeff", "").replace("\u200b", "").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_for_write(text):
    """清洗单元格写入值，避免把富文本结构串直接写回。"""
    text = normalize_text(text)
    if not text:
        return ""
    if "http" in text and ("cellPosition" in text or "type" in text):
        m = re.search(r"https?://[^\s'\"\]\}]+", text)
        if m:
            return m.group(0)
    return text


def _resolve_lark_cli_exe():
    # Windows 上优先使用 .cmd，避免命中 lark-cli.ps1 带来的参数转义问题
    if os.name == "nt":
        for name in ("lark-cli.cmd", "lark-cli"):
            p = shutil.which(name)
            if p:
                return p
        return "lark-cli.cmd"

    for name in ("lark-cli", "lark-cli.cmd"):
        p = shutil.which(name)
        if p:
            return p
    return "lark-cli"


def run_lark_cli(cmd):
    result = subprocess.run(cmd, shell=True, capture_output=True)
    if result.returncode != 0:
        print(f"命令执行失败: {cmd}")
        print(f"错误: {result.stderr.decode('utf-8', errors='replace')}")
        return None
    stdout = result.stdout.decode("utf-8", errors="replace")
    return json.loads(stdout)


def native_sheets_write(sheet_id, range_str, values_2d):
    """使用 lark-cli 原生 API 写入单元格。"""
    full_range = f"{sheet_id}!{range_str}" if "!" not in range_str else range_str
    payload = {
        "valueRange": {
            "range": full_range,
            "values": values_2d,
        }
    }
    path = f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    try:
        exe = _resolve_lark_cli_exe()
        r = subprocess.run(
            [exe, "api", "PUT", path, "--data", json.dumps(payload, ensure_ascii=False)],
            capture_output=True,
            timeout=30,
        )
        if r.returncode != 0:
            print(f"  命令失败: {r.stderr.decode('utf-8', errors='replace')[:200]}")
            return False
        out = r.stdout.decode("utf-8", errors="replace").strip()
        resp = json.loads(out) if out else {}
        if resp.get("code") == 0 or resp.get("ok") is True:
            return True
        print(f"  写入失败: {resp.get('msg') or resp.get('error', {}).get('message') or 'unknown'}")
        return False
    except Exception as e:
        print(f"  写入异常: {e}")
        return False


def index_to_col(n):
    result = ""
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_header_mapping(sheet_id, header_row):
    range_str = f"{sheet_id}!A{header_row}:BZ{header_row}"
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} '
        f'--sheet-id {sheet_id} --range "{range_str}"'
    )
    result = run_lark_cli(cmd)
    if not result or not result.get("ok"):
        return {}
    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return {}
    header_row_data = values[0]
    mapping = {}
    for idx, cell in enumerate(header_row_data):
        if cell is None:
            continue
        cell_str = normalize_text(cell)
        if cell_str:
            mapping[cell_str] = idx
    return mapping


def read_sheet_data(sheet_id, start_row, row_count=500):
    range_str = f"{sheet_id}!A{start_row}:BZ{start_row + row_count}"
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} '
        f'--sheet-id {sheet_id} --range "{range_str}"'
    )
    result = run_lark_cli(cmd)
    if not result or not result.get("ok"):
        return []
    return result.get("data", {}).get("valueRange", {}).get("values", [])


def get_sheet_row_count(sheet_id):
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        return 1
    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    for sheet in sheets:
        if sheet.get("sheet_id") == sheet_id:
            return sheet.get("grid_properties", {}).get("row_count", 1)
    return 1


def extract_cell_value(row_data, col_idx):
    if col_idx is None or col_idx >= len(row_data):
        return ""
    val = row_data[col_idx]
    if val is None:
        return ""
    if isinstance(val, dict):
        return val.get("text") or val.get("link") or ""
    return str(val)


def sync_source_to_target(source_name, target_name, col_mapping, start_row=None):
    """
    从源sheet读取数据，按列映射upsert到目标sheet。
    col_mapping: {目标列名: 源列名}
    匹配键: 目标表的第一个映射列（通常是"小红书昵称"）
    """
    source_cfg = SHEETS[source_name]
    target_cfg = SHEETS[target_name]

    print(f"\n{'='*50}")
    print(f"同步: {source_name} -> {target_name}")

    source_col_map = get_header_mapping(source_cfg["sheet_id"], source_cfg["header_row"])
    target_col_map = get_header_mapping(target_cfg["sheet_id"], target_cfg["header_row"])

    print(f"  源表列: {list(source_col_map.keys())[:15]}...")
    print(f"  目标表列: {list(target_col_map.keys())[:15]}...")

    source_col_indices = {}
    for target_col, source_col in col_mapping.items():
        src_idx = source_col_map.get(source_col)
        if src_idx is None:
            print(f"  警告: 源表未找到列 '{source_col}'")
        source_col_indices[target_col] = src_idx

    target_col_indices = {}
    for target_col in col_mapping:
        tgt_idx = target_col_map.get(target_col)
        if tgt_idx is None:
            print(f"  警告: 目标表未找到列 '{target_col}'")
        target_col_indices[target_col] = tgt_idx

    match_key_target_col = list(col_mapping.keys())[0]
    match_key_source_col = col_mapping[match_key_target_col]
    match_key_target_idx = target_col_indices.get(match_key_target_col)
    match_key_source_idx = source_col_indices.get(match_key_target_col)

    if match_key_target_idx is None or match_key_source_idx is None:
        print(f"  错误: 匹配键列 '{match_key_target_col}' 未找到，无法同步")
        return

    data_start_row = start_row if start_row else source_cfg["header_row"] + 1
    rows = read_sheet_data(source_cfg["sheet_id"], data_start_row)
    print(f"  读取源表 {len(rows)} 行")

    source_data = []
    for row in rows:
        key_val = normalize_text(extract_cell_value(row, match_key_source_idx))
        if not key_val:
            continue
        row_data = {}
        for target_col, source_col in col_mapping.items():
            src_idx = source_col_indices.get(target_col)
            if src_idx is not None:
                row_data[target_col] = clean_for_write(extract_cell_value(row, src_idx))
            else:
                row_data[target_col] = ""
        source_data.append((key_val, row_data))

    print(f"  有效数据 {len(source_data)} 行")

    target_row_count = get_sheet_row_count(target_cfg["sheet_id"])
    if target_row_count <= target_cfg["header_row"]:
        target_row_count = target_cfg["header_row"] + 200

    all_target_cols = [v for v in target_col_indices.values() if v is not None]
    if not all_target_cols:
        print("  错误: 没有可写入的列")
        return

    min_col = min(all_target_cols)
    max_col = max(all_target_cols)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)
    range_str = (
        f"{target_cfg['sheet_id']}!"
        f"{start_col_letter}{target_cfg['header_row']}:{end_col_letter}{target_row_count}"
    )
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} '
        f'--sheet-id {target_cfg["sheet_id"]} --range "{range_str}"'
    )
    result = run_lark_cli(cmd)

    existing_rows = {}
    last_data_row = target_cfg["header_row"]
    if result and result.get("ok"):
        values = result.get("data", {}).get("valueRange", {}).get("values", [])
        for idx, row in enumerate(values):
            row_num = idx + target_cfg["header_row"]
            if row_num <= target_cfg["header_row"]:
                continue
            col_offset = match_key_target_idx - min_col
            if row and len(row) > col_offset:
                nick_val = row[col_offset]
                nick_norm = normalize_text(nick_val)
                if nick_norm:
                    existing_rows[nick_norm] = (row_num, row)
                    last_data_row = row_num

    print(f"  目标表已有 {len(existing_rows)} 个不重复博主，最后数据行: {last_data_row}")

    to_update = []
    to_append = []

    protect_non_empty_targets = {"确认执行-KOL-4月", "确认执行-KOL-5月"}

    for key_val, row_data in source_data:
        if key_val in existing_rows:
            row_num, orig_row = existing_rows[key_val]
            changes = {}
            for target_col, new_val in row_data.items():
                tgt_idx = target_col_indices.get(target_col)
                if tgt_idx is None:
                    continue
                col_offset = tgt_idx - min_col
                cur_val = normalize_text(orig_row[col_offset]) if orig_row and len(orig_row) > col_offset else ""

                # 对确认执行-KOL-4月/5月：目标单元格已有值则不覆盖
                if target_name in protect_non_empty_targets and cur_val:
                    continue

                if cur_val != new_val and new_val:
                    changes[target_col] = new_val
            if changes:
                to_update.append((row_num, changes))
        else:
            to_append.append(row_data)

    print(f"  更新已有 {len(to_update)} 条，新增 {len(to_append)} 条")

    update_count = 0
    for row_num, changes in to_update:
        for target_col, new_val in changes.items():
            tgt_idx = target_col_indices.get(target_col)
            if tgt_idx is None:
                continue
            col_letter = index_to_col(tgt_idx)
            cell_range = f"{col_letter}{row_num}:{col_letter}{row_num}"
            if native_sheets_write(target_cfg["sheet_id"], cell_range, [[new_val]]):
                update_count += 1

    print(f"  已更新 {update_count} 个单元格")

    if to_append:
        append_start = last_data_row + 1
        values = []
        for row_data in to_append:
            row = []
            for col_idx in range(min_col, max_col + 1):
                found = False
                for target_col, tgt_idx in target_col_indices.items():
                    if tgt_idx == col_idx:
                        row.append(row_data.get(target_col, ""))
                        found = True
                        break
                if not found:
                    row.append("")
            values.append(row)

        batch_size = 1
        total_appended = 0
        for batch_start in range(0, len(values), batch_size):
            batch = values[batch_start:batch_start + batch_size]
            batch_row_start = append_start + batch_start
            range_str = f"{start_col_letter}{batch_row_start}:{end_col_letter}{batch_row_start + len(batch) - 1}"
            print(f"  追加第 {batch_start+1}-{batch_start+len(batch)} 行到 {target_name}...")
            if native_sheets_write(target_cfg["sheet_id"], range_str, batch):
                total_appended += len(batch)
        print(f"  共追加 {total_appended} 行到 {target_name}")


def sync_to_machine_flow():
    """从确认执行-KOL-4月和5月同步到机器流转统计"""
    print(f"\n{'='*50}")
    print("同步: 确认执行-KOL-4月 + 确认执行-KOL-5月 -> 机器流转统计")

    target_cfg = SHEETS["机器流转统计"]
    target_col_map = get_header_mapping(target_cfg["sheet_id"], target_cfg["header_row"])

    target_col_indices = {}
    for target_col in TASK3_MAPPING:
        tgt_idx = target_col_map.get(target_col)
        if tgt_idx is None:
            print(f"  警告: 机器流转统计未找到列 '{target_col}'")
        target_col_indices[target_col] = tgt_idx

    need_send_col = target_col_map.get("是否需要寄走机器")

    all_source_data = {}
    for source_name in ["确认执行-KOL-4月", "确认执行-KOL-5月"]:
        source_cfg = SHEETS[source_name]
        source_col_map = get_header_mapping(source_cfg["sheet_id"], source_cfg["header_row"])

        source_col_indices = {}
        for target_col, source_col in TASK3_MAPPING.items():
            src_idx = source_col_map.get(source_col)
            if src_idx is None:
                print(f"  警告: {source_name}未找到列 '{source_col}'")
            source_col_indices[target_col] = src_idx

        data_start_row = source_cfg["header_row"] + 1
        rows = read_sheet_data(source_cfg["sheet_id"], data_start_row)
        print(f"  {source_name}: 读取 {len(rows)} 行")

        for row in rows:
            key_idx = source_col_indices.get("博主")
            if key_idx is None:
                continue
            key_val = normalize_text(extract_cell_value(row, key_idx))
            if not key_val:
                continue

            row_data = {}
            for target_col, source_col in TASK3_MAPPING.items():
                src_idx = source_col_indices.get(target_col)
                if src_idx is not None:
                    raw_val = clean_for_write(extract_cell_value(row, src_idx))
                    if target_col == "是否已发布":
                        raw_val = "是" if "已发布" in raw_val else "否"
                    row_data[target_col] = raw_val
                else:
                    row_data[target_col] = ""

            row_data["是否需要寄走机器"] = "是"
            all_source_data[key_val] = row_data

    print(f"  合并后共 {len(all_source_data)} 个博主")

    target_row_count = get_sheet_row_count(target_cfg["sheet_id"])
    if target_row_count <= target_cfg["header_row"]:
        target_row_count = target_cfg["header_row"] + 200

    all_target_cols = [v for v in target_col_indices.values() if v is not None]
    if need_send_col is not None:
        all_target_cols.append(need_send_col)
    if not all_target_cols:
        print("  错误: 没有可写入的列")
        return

    min_col = min(all_target_cols)
    max_col = max(all_target_cols)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)
    range_str = (
        f"{target_cfg['sheet_id']}!"
        f"{start_col_letter}{target_cfg['header_row']}:{end_col_letter}{target_row_count}"
    )
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} '
        f'--sheet-id {target_cfg["sheet_id"]} --range "{range_str}"'
    )
    result = run_lark_cli(cmd)

    match_key_target_idx = target_col_indices.get("博主")
    if match_key_target_idx is None:
        print("  错误: 机器流转统计未找到'博主'列")
        return

    existing_rows = {}
    last_data_row = target_cfg["header_row"]
    if result and result.get("ok"):
        values = result.get("data", {}).get("valueRange", {}).get("values", [])
        for idx, row in enumerate(values):
            row_num = idx + target_cfg["header_row"]
            if row_num <= target_cfg["header_row"]:
                continue
            col_offset = match_key_target_idx - min_col
            if row and len(row) > col_offset:
                nick_val = row[col_offset]
                nick_norm = normalize_text(nick_val)
                if nick_norm:
                    existing_rows[nick_norm] = (row_num, row)
                    last_data_row = row_num

    print(f"  机器流转统计已有 {len(existing_rows)} 个博主，最后数据行: {last_data_row}")

    to_update = []
    to_append = []

    for key_val, row_data in all_source_data.items():
        if key_val in existing_rows:
            row_num, orig_row = existing_rows[key_val]
            changes = {}
            for target_col, new_val in row_data.items():
                tgt_idx = target_col_indices.get(target_col) or (need_send_col if target_col == "是否需要寄走机器" else None)
                if tgt_idx is None:
                    continue
                col_offset = tgt_idx - min_col
                cur_val = normalize_text(orig_row[col_offset]) if orig_row and len(orig_row) > col_offset else ""
                if cur_val != new_val and new_val:
                    changes[target_col] = (tgt_idx, new_val)
            if changes:
                to_update.append((row_num, changes))
        else:
            to_append.append(row_data)

    print(f"  更新已有 {len(to_update)} 条，新增 {len(to_append)} 条")

    update_count = 0
    for row_num, changes in to_update:
        for target_col, (tgt_idx, new_val) in changes.items():
            col_letter = index_to_col(tgt_idx)
            cell_range = f"{col_letter}{row_num}:{col_letter}{row_num}"
            if native_sheets_write(target_cfg["sheet_id"], cell_range, [[new_val]]):
                update_count += 1

    print(f"  已更新 {update_count} 个单元格")

    if to_append:
        append_start = last_data_row + 1
        values = []
        for row_data in to_append:
            row = []
            for col_idx in range(min_col, max_col + 1):
                found = False
                for target_col, tgt_idx in target_col_indices.items():
                    if tgt_idx == col_idx:
                        row.append(row_data.get(target_col, ""))
                        found = True
                        break
                if not found and need_send_col is not None and need_send_col == col_idx:
                    row.append(row_data.get("是否需要寄走机器", "是"))
                    found = True
                if not found:
                    row.append("")
            values.append(row)

        batch_size = 1
        total_appended = 0
        for batch_start in range(0, len(values), batch_size):
            batch = values[batch_start:batch_start + batch_size]
            batch_row_start = append_start + batch_start
            range_str = f"{start_col_letter}{batch_row_start}:{end_col_letter}{batch_row_start + len(batch) - 1}"
            print(f"  追加第 {batch_start+1}-{batch_start+len(batch)} 行到机器流转统计...")
            if native_sheets_write(target_cfg["sheet_id"], range_str, batch):
                total_appended += len(batch)
        print(f"  共追加 {total_appended} 行到机器流转统计")


def main():
    _configure_console_encoding()

    run_task1 = True
    run_task2 = True
    run_task3 = True

    if len(sys.argv) > 1:
        run_task1 = "1" in sys.argv[1]
        run_task2 = "2" in sys.argv[1]
        run_task3 = "3" in sys.argv[1]

    print("=" * 50)
    print("小红书KOL执行表同步工具")
    print(f"Spreadsheet: {SPREADSHEET_TOKEN}")
    print(f"任务1(二核-KOL->确认执行-KOL-4月): {'是' if run_task1 else '否'}")
    print(f"任务2(二核-KOL->确认执行-KOL-5月): {'是' if run_task2 else '否'}")
    print(f"任务3(确认执行->机器流转统计): {'是' if run_task3 else '否'}")

    if run_task1:
        sync_source_to_target("二核-KOL", "确认执行-KOL-4月", TASK1_MAPPING)

    if run_task2:
        sync_source_to_target("二核-KOL", "确认执行-KOL-5月", TASK1_MAPPING)

    if run_task3:
        sync_to_machine_flow()

    print("\n同步完成!")


if __name__ == "__main__":
    main()
