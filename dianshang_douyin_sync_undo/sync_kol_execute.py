#!/usr/bin/env python3
"""
小红书KOL执行表同步脚本
从二核-KOL同步到确认执行-KOL-4月/5月，再从确认执行表同步到机器流转统计
支持跨文档读取3月数据
"""

import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, date, timedelta

# 主文档（4-5月及统计表所在文档）
SPREADSHEET_TOKEN = "OQKNwQnKriY2NBkXYc0cwBL5nmb"
# 3月数据文档
SPREADSHEET_TOKEN_MAR = "HsVGwSkx5igeUiksE1Mc1qyun5e"

SHEETS = {
    "二核-KOL": {"sheet_id": "xbVzrU", "header_row": 1},
    "确认执行-KOL-4月": {"sheet_id": "C829uw", "header_row": 2},
    "确认执行-KOL-5月": {"sheet_id": "790USh", "header_row": 2},
    "机器流转统计": {"sheet_id": "keg3mV", "header_row": 1},
}

TASK3_SOURCES = [
    {"token": SPREADSHEET_TOKEN, "sheet_id": "C829uw", "title": "确认执行-KOL-4月", "header_row": 2},
    {"token": SPREADSHEET_TOKEN, "sheet_id": "790USh", "title": "确认执行-KOL-5月", "header_row": 2},
    {"token": SPREADSHEET_TOKEN_MAR, "sheet_id": "NOLPy9", "title": "确认执行-KOL-3月", "header_row": 1},
]

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
    "燃气类型": ["气源", "天然气/液化气"],
    "地址/收件人/联系电话": ["产品邮寄地址", "地址/收件人/联系电话"],
    "是否已发布": "稿件进度",
    "发布时间": ["档期", "发布时间"],
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
    text = normalize_text(text)
    if not text:
        return ""
    if "http" in text and ("cellPosition" in text or "type" in text):
        m = re.search(r"https?://[^\s'\"\]\}]+", text)
        if m:
            return m.group(0)
    if "(Programmatic)" in text:
        text = re.sub(r',?\s*\(Programmatic\)\s+\w+', '', text).strip()
        text = re.sub(r',?\s*Type:\s*\(Programmatic\)[^,]*', '', text).strip()
        text = re.sub(r'\s*,\s*', ', ', text).strip().rstrip(',')
    return text


def _resolve_lark_cli_exe():
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
    result = subprocess.run(cmd, shell=True, capture_output=True, encoding="utf-8")
    if result.returncode != 0:
        return None
    try:
        return json.loads(result.stdout)
    except Exception:
        return None


def native_sheets_write(token, sheet_id, range_str, values_2d):
    full_range = f"{sheet_id}!{range_str}" if "!" not in range_str else range_str
    payload = {
        "valueRange": {
            "range": full_range,
            "values": values_2d,
        }
    }
    path = f"/open-apis/sheets/v2/spreadsheets/{token}/values"
    try:
        exe = _resolve_lark_cli_exe()
        body = json.dumps(payload, ensure_ascii=False)
        r = subprocess.run(
            [exe, "api", "PUT", path, "--data", "-"],
            input=body,
            text=True,
            capture_output=True,
            encoding="utf-8",
            timeout=30,
        )
        if r.returncode != 0:
            return False
        out = (r.stdout or "").strip()
        resp = json.loads(out) if out else {}
        return resp.get("code") == 0 or resp.get("ok") is True
    except Exception:
        return False


def native_sheets_style(token, sheet_id, range_str, style_obj):
    path = f"/open-apis/sheets/v2/spreadsheets/{token}/style"
    payload = {
        "appendStyle": {
            "range": f"{sheet_id}!{range_str}" if "!" not in range_str else range_str,
            "style": style_obj
        }
    }
    try:
        exe = _resolve_lark_cli_exe()
        body = json.dumps(payload, ensure_ascii=False)
        subprocess.run(
            [exe, "api", "POST", path, "--data", "-"],
            input=body,
            text=True,
            capture_output=True,
            encoding="utf-8",
            timeout=30,
        )
        return True
    except Exception:
        return False


def normalize_date_str(raw_val):
    text = normalize_text(raw_val)
    if not text:
        return ""
    try:
        f_val = float(raw_val)
        if 1 <= f_val <= 12.31:
            month = int(f_val)
            day_raw = round((f_val - month) * 100)
            if 1 <= day_raw <= 31:
                return f"{month}.{day_raw:02d}"
    except (ValueError, TypeError):
        pass
    m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})$', text.strip())
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{month}.{day:02d}"
    return text


def parse_sheet_date(val):
    if val is None or val == "" or val == "未知":
        return None
    try:
        if isinstance(val, (int, float)) or re.match(r'^\d+\.?\d*$', str(val).strip()):
            f_val = float(val)
            month = int(f_val)
            day_raw = round((f_val - month) * 100)
            if 1 <= month <= 12 and 1 <= day_raw <= 31:
                return date(datetime.now().year, month, day_raw)
        s_val = str(val).strip()
        m = re.match(r'^(\d{1,2})[.\-/](\d{1,2})$', s_val)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            if 1 <= month <= 12 and 1 <= day <= 31:
                return date(datetime.now().year, month, day)
        return None
    except Exception:
        return None


def index_to_col(n):
    result = ""
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_header_mapping(token, sheet_id, header_row):
    range_str = f"{sheet_id}!A{header_row}:BZ{header_row}"
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {token} '
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
        cell_str = normalize_text(cell)
        if cell_str:
            mapping[cell_str] = idx
    return mapping


def read_sheet_data(token, sheet_id, start_row, row_count=500):
    range_str = f"{sheet_id}!A{start_row}:BZ{start_row + row_count}"
    cmd = (
        f'lark-cli sheets +read --spreadsheet-token {token} '
        f'--sheet-id {sheet_id} --range "{range_str}"'
    )
    result = run_lark_cli(cmd)
    if not result or not result.get("ok"):
        return []
    return result.get("data", {}).get("valueRange", {}).get("values", [])


def get_sheet_row_count(token, sheet_id):
    cmd = f'lark-cli sheets +info --spreadsheet-token {token}'
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


def sync_source_to_target(source_name, target_name, col_mapping, start_row=None, allow_append=True):
    source_cfg = SHEETS[source_name]
    target_cfg = SHEETS[target_name]
    print(f"\n{'='*50}\n同步: {source_name} -> {target_name}")
    source_col_map = get_header_mapping(SPREADSHEET_TOKEN, source_cfg["sheet_id"], source_cfg["header_row"])
    target_col_map = get_header_mapping(SPREADSHEET_TOKEN, target_cfg["sheet_id"], target_cfg["header_row"])

    source_col_indices = {k: source_col_map.get(v) for k, v in col_mapping.items()}
    target_col_indices = {k: target_col_map.get(k) for k in col_mapping}
    
    match_key_target_col = list(col_mapping.keys())[0]
    match_key_target_idx = target_col_indices.get(match_key_target_col)
    match_key_source_idx = source_col_indices.get(match_key_target_col)

    if match_key_target_idx is None or match_key_source_idx is None:
        return

    rows = read_sheet_data(SPREADSHEET_TOKEN, source_cfg["sheet_id"], start_row if start_row else source_cfg["header_row"] + 1)
    source_data = []
    for row in rows:
        key_val = normalize_text(extract_cell_value(row, match_key_source_idx))
        if not key_val: continue
        row_data = {k: clean_for_write(extract_cell_value(row, v)) for k, v in source_col_indices.items() if v is not None}
        source_data.append((key_val, row_data))

    target_row_count = get_sheet_row_count(SPREADSHEET_TOKEN, target_cfg["sheet_id"])
    if target_row_count <= target_cfg["header_row"]: target_row_count = target_cfg["header_row"] + 200
    
    all_target_cols = [v for v in target_col_indices.values() if v is not None]
    min_col, max_col = min(all_target_cols), max(all_target_cols)
    range_str = f"{target_cfg['sheet_id']}!{index_to_col(min_col)}{target_cfg['header_row']}:{index_to_col(max_col)}{target_row_count}"
    result = run_lark_cli(f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_cfg["sheet_id"]} --range "{range_str}"')

    existing_rows = {}
    last_data_row = target_cfg["header_row"]
    if result and result.get("ok"):
        values = result.get("data", {}).get("valueRange", {}).get("values", [])
        for idx, row in enumerate(values):
            row_num = idx + target_cfg["header_row"]
            if row_num <= target_cfg["header_row"]: continue
            col_offset = match_key_target_idx - min_col
            nick_norm = normalize_text(row[col_offset]) if row and len(row) > col_offset else ""
            if nick_norm:
                existing_rows[nick_norm] = (row_num, row)
                last_data_row = row_num

    to_update, to_append = [], []
    protect_sheets = {"确认执行-KOL-4月", "确认执行-KOL-5月"}

    for key_val, row_data in source_data:
        if key_val in existing_rows:
            row_num, orig_row = existing_rows[key_val]
            changes = {}
            for target_col, new_val in row_data.items():
                tgt_idx = target_col_indices.get(target_col)
                if tgt_idx is None: continue
                cur_val = normalize_text(orig_row[tgt_idx - min_col]) if len(orig_row) > (tgt_idx - min_col) else ""
                if target_name in protect_sheets and cur_val: continue
                if cur_val != new_val and new_val: changes[target_col] = (tgt_idx, new_val)
            if changes: to_update.append((row_num, changes))
        elif allow_append: to_append.append(row_data)

    for row_num, changes in to_update:
        change_indices = [v[0] for v in changes.values()]
        r_min, r_max = min(change_indices), max(change_indices)
        _, orig_row_values = existing_rows[next(k for k, v in existing_rows.items() if v[0] == row_num)]
        row_vals = []
        for c_idx in range(r_min, r_max + 1):
            col_name = next((n for n, i in target_col_indices.items() if i == c_idx), None)
            if col_name in changes: row_vals.append(changes[col_name][1])
            else: row_vals.append(orig_row_values[c_idx - min_col] if len(orig_row_values) > (c_idx - min_col) else "")
        native_sheets_write(SPREADSHEET_TOKEN, target_cfg["sheet_id"], f"{index_to_col(r_min)}{row_num}:{index_to_col(r_max)}{row_num}", [row_vals])

    if to_append:
        values = []
        for rd in to_append:
            row = [rd.get(next((n for n, i in target_col_indices.items() if i == ci), "")) for ci in range(min_col, max_col + 1)]
            values.append(row)
        for i in range(0, len(values), 50):
            batch = values[i:i+50]
            native_sheets_write(SPREADSHEET_TOKEN, target_cfg['sheet_id'], f"{index_to_col(min_col)}{last_data_row+1+i}:{index_to_col(max_col)}{last_data_row+i+len(batch)}", batch)


def sync_to_machine_flow():
    print(f"\n{'='*50}\n同步: 确认执行(多文档) -> 机器流转统计")
    target_cfg = SHEETS["机器流转统计"]
    target_col_map = get_header_mapping(SPREADSHEET_TOKEN, target_cfg["sheet_id"], target_cfg["header_row"])
    
    target_indices = {k: target_col_map.get(k) for k in ["博主", "燃气类型", "地址/收件人/联系电话", "是否已发布", "发布时间", "是否需要寄走机器", "流转沟通情况", "流转单号（接收）", "流转单号（寄出）"]}

    all_source_data = {}
    for src_cfg in TASK3_SOURCES:
        token, sid, s_title, h_row = src_cfg["token"], src_cfg["sheet_id"], src_cfg["title"], src_cfg["header_row"]
        s_map = get_header_mapping(token, sid, h_row)
        
        # 建立映射索引
        s_indices = {}
        for tk, sv_list in TASK3_MAPPING.items():
            if isinstance(sv_list, str): sv_list = [sv_list]
            idx = None
            for sv in sv_list:
                if sv in s_map:
                    idx = s_map[sv]
                    break
            s_indices[tk] = idx

        rows = read_sheet_data(token, sid, h_row + 1)
        for r in rows:
            name = normalize_text(extract_cell_value(r, s_indices.get("博主")))
            if not name: continue
            
            data = {}
            for tk, idx in s_indices.items():
                if tk == "是否已发布":
                    raw_pub = normalize_text(extract_cell_value(r, idx))
                    data[tk] = "是" if ("已发布" in raw_pub or raw_pub == "是") else "否"
                elif tk == "发布时间":
                    data[tk] = normalize_date_str(extract_cell_value(r, idx))
                else:
                    data[tk] = clean_for_write(extract_cell_value(r, idx))
            
            if "抠图" in data.get("地址/收件人/联系电话", ""): continue
            data["是否需要寄走机器"] = "是"
            all_source_data[name] = data

    min_col = min(i for i in target_indices.values() if i is not None)
    max_col = max(i for i in target_indices.values() if i is not None)
    t_rows_res = run_lark_cli(f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_cfg["sheet_id"]} --range "{target_cfg["sheet_id"]}!{index_to_col(min_col)}{target_cfg["header_row"]}:{index_to_col(max_col)}2000"')
    
    existing = {}
    last_row = target_cfg["header_row"]
    if t_rows_res and t_rows_res.get("ok"):
        vals = t_rows_res.get("data", {}).get("valueRange", {}).get("values", [])
        for i, row in enumerate(vals):
            rn = i + target_cfg["header_row"]
            if rn <= target_cfg["header_row"]: continue
            name_idx = target_indices["博主"] - min_col
            name = normalize_text(row[name_idx]) if len(row) > name_idx else ""
            if name: 
                existing[name] = (rn, row)
                last_row = rn

    today = date.today()
    to_up, to_ap = [], []
    for name, sd in all_source_data.items():
        if name in existing:
            rn, orow = existing[name]
            changes = {}
            for tk, nv in sd.items():
                ti = target_indices.get(tk)
                if ti is None: continue
                cv = normalize_text(orow[ti - min_col]) if len(orow) > (ti - min_col) else ""
                if tk in ["是否已发布", "发布时间"]:
                    if cv != nv: changes[tk] = nv
                elif cv != nv and nv:
                    changes[tk] = nv
            if changes: to_up.append((rn, changes))
        else: to_ap.append(sd)

    # 全量扫描目标表，补全预警状态
    for name, (rn, orow) in existing.items():
        cur_comm = normalize_text(orow[target_indices["流转沟通情况"] - min_col]) if len(orow) > (target_indices["流转沟通情况"] - min_col) else ""
        if cur_comm and cur_comm not in ["需安排流转机器", "需沟通寄走机器"]: continue
        
        existing_up = next((item for item in to_up if item[0] == rn), None)
        if existing_up and "流转沟通情况" in existing_up[1]: continue
            
        cur_pub = normalize_text(orow[target_indices["是否已发布"] - min_col]) if target_indices["是否已发布"] else ""
        canon_pub = "是" if ("已发布" in cur_pub or cur_pub == "是") else "否"
        
        pub_time_str = normalize_text(orow[target_indices["发布时间"] - min_col]) if target_indices.get("发布时间") else ""
        pub_date = parse_sheet_date(pub_time_str)
        need_send = normalize_text(orow[target_indices["是否需要寄走机器"] - min_col]) if target_indices.get("是否需要寄走机器") else ""
        recv_id = normalize_text(orow[target_indices["流转单号（接收）"] - min_col]) if target_indices.get("流转单号（接收）") else ""
        send_id = normalize_text(orow[target_indices["流转单号（寄出）"] - min_col]) if target_indices.get("流转单号（寄出）") else ""
        
        new_comm = ""
        if pub_date and abs((pub_date - today).days) < 10 and not recv_id:
            new_comm = "需安排流转机器"
        elif (canon_pub == "是") and ("是" in need_send or need_send == "是") and recv_id and not send_id:
            new_comm = "需沟通寄走机器"
            
        if new_comm:
            if existing_up: 
                existing_up[1]["流转沟通情况"] = new_comm
                if canon_pub != cur_pub: existing_up[1]["是否已发布"] = canon_pub
            else:
                to_up.append((rn, {"流转沟通情况": new_comm, "是否已发布": canon_pub}))
        elif canon_pub != cur_pub:
            if existing_up: existing_up[1]["是否已发布"] = canon_pub
            else: to_up.append((rn, {"是否已发布": canon_pub}))

    for rn, chgs in to_up:
        _, orig_row_vals = existing[next(k for k,v in existing.items() if v[0] == rn)]
        row_vals = []
        for c_idx in range(min_col, max_col + 1):
            col_name = next((k for k,v in target_indices.items() if v == c_idx), None)
            if col_name in chgs: row_vals.append(chgs[col_name])
            else: row_vals.append(orig_row_vals[c_idx - min_col] if len(orig_row_vals) > (c_idx-min_col) else "")
        
        range_str = f"{index_to_col(min_col)}{rn}:{index_to_col(max_col)}{rn}"
        if native_sheets_write(SPREADSHEET_TOKEN, target_cfg["sheet_id"], range_str, [row_vals]):
            for k, v in chgs.items():
                if k == "流转沟通情况" and v in ["需安排流转机器", "需沟通寄走机器"]:
                    ti = target_indices[k]
                    native_sheets_style(SPREADSHEET_TOKEN, target_cfg["sheet_id"], f"{index_to_col(ti)}{rn}", {"font": {"foreColor": "#FF0000"}})
    
    if to_ap:
        batch = [[rd.get(next((k for k,v in target_indices.items() if v == ci), "")) for ci in range(min_col, max_col + 1)] for rd in to_ap]
        for i in range(0, len(batch), 50):
            b = batch[i:i+50]
            native_sheets_write(SPREADSHEET_TOKEN, target_cfg['sheet_id'], f"{index_to_col(min_col)}{last_row+1+i}:{index_to_col(max_col)}{last_row+i+len(b)}", b)
    print(f"  流转同步完成: 更新 {len(to_up)}, 追加 {len(to_ap)}")


def main():
    _configure_console_encoding()
    run_task1, run_task2, run_task3 = True, True, True
    if len(sys.argv) > 1:
        run_task1 = "1" in sys.argv[1]; run_task2 = "2" in sys.argv[1]; run_task3 = "3" in sys.argv[1]
    
    if run_task1: sync_source_to_target("二核-KOL", "确认执行-KOL-4月", TASK1_MAPPING, allow_append=False)
    if run_task2: sync_source_to_target("二核-KOL", "确认执行-KOL-5月", TASK1_MAPPING, allow_append=False)
    if run_task3: sync_to_machine_flow()

if __name__ == "__main__":
    main()
