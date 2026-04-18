#!/usr/bin/env python3
"""
小红书KOL执行表同步脚本
从二核-KOL同步到确认执行-KOL-4月/5月，再从确认执行表同步到机器流转统计
支持跨文档读取3月数据

优化版本：
- 日志记录
- API调用缓存
- 批量写入优化（逐列写入）
- 配置文件支持
- 代码抽象
"""

import argparse
import json
import logging
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, date, timedelta

# ============================================================
# 日志配置
# ============================================================
LOG_FILE = "sync_kol_execute.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# 文档配置
# ============================================================
DEFAULT_SPREADSHEET_TOKEN = "HifSwDasHiPCtUkcEoAcfDi6nMe"
DEFAULT_SPREADSHEET_TOKEN_MAR = "D8GlwSmi2iwZaYkZt9mcOcjmn4c"
SPREADSHEET_TOKEN = DEFAULT_SPREADSHEET_TOKEN
SPREADSHEET_TOKEN_MAR = DEFAULT_SPREADSHEET_TOKEN_MAR

# ============================================================
# Sheet配置 - 便于扩展新月份
# ============================================================
SHEETS = {
    "二核-KOL": {"sheet_id": "xbVzrU", "header_row": 1},
    "二核-KOC": {"sheet_id": "2uAcNH", "header_row": 1},
    "确认执行-KOL-4月": {"sheet_id": "C829uw", "header_row": 2},
    "确认执行-KOL-5月": {"sheet_id": "790USh", "header_row": 2},
    "确认执行-KOL-6月": {"sheet_id": "picjM5", "header_row": 2},
    "确认执行-KOC-4月": {"sheet_id": "PcLSfB", "header_row": 2},
    "确认执行-KOC-5月": {"sheet_id": "SOxZPD", "header_row": 2},
    "确认执行-KOC-6月": {"sheet_id": "JEdnGz", "header_row": 2},
    "机器流转统计3-4月": {"sheet_id": "keg3mV", "header_row": 1},
    "机器流转统计5月": {"sheet_id": "hhJ7nE", "header_row": 1},
}

# 月份到目标表的映射配置
MONTH_TO_TARGET = {
    "3月": {"name": "机器流转统计-3月特殊", "token": "SPREADSHEET_TOKEN_MAR", "sheet_id": "4KVKbI", "header_row": 1},
    "4月": {"name": "机器流转统计3-4月", "token": "SPREADSHEET_TOKEN", "sheet_id": "keg3mV", "header_row": 1},
    "5月": {"name": "机器流转统计5月", "token": "SPREADSHEET_TOKEN", "sheet_id": "hhJ7nE", "header_row": 1},
    # "6月": {"name": "机器流转统计6月", "token": "SPREADSHEET_TOKEN", "sheet_id": "xxx", "header_row": 1},
}

# 任务3源表配置（支持跨文档）
TASK3_SOURCES = [
    {"token": SPREADSHEET_TOKEN, "sheet_id": "C829uw", "title": "确认执行-KOL-4月", "header_row": 2},
    {"token": SPREADSHEET_TOKEN, "sheet_id": "790USh", "title": "确认执行-KOL-5月", "header_row": 2},
    {"token": SPREADSHEET_TOKEN, "sheet_id": "picjM5", "title": "确认执行-KOL-6月", "header_row": 2},
    {"token": SPREADSHEET_TOKEN_MAR, "sheet_id": "NOLPy9", "title": "确认执行-KOL-3月", "header_row": 1},
]

# ============================================================
# 字段映射配置
# ============================================================
TASK1_MAPPING = {
    "小红书昵称": "小红书昵称",
    "博主ID": "ID",
    "达人类型": "达人类型",
    "蒲公英链接": "蒲公英链接",
    "主页链接": "主页链接",
    "来源": "来源",
    "形式": "形式",
    "达人平台裸价": "达人平台裸价",
}

TASK4_MAPPING = {
    "小红书昵称": "小红书昵称",
    "博主ID": "ID",
    "蒲公英链接": "蒲公英链接",
    "主页链接": "主页链接",
    "粉丝数 （w）": "粉丝数 （w）",
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

# ============================================================
# API调用缓存
# ============================================================
_sheet_info_cache = None
_header_mapping_cache = {}


def clear_cache():
    """清除所有缓存"""
    global _sheet_info_cache, _header_mapping_cache
    _sheet_info_cache = None
    _header_mapping_cache = {}
    logger.debug("缓存已清除")


def get_header_mapping_cached(token, sheet_id, header_row=1):
    """读取表头，返回列名到索引的映射（带缓存）"""
    cache_key = f"{token}_{sheet_id}_{header_row}"
    if cache_key in _header_mapping_cache:
        return _header_mapping_cache[cache_key]
    
    mapping = get_header_mapping(token, sheet_id, header_row)
    _header_mapping_cache[cache_key] = mapping
    return mapping


# ============================================================
# 工具函数
# ============================================================
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
    if isinstance(value, list):
        parts = []
        for item in value:
            if isinstance(item, dict):
                text = item.get("text", "")
                if text:
                    parts.append(text)
            elif isinstance(item, str):
                parts.append(item)
        return " ".join(parts)
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
            logger.error(f"写入失败: {r.stderr}")
            return False
        out = (r.stdout or "").strip()
        resp = json.loads(out) if out else {}
        return resp.get("code") == 0 or resp.get("ok") is True
    except Exception as e:
        logger.error(f"写入异常: {e}")
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
    except Exception as e:
        logger.error(f"设置样式异常: {e}")
        return False


def normalize_date_str(raw_val):
    """
    标准化日期字符串。
    
    处理飞书 API 返回的各种数据格式：
    - 字符串：直接返回（保留原始格式如 "4.30"）
    - 数字：转换为字符串
    - 字典：提取 text 字段
    - 列表：提取所有 text 字段拼接
    """
    if raw_val is None or raw_val == "":
        return ""
    if isinstance(raw_val, str):
        return raw_val
    if isinstance(raw_val, dict):
        text = raw_val.get("text", "")
        if text:
            return str(text)
        return str(raw_val)
    if isinstance(raw_val, list):
        parts = []
        for item in raw_val:
            if isinstance(item, dict) and "text" in item:
                parts.append(item["text"])
            elif isinstance(item, str):
                parts.append(item)
        if parts:
            return " ".join(parts)
        return str(raw_val)
    if isinstance(raw_val, float):
        if raw_val == int(raw_val):
            return str(int(raw_val))
        return str(raw_val)
    if isinstance(raw_val, int):
        return str(raw_val)
    return normalize_text(str(raw_val))


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
    global _sheet_info_cache
    if _sheet_info_cache is None:
        cmd = f'lark-cli sheets +info --spreadsheet-token {token}'
        _sheet_info_cache = run_lark_cli(cmd)
    
    if not _sheet_info_cache:
        return 1
    sheets = _sheet_info_cache.get("data", {}).get("sheets", {}).get("sheets", [])
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
    if isinstance(val, list):
        parts = []
        for item in val:
            if isinstance(item, dict):
                text = item.get("text", "")
                if text:
                    parts.append(text)
            elif isinstance(item, str):
                parts.append(item)
        return " ".join(parts)
    if isinstance(val, dict):
        return val.get("text") or val.get("link") or ""
    return str(val)


# ============================================================
# 同步函数
# ============================================================
def sync_source_to_target(source_name, target_name, col_mapping, start_row=None, allow_append=True):
    """
    通用同步函数：从源表同步到目标表
    
    Args:
        source_name: 源表名称
        target_name: 目标表名称
        col_mapping: 列映射字典 {目标列名: 源列名}
        start_row: 起始行（可选）
        allow_append: 是否允许追加新行
    """
    source_cfg = SHEETS[source_name]
    target_cfg = SHEETS[target_name]
    
    print(f"\n{'='*50}\n同步: {source_name} -> {target_name}")
    logger.info(f"同步: {source_name} -> {target_name}")
    
    # 使用缓存的表头映射
    source_col_map = get_header_mapping_cached(SPREADSHEET_TOKEN, source_cfg["sheet_id"], source_cfg["header_row"])
    target_col_map = get_header_mapping_cached(SPREADSHEET_TOKEN, target_cfg["sheet_id"], target_cfg["header_row"])

    source_col_indices = {k: source_col_map.get(v) for k, v in col_mapping.items()}
    target_col_indices = {k: target_col_map.get(k) for k in col_mapping}
    
    match_key_target_col = list(col_mapping.keys())[0]
    match_key_target_idx = target_col_indices.get(match_key_target_col)
    match_key_source_idx = source_col_indices.get(match_key_target_col)

    if match_key_target_idx is None or match_key_source_idx is None:
        logger.error(f"无法找到匹配列: {match_key_target_col}")
        return

    # 读取源表数据
    rows = read_sheet_data(SPREADSHEET_TOKEN, source_cfg["sheet_id"], start_row if start_row else source_cfg["header_row"] + 1)
    source_data = []
    for row in rows:
        key_val = normalize_text(extract_cell_value(row, match_key_source_idx))
        if not key_val:
            continue
        row_data = {k: clean_for_write(extract_cell_value(row, v)) for k, v in source_col_indices.items() if v is not None}
        source_data.append((key_val, row_data))
    
    logger.info(f"  源表读取到 {len(source_data)} 行")

    # 读取目标表现有数据
    target_row_count = get_sheet_row_count(SPREADSHEET_TOKEN, target_cfg["sheet_id"])
    if target_row_count <= target_cfg["header_row"]:
        target_row_count = target_cfg["header_row"] + 200
    
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
            if row_num <= target_cfg["header_row"]:
                continue
            col_offset = match_key_target_idx - min_col
            nick_norm = normalize_text(row[col_offset]) if row and len(row) > col_offset else ""
            if nick_norm:
                existing_rows[nick_norm] = (row_num, row)
                last_data_row = row_num
    
    logger.info(f"  目标表已有 {len(existing_rows)} 行数据")

    to_update, to_append = [], []

    for key_val, row_data in source_data:
        if key_val in existing_rows:
            row_num, orig_row = existing_rows[key_val]
            changes = {}
            for target_col, new_val in row_data.items():
                tgt_idx = target_col_indices.get(target_col)
                if tgt_idx is None:
                    continue
                cur_val = normalize_text(orig_row[tgt_idx - min_col]) if len(orig_row) > (tgt_idx - min_col) else ""
                if cur_val != new_val and new_val:
                    changes[target_col] = (tgt_idx, new_val)
            if changes:
                to_update.append((row_num, changes))
        elif allow_append:
            to_append.append(row_data)

    logger.info(f"  需要更新 {len(to_update)} 行")

    # 逐列写入更新（避免覆盖中间列）
    updated_cells = 0
    for row_num, changes in to_update:
        for col_name, (col_idx, new_val) in changes.items():
            range_str = f"{index_to_col(col_idx)}{row_num}:{index_to_col(col_idx)}{row_num}"
            if native_sheets_write(SPREADSHEET_TOKEN, target_cfg["sheet_id"], range_str, [[new_val]]):
                updated_cells += 1
    
    logger.info(f"  已更新 {updated_cells} 个单元格")

    # 批量追加
    if to_append:
        values = []
        for rd in to_append:
            row = [rd.get(next((n for n, i in target_col_indices.items() if i == ci), "")) for ci in range(min_col, max_col + 1)]
            values.append(row)
        for i in range(0, len(values), 50):
            batch = values[i:i+50]
            native_sheets_write(SPREADSHEET_TOKEN, target_cfg['sheet_id'], f"{index_to_col(min_col)}{last_data_row+1+i}:{index_to_col(max_col)}{last_data_row+i+len(batch)}", batch)
        logger.info(f"  已追加 {len(to_append)} 行")


def sync_to_machine_flow():
    """同步确认执行表到机器流转统计（按月份自动分流）"""
    print(f"\n{'='*50}\n同步: 确认执行(多文档) -> 机器流转统计")
    logger.info("同步: 确认执行(多文档) -> 机器流转统计")

    # 收集所有源表数据，记录每个博主来自哪个月份
    blogger_to_month = {}
    all_source_data = {}
    
    for src_cfg in TASK3_SOURCES:
        token, sid, s_title, h_row = src_cfg["token"], src_cfg["sheet_id"], src_cfg["title"], src_cfg["header_row"]
        
        # 从标题提取月份（如"确认执行-KOL-4月" -> "4月"）
        month = None
        for m in ["3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]:
            if m in s_title:
                month = m
                break
        
        s_map = get_header_mapping_cached(token, sid, h_row)
        
        # 建立映射索引
        s_indices = {}
        for tk, sv_list in TASK3_MAPPING.items():
            if isinstance(sv_list, str):
                sv_list = [sv_list]
            idx = None
            for sv in sv_list:
                if sv in s_map:
                    idx = s_map[sv]
                    break
            s_indices[tk] = idx

        rows = read_sheet_data(token, sid, h_row + 1)
        for r in rows:
            name = normalize_text(extract_cell_value(r, s_indices.get("博主")))
            if not name:
                continue
            
            data = {"博主": name}
            for tk, idx in s_indices.items():
                if tk == "博主":
                    continue
                if tk == "是否已发布":
                    raw_pub = normalize_text(extract_cell_value(r, idx))
                    data[tk] = "是" if ("已发布" in raw_pub or raw_pub == "是") else "否"
                elif tk == "发布时间":
                    data[tk] = normalize_date_str(extract_cell_value(r, idx))
                else:
                    data[tk] = clean_for_write(extract_cell_value(r, idx))
            
            # 过滤抠图相关
            if "抠图" in data.get("地址/收件人/联系电话", ""):
                continue
            all_source_data[name] = data
            blogger_to_month[name] = month
    
    logger.info(f"  源表收集到 {len(all_source_data)} 个博主")

    # 按月份分组博主
    month_to_bloggers = {}
    for name, month in blogger_to_month.items():
        if month not in month_to_bloggers:
            month_to_bloggers[month] = []
        month_to_bloggers[month].append(name)

    # 处理每个月份对应的目标表
    total_updated = 0
    total_appended = 0
    
    for month, bloggers in month_to_bloggers.items():
        if month not in MONTH_TO_TARGET:
            logger.warning(f"  月份 {month} 没有配置目标表，跳过 {len(bloggers)} 个博主")
            continue
        
        target_cfg = MONTH_TO_TARGET[month]
        target_name = target_cfg["name"]
        target_token = SPREADSHEET_TOKEN_MAR if target_cfg["token"] == "SPREADSHEET_TOKEN_MAR" else SPREADSHEET_TOKEN
        target_sheet_id = target_cfg["sheet_id"]
        target_header_row = target_cfg["header_row"]
        
        print(f"\n  处理目标表: {target_name} (月份: {month})")
        logger.info(f"  处理目标表: {target_name} (月份: {month})")
        
        target_col_map = get_header_mapping_cached(target_token, target_sheet_id, target_header_row)
        
        # 根据目标表类型确定列映射
        if target_name == "机器流转统计-3月特殊":
            target_indices = {
                "博主": target_col_map.get("博主"),
                "燃气类型": target_col_map.get("燃气类型"),
                "地址/收件人/联系电话": target_col_map.get("地址"),
            }
            is_special_sheet = True
        else:
            target_indices = {k: target_col_map.get(k) for k in ["博主", "燃气类型", "地址/收件人/联系电话", "是否已发布", "发布时间", "流转沟通情况", "流转单号（接收）", "流转单号（寄出）"]}
            is_special_sheet = False
        need_send_col_idx = target_col_map.get("是否需要寄走机器")

        # 过滤出属于该目标表的源数据
        target_source_data = {k: v for k, v in all_source_data.items() if k in bloggers}
        
        if not target_source_data:
            logger.info(f"    无需处理的博主，跳过")
            continue

        # 检查是否有有效的列索引
        valid_indices = [i for i in target_indices.values() if i is not None]
        if not valid_indices:
            logger.warning(f"    目标表没有匹配的列，跳过")
            continue
            
        # 读取目标表现有数据
        min_col = min(valid_indices)
        max_col = max(valid_indices)
        t_rows_res = run_lark_cli(f'lark-cli sheets +read --spreadsheet-token {target_token} --sheet-id {target_sheet_id} --range "{target_sheet_id}!{index_to_col(min_col)}{target_header_row}:{index_to_col(max_col)}2000"')
        
        existing = {}
        last_row = target_header_row
        if t_rows_res and t_rows_res.get("ok"):
            vals = t_rows_res.get("data", {}).get("valueRange", {}).get("values", [])
            for i, row in enumerate(vals):
                rn = i + target_header_row
                if rn <= target_header_row:
                    continue
                name_idx = target_indices["博主"] - min_col
                name = normalize_text(row[name_idx]) if len(row) > name_idx else ""
                if name: 
                    existing[name] = (rn, row)
                    last_row = rn
        
        logger.info(f"    目标表已有 {len(existing)} 行数据")

        today = date.today()
        to_up, to_ap = [], []
        
        # 处理源数据
        for name, sd in target_source_data.items():
            if name in existing:
                rn, orow = existing[name]
                changes = {}
                for tk, nv in sd.items():
                    if tk == "博主":
                        continue
                    ti = target_indices.get(tk)
                    if ti is None:
                        continue
                    cv = normalize_text(orow[ti - min_col]) if len(orow) > (ti - min_col) else ""
                    # 地址列：只要新值不为空就更新（修复何一荷问题）
                    if tk == "地址/收件人/联系电话":
                        if nv and cv != nv:
                            changes[tk] = nv
                    elif tk in ["是否已发布", "发布时间"]:
                        if cv != nv:
                            changes[tk] = nv
                    elif cv != nv and nv:
                        changes[tk] = nv
                if changes:
                    to_up.append((rn, changes, name))
            else:
                to_ap.append((sd, name))

        # 全量扫描目标表，补全预警状态（仅对非特殊sheet）
        if not is_special_sheet:
            for name, (rn, orow) in existing.items():
                cur_comm = normalize_text(orow[target_indices["流转沟通情况"] - min_col]) if len(orow) > (target_indices["流转沟通情况"] - min_col) else ""
                if cur_comm and cur_comm not in ["需安排流转机器", "需沟通寄走机器"]:
                    continue
                
                existing_up = next((item for item in to_up if item[0] == rn), None)
                if existing_up and "流转沟通情况" in existing_up[1]:
                    continue
                    
                cur_pub = normalize_text(orow[target_indices["是否已发布"] - min_col]) if target_indices["是否已发布"] else ""
                canon_pub = "是" if ("已发布" in cur_pub or cur_pub == "是") else "否"
                
                pub_time_str = normalize_text(orow[target_indices["发布时间"] - min_col]) if target_indices.get("发布时间") else ""
                pub_date = parse_sheet_date(pub_time_str)
                need_send = normalize_text(orow[need_send_col_idx - min_col]) if need_send_col_idx is not None else ""
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
                        if canon_pub != cur_pub:
                            existing_up[1]["是否已发布"] = canon_pub
                    else:
                        to_up.append((rn, {"流转沟通情况": new_comm, "是否已发布": canon_pub}, name))
                elif canon_pub != cur_pub:
                    if existing_up:
                        existing_up[1]["是否已发布"] = canon_pub
                    else:
                        to_up.append((rn, {"是否已发布": canon_pub}, name))

        logger.info(f"    更新 {len(to_up)} 行，追加 {len(to_ap)} 行")

        # 逐列写入更新
        for rn, chgs, blogger_name in to_up:
            for col_name, new_val in chgs.items():
                ti = target_indices.get(col_name)
                if ti is None:
                    continue
                range_str = f"{index_to_col(ti)}{rn}:{index_to_col(ti)}{rn}"
                if native_sheets_write(target_token, target_sheet_id, range_str, [[new_val]]):
                    # 设置红色字体
                    if col_name == "流转沟通情况" and new_val in ["需安排流转机器", "需沟通寄走机器"]:
                        native_sheets_style(target_token, target_sheet_id, f"{index_to_col(ti)}{rn}", {"font": {"foreColor": "#FF0000"}})
        
        # 批量追加
        if to_ap:
            batch = []
            for sd, blogger_name in to_ap:
                row = [sd.get(next((k for k, v in target_indices.items() if v == ci), "")) for ci in range(min_col, max_col + 1)]
                batch.append(row)
            for i in range(0, len(batch), 50):
                b = batch[i:i+50]
                native_sheets_write(target_token, target_sheet_id, f"{index_to_col(min_col)}{last_row+1+i}:{index_to_col(max_col)}{last_row+i+len(b)}", b)
        
        total_updated += len(to_up)
        total_appended += len(to_ap)
        logger.info(f"    {target_name} 同步完成: 更新 {len(to_up)} 个单元格，追加 {len(to_ap)} 行")
    
    logger.info(f"  全部同步完成: 总更新 {total_updated} 个单元格，总追加 {total_appended} 行")
    print(f"  流转同步完成: 总更新 {total_updated}, 总追加 {total_appended}")


# ============================================================
# 配置文件支持
# ============================================================
def load_config_from_file(config_path):
    """从配置文件加载参数"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info(f"从配置文件加载参数: {config_path}")
        return config
    except Exception as e:
        logger.error(f"加载配置文件失败: {e}")
        return None


def parse_tasks_param(tasks_param):
    """解析任务参数
    
    Args:
        tasks_param: 可以是字符串"123"或列表[1,1,1]
    
    Returns:
        dict: 任务执行状态字典
    """
    tasks = {
        "kol": True,  # KOL任务：同步到确认执行-KOL-4月/5月/6月
        "koc": True,  # KOC任务：同步到确认执行-KOC-4月/5月/6月
        "flow": True, # 流转任务：确认执行表 -> 机器流转统计
    }
    if isinstance(tasks_param, str):
        tasks["kol"] = "1" in tasks_param
        tasks["koc"] = "2" in tasks_param
        tasks["flow"] = "3" in tasks_param
    elif isinstance(tasks_param, list) and len(tasks_param) >= 3:
        tasks["kol"] = bool(tasks_param[0])
        tasks["koc"] = bool(tasks_param[1])
        tasks["flow"] = bool(tasks_param[2])
    return tasks


# ============================================================
# 主函数
# ============================================================
def main():
    global SPREADSHEET_TOKEN, SPREADSHEET_TOKEN_MAR
    _configure_console_encoding()
    
    parser = argparse.ArgumentParser(description='小红书KOL执行表同步脚本')
    parser.add_argument('tasks', nargs='?', default='123', help='要执行的任务: 1=KOL任务, 2=KOC任务, 3=流转任务')
    parser.add_argument('--config', '-c', help='配置文件路径')
    parser.add_argument('--token', help='自定义 SPREADSHEET_TOKEN')
    parser.add_argument('--token-mar', help='自定义 SPREADSHEET_TOKEN_MAR（3月数据文档）')
    args = parser.parse_args()
    
    if args.token:
        SPREADSHEET_TOKEN = args.token
        logger.info(f"使用自定义 SPREADSHEET_TOKEN: {args.token}")
    
    if args.token_mar:
        SPREADSHEET_TOKEN_MAR = args.token_mar
        logger.info(f"使用自定义 SPREADSHEET_TOKEN_MAR: {args.token_mar}")
    
    if args.config:
        config = load_config_from_file(args.config)
        if config and "tasks" in config:
            tasks = parse_tasks_param(config["tasks"])
        else:
            tasks = parse_tasks_param(args.tasks)
    else:
        tasks = parse_tasks_param(args.tasks)
    
    # 清除缓存
    clear_cache()
    
    # 执行任务
    print("\n" + "="*60)
    print("小红书KOL执行表同步脚本")
    print("="*60)
    
    # KOL任务：同步到确认执行-KOL-4月/5月/6月
    if tasks["kol"]:
        print("\n" + "="*60)
        print("执行KOL任务: 二核-KOL -> 确认执行-KOL-4月/5月/6月")
        print("="*60)
        logger.info("执行KOL任务: 二核-KOL -> 确认执行-KOL-4月/5月/6月")
        sync_source_to_target("二核-KOL", "确认执行-KOL-4月", TASK1_MAPPING, allow_append=False)
        sync_source_to_target("二核-KOL", "确认执行-KOL-5月", TASK1_MAPPING, allow_append=False)
        sync_source_to_target("二核-KOL", "确认执行-KOL-6月", TASK1_MAPPING, allow_append=False)
    
    # KOC任务：同步到确认执行-KOC-4月/5月/6月
    if tasks["koc"]:
        print("\n" + "="*60)
        print("执行KOC任务: 二核-KOC -> 确认执行-KOC-4月/5月/6月")
        print("="*60)
        logger.info("执行KOC任务: 二核-KOC -> 确认执行-KOC-4月/5月/6月")
        sync_source_to_target("二核-KOC", "确认执行-KOC-4月", TASK4_MAPPING, allow_append=False)
        sync_source_to_target("二核-KOC", "确认执行-KOC-5月", TASK4_MAPPING, allow_append=False)
        sync_source_to_target("二核-KOC", "确认执行-KOC-6月", TASK4_MAPPING, allow_append=False)
    
    # 流转任务：确认执行表 -> 机器流转统计
    if tasks["flow"]:
        print("\n" + "="*60)
        print("执行流转任务: 确认执行表 -> 机器流转统计")
        print("="*60)
        logger.info("执行流转任务: 确认执行表 -> 机器流转统计")
        sync_to_machine_flow()
    
    print("\n" + "="*60)
    print("所有任务执行完成!")
    print("="*60)
    
    print("\n" + "="*60)
    print("执行汇总")
    print("="*60)
    print(f"KOL任务 (二核-KOL -> 确认执行-KOL-4月/5月/6月): {'已执行' if tasks['kol'] else '已跳过'}")
    print(f"KOC任务 (二核-KOC -> 确认执行-KOC-4月/5月/6月): {'已执行' if tasks['koc'] else '已跳过'}")
    print(f"流转任务 (确认执行表 -> 机器流转统计): {'已执行' if tasks['flow'] else '已跳过'}")
    
    logger.info("所有任务执行完成")
    for task_name, executed in tasks.items():
        logger.info(f"{task_name}: {'已执行' if executed else '已跳过'}")


if __name__ == "__main__":
    main()
