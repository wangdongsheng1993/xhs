#!/usr/bin/env python3
"""
抖音KOL执行表同步脚本
任务1: 从抖音提报-KOL同步到确认执行-抖音-3月/4月/5月
任务2: 从确认执行表同步到机器流转统计3-4月

优化内容:
- 日志记录
- API调用缓存
- 批量写入优化
- 配置文件支持
- 重复代码抽象
"""

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
# 常量配置
# ============================================================
DEFAULT_SPREADSHEET_TOKEN = "PY4ysYkCihuiyttOl3TcnKdpn0g"
SPREADSHEET_TOKEN = DEFAULT_SPREADSHEET_TOKEN

# Sheet配置 - 便于扩展新月份
SHEETS = {
    "抖音提报-KOL": {"sheet_id": "c3cac9", "header_row": 1},
    "确认执行-抖音-3月": {"sheet_id": "6jbST9", "header_row": 1},
    "确认执行-抖音-4月": {"sheet_id": "7sKHAw", "header_row": 2},
    "确认执行-抖音-5月": {"sheet_id": "a12h5b", "header_row": 2},
    "确认执行-抖音-6月": {"sheet_id": "Qd1wQU", "header_row": 2},
    "机器流转统计3-4月": {"sheet_id": "TyVk3J", "header_row": 1},
}

# 任务2源表配置
TASK2_SOURCES = [
    {"sheet_id": "6jbST9", "title": "确认执行-抖音-3月", "header_row": 1},
    {"sheet_id": "7sKHAw", "title": "确认执行-抖音-4月", "header_row": 2},
    {"sheet_id": "a12h5b", "title": "确认执行-抖音-5月", "header_row": 2},
    {"sheet_id": "Qd1wQU", "title": "确认执行-抖音-6月", "header_row": 2},
]

# 任务1字段映射: 目标列名 -> 源列名
TASK1_MAPPING = {
    "KOL/KOC名称": "KOL名称",
    "主页链接": "主页链接",
    "ID": "ID",
    "合作形式": "合作形式",
}

# 任务2字段映射: 目标列名 -> 源列名（支持多个候选）
TASK2_MAPPING = {
    "博主": ["KOL/KOC名称"],
    "燃气类型": ["气源"],
    "地址/收件人/联系电话": ["产品邮寄地址"],
    "是否已发布": ["审核进度"],
    "发布时间": ["发布时间"],
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


def get_sheet_info():
    """获取表格信息（带缓存）"""
    global _sheet_info_cache
    if _sheet_info_cache is not None:
        return _sheet_info_cache
    
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if result:
        _sheet_info_cache = result
    return _sheet_info_cache


def get_header_mapping_cached(sheet_id, header_row=1):
    """读取表头，返回列名到索引的映射（带缓存）"""
    cache_key = f"{sheet_id}_{header_row}"
    if cache_key in _header_mapping_cache:
        return _header_mapping_cache[cache_key]
    
    mapping = get_header_mapping(SPREADSHEET_TOKEN, sheet_id, header_row)
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
    if isinstance(value, dict):
        return value.get("text") or value.get("link") or ""
    if isinstance(value, list):
        urls = []
        for item in value:
            if isinstance(item, dict) and "link" in item:
                urls.append(item["link"])
            elif isinstance(item, dict) and "text" in item:
                urls.append(item["text"])
        if urls:
            return ", ".join(urls)
        return str(value)
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
    """执行lark-cli命令"""
    logger.debug(f"执行命令: {cmd}")
    result = subprocess.run(cmd, shell=True, capture_output=True, encoding="utf-8")
    if result.returncode != 0:
        logger.error(f"命令执行失败: {cmd}")
        logger.error(f"错误: {result.stderr}")
        return None
    try:
        return json.loads(result.stdout)
    except Exception as e:
        logger.error(f"解析JSON失败: {e}")
        return None


def native_sheets_write(token, sheet_id, range_str, values_2d):
    """使用原生API写入数据"""
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
    """设置单元格样式"""
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
    """解析日期"""
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
    """索引转列名"""
    result = ""
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_header_mapping(token, sheet_id, header_row):
    """读取表头映射"""
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
    """读取sheet数据"""
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
    """获取sheet行数"""
    info = get_sheet_info()
    if not info:
        return 1
    sheets = info.get("data", {}).get("sheets", {}).get("sheets", [])
    for sheet in sheets:
        if sheet.get("sheet_id") == sheet_id:
            return sheet.get("grid_properties", {}).get("row_count", 1)
    return 1


def extract_cell_value(row_data, col_idx):
    """提取单元格值"""
    if col_idx is None or col_idx >= len(row_data):
        return ""
    val = row_data[col_idx]
    if val is None:
        return ""
    if isinstance(val, dict):
        return val.get("text") or val.get("link") or ""
    return str(val)


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


def save_config_to_file(config_path, config):
    """保存配置到文件"""
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        logger.info(f"配置已保存到: {config_path}")
        return True
    except Exception as e:
        logger.error(f"保存配置文件失败: {e}")
        return False


# ============================================================
# 通用同步函数
# ============================================================
def sync_source_to_target(source_name, target_name, col_mapping, allow_append=False):
    """
    从源表同步数据到目标表（通用函数）
    
    数据保护：如果目标单元格已有数据，不覆盖
    逐列写入：避免覆盖中间列数据
    """
    source_cfg = SHEETS[source_name]
    target_cfg = SHEETS[target_name]
    
    logger.info(f"同步: {source_name} -> {target_name}")
    print(f"\n{'='*50}")
    print(f"同步: {source_name} -> {target_name}")
    
    source_col_map = get_header_mapping_cached(source_cfg["sheet_id"], source_cfg["header_row"])
    target_col_map = get_header_mapping_cached(target_cfg["sheet_id"], target_cfg["header_row"])

    source_col_indices = {k: source_col_map.get(v) for k, v in col_mapping.items()}
    target_col_indices = {k: target_col_map.get(k) for k in col_mapping}
    
    match_key_target_col = list(col_mapping.keys())[0]
    match_key_target_idx = target_col_indices.get(match_key_target_col)
    match_key_source_idx = source_col_indices.get(match_key_target_col)

    if match_key_target_idx is None or match_key_source_idx is None:
        logger.error(f"未找到匹配键列 '{match_key_target_col}'")
        print(f"  错误: 未找到匹配键列 '{match_key_target_col}'")
        return

    rows = read_sheet_data(SPREADSHEET_TOKEN, source_cfg["sheet_id"], source_cfg["header_row"] + 1)
    source_data = []
    for row in rows:
        key_val = normalize_text(extract_cell_value(row, match_key_source_idx))
        if not key_val:
            continue
        row_data = {k: clean_for_write(extract_cell_value(row, v)) for k, v in source_col_indices.items() if v is not None}
        source_data.append((key_val, row_data))

    logger.info(f"  源表读取到 {len(source_data)} 行")
    print(f"  源表读取到 {len(source_data)} 行")

    target_row_count = get_sheet_row_count(SPREADSHEET_TOKEN, target_cfg["sheet_id"])
    if target_row_count <= target_cfg["header_row"]:
        target_row_count = target_cfg["header_row"] + 200
    
    all_target_cols = [v for v in target_col_indices.values() if v is not None]
    if not all_target_cols:
        logger.error("没有可读取的目标列")
        print(f"  错误: 没有可读取的目标列")
        return
    
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
    print(f"  目标表已有 {len(existing_rows)} 行数据")

    to_update = []
    for key_val, row_data in source_data:
        if key_val not in existing_rows:
            continue
        
        row_num, orig_row = existing_rows[key_val]
        changes = {}
        for target_col, new_val in row_data.items():
            tgt_idx = target_col_indices.get(target_col)
            if tgt_idx is None:
                continue
            cur_val = normalize_text(orig_row[tgt_idx - min_col]) if len(orig_row) > (tgt_idx - min_col) else ""
            
            if cur_val:
                continue
            
            if cur_val != new_val and new_val:
                changes[target_col] = (tgt_idx, new_val)
        
        if changes:
            to_update.append((row_num, changes))

    logger.info(f"  需要更新 {len(to_update)} 行")
    print(f"  需要更新 {len(to_update)} 行")

    updated_count = 0
    for row_num, changes in to_update:
        for col_name, (col_idx, new_val) in changes.items():
            range_str = f"{index_to_col(col_idx)}{row_num}:{index_to_col(col_idx)}{row_num}"
            if native_sheets_write(SPREADSHEET_TOKEN, target_cfg["sheet_id"], range_str, [[new_val]]):
                updated_count += 1

    logger.info(f"  已更新 {updated_count} 个单元格")
    print(f"  已更新 {updated_count} 个单元格")
    
    source_nicknames = set(key_val for key_val, _ in source_data)
    missing_in_source = []
    for nick_norm in existing_rows.keys():
        if nick_norm not in source_nicknames:
            missing_in_source.append(nick_norm)
    
    if missing_in_source:
        logger.warning(f"  {len(missing_in_source)} 个博主在源表中不存在")
        print(f"\n  ⚠ 以下 {len(missing_in_source)} 个博主在源表中不存在，无法同步:")
        for nick in sorted(missing_in_source):
            print(f"    - {nick}")


def sync_to_machine_flow():
    """
    任务2: 从确认执行-抖音-3月/4月/5月同步到机器流转统计3-4月
    
    筛选条件:
    1. 达人量级包含"KOL"
    2. 产品邮寄地址不包含"抠图"或"扣图"
    3. 产品邮寄地址不是"自有"
    
    逐列写入：避免覆盖中间列数据
    """
    logger.info("同步: 确认执行表 -> 机器流转统计3-4月")
    print(f"\n{'='*50}")
    print(f"同步: 确认执行表 -> 机器流转统计3-4月")
    
    target_cfg = SHEETS["机器流转统计3-4月"]
    target_col_map = get_header_mapping_cached(target_cfg["sheet_id"], target_cfg["header_row"])
    
    target_indices = {k: target_col_map.get(k) for k in ["博主", "燃气类型", "地址/收件人/联系电话", "是否已发布", "发布时间"]}

    all_source_data = {}
    for src_cfg in TASK2_SOURCES:
        sid, s_title, h_row = src_cfg["sheet_id"], src_cfg["title"], src_cfg["header_row"]
        s_map = get_header_mapping_cached(sid, h_row)
        
        s_indices = {}
        for tk, sv_list in TASK2_MAPPING.items():
            if isinstance(sv_list, str):
                sv_list = [sv_list]
            idx = None
            for sv in sv_list:
                if sv in s_map:
                    idx = s_map[sv]
                    break
            s_indices[tk] = idx
        
        kol_col = s_map.get("达人量级")
        addr_col = s_map.get("产品邮寄地址")
        s_indices["达人量级"] = kol_col
        s_indices["产品邮寄地址"] = addr_col

        rows = read_sheet_data(SPREADSHEET_TOKEN, sid, h_row + 1)
        for r in rows:
            name = normalize_text(extract_cell_value(r, s_indices.get("博主")))
            if not name:
                continue
            
            kol_level = normalize_text(extract_cell_value(r, kol_col)) if kol_col is not None else ""
            address = normalize_text(extract_cell_value(r, addr_col)) if addr_col is not None else ""
            
            if "KOL" not in kol_level.upper():
                continue
            if "抠图" in address or "扣图" in address or address == "自有":
                continue
            
            data = {}
            for tk, idx in s_indices.items():
                if tk in ["达人量级", "产品邮寄地址"]:
                    continue
                if tk == "是否已发布":
                    raw_pub = normalize_text(extract_cell_value(r, idx))
                    data[tk] = "是" if "已发布" in raw_pub else "否"
                elif tk == "发布时间":
                    data[tk] = normalize_date_str(extract_cell_value(r, idx))
                else:
                    data[tk] = clean_for_write(extract_cell_value(r, idx))
            
            all_source_data[name] = data

    logger.info(f"  源表收集到 {len(all_source_data)} 个博主")
    print(f"  源表收集到 {len(all_source_data)} 个博主")

    min_col = min(i for i in target_indices.values() if i is not None)
    max_col = max(i for i in target_indices.values() if i is not None)
    
    t_rows_res = run_lark_cli(
        f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} '
        f'--sheet-id {target_cfg["sheet_id"]} --range "{target_cfg["sheet_id"]}!{index_to_col(min_col)}{target_cfg["header_row"]}:{index_to_col(max_col)}2000"'
    )
    
    existing = {}
    last_row = target_cfg["header_row"]
    if t_rows_res and t_rows_res.get("ok"):
        vals = t_rows_res.get("data", {}).get("valueRange", {}).get("values", [])
        for i, row in enumerate(vals):
            rn = i + target_cfg["header_row"]
            if rn <= target_cfg["header_row"]:
                continue
            name_idx = target_indices["博主"] - min_col
            name = normalize_text(row[name_idx]) if len(row) > name_idx else ""
            if name:
                existing[name] = (rn, row)
                last_row = rn

    logger.info(f"  目标表已有 {len(existing)} 行数据")
    print(f"  目标表已有 {len(existing)} 行数据")

    to_update = []
    to_append = []
    
    for name, sd in all_source_data.items():
        if name in existing:
            rn, orow = existing[name]
            changes = {}
            for tk, nv in sd.items():
                ti = target_indices.get(tk)
                if ti is None:
                    continue
                cv = normalize_text(orow[ti - min_col]) if len(orow) > (ti - min_col) else ""
                if tk == "发布时间":
                    if nv and nv != cv and nv != "未知":
                        changes[tk] = nv
                elif tk == "是否已发布":
                    if cv != nv:
                        changes[tk] = nv
                elif cv != nv and nv:
                    changes[tk] = nv
            if changes:
                to_update.append((rn, changes, sd))
        else:
            to_append.append(sd)

    logger.info(f"  更新 {len(to_update)} 行，追加 {len(to_append)} 行")
    print(f"  更新 {len(to_update)} 行，追加 {len(to_append)} 行")

    updated_count = 0
    for rn, chgs, _ in to_update:
        for col_name, new_val in chgs.items():
            col_idx = target_indices.get(col_name)
            if col_idx is None:
                continue
            range_str = f"{index_to_col(col_idx)}{rn}:{index_to_col(col_idx)}{rn}"
            if native_sheets_write(SPREADSHEET_TOKEN, target_cfg["sheet_id"], range_str, [[new_val]]):
                updated_count += 1
    
    if to_append:
        batch = []
        for rd in to_append:
            row = []
            for c_idx in range(min_col, max_col + 1):
                col_name = next((k for k, v in target_indices.items() if v == c_idx), "")
                row.append(rd.get(col_name, ""))
            batch.append(row)
        
        for i in range(0, len(batch), 50):
            b = batch[i:i+50]
            native_sheets_write(
                SPREADSHEET_TOKEN, 
                target_cfg['sheet_id'], 
                f"{index_to_col(min_col)}{last_row+1+i}:{index_to_col(max_col)}{last_row+i+len(b)}", 
                b
            )

    logger.info(f"  同步完成: 更新 {updated_count} 个单元格，追加 {len(to_append)} 行")
    print(f"  同步完成: 更新 {updated_count} 个单元格，追加 {len(to_append)} 行")


# ============================================================
# 主函数
# ============================================================
def main():
    global SPREADSHEET_TOKEN
    _configure_console_encoding()
    
    config_file = None
    token_override = None
    
    args = sys.argv[1:]
    filtered_args = []
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--config" and i + 1 < len(args):
            config_file = args[i + 1]
            i += 2
        elif arg == "--token" and i + 1 < len(args):
            token_override = args[i + 1]
            i += 2
        else:
            filtered_args.append(arg)
            i += 1
    
    if token_override:
        SPREADSHEET_TOKEN = token_override
        logger.info(f"使用自定义 SPREADSHEET_TOKEN: {token_override}")
    
    run_task1 = True
    run_task2 = True
    
    if config_file:
        config = load_config_from_file(config_file)
        if config:
            tasks = config.get("tasks", [1, 1])
            run_task1 = bool(tasks[0]) if len(tasks) > 0 else True
            run_task2 = bool(tasks[1]) if len(tasks) > 1 else True
        else:
            sys.exit(1)
    elif filtered_args:
        run_task1 = "1" in filtered_args[0]
        run_task2 = "2" in filtered_args[0]
    
    logger.info("="*60)
    logger.info("任务执行计划:")
    logger.info(f"  任务1 (抖音提报-KOL -> 确认执行表): {'是' if run_task1 else '否'}")
    logger.info(f"  任务2 (确认执行表 -> 机器流转统计): {'是' if run_task2 else '否'}")
    logger.info("="*60)
    
    print("="*60)
    print("任务执行计划:")
    print(f"  任务1 (抖音提报-KOL -> 确认执行表): {'是' if run_task1 else '否'}")
    print(f"  任务2 (确认执行表 -> 机器流转统计): {'是' if run_task2 else '否'}")
    print("="*60)
    
    if run_task1:
        print("\n" + "="*60)
        print("执行任务1: 抖音提报-KOL -> 确认执行表")
        print("="*60)
        logger.info("执行任务1: 抖音提报-KOL -> 确认执行表")
        
        for target_name in ["确认执行-抖音-3月", "确认执行-抖音-4月", "确认执行-抖音-5月", "确认执行-抖音-6月"]:
            sync_source_to_target("抖音提报-KOL", target_name, TASK1_MAPPING)
    
    if run_task2:
        print("\n" + "="*60)
        print("执行任务2: 确认执行表 -> 机器流转统计3-4月")
        print("="*60)
        logger.info("执行任务2: 确认执行表 -> 机器流转统计3-4月")
        sync_to_machine_flow()
    
    print("\n" + "="*60)
    print("所有任务执行完成!")
    print("="*60)
    
    print("\n" + "="*60)
    print("执行汇总")
    print("="*60)
    print(f"任务1 (抖音提报-KOL -> 确认执行表): {'已执行' if run_task1 else '未执行'}")
    print(f"任务2 (确认执行表 -> 机器流转统计): {'已执行' if run_task2 else '未执行'}")
    
    logger.info("所有任务执行完成")
    logger.info(f"任务1: {'已执行' if run_task1 else '未执行'}")
    logger.info(f"任务2: {'已执行' if run_task2 else '未执行'}")


if __name__ == "__main__":
    main()
