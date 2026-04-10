#!/usr/bin/env python3
"""
同步机器流转规划脚本
从4个确认执行sheet读取数据，筛选后更新到目标sheet
"""

import json
import re
import shutil
import subprocess
import sys
from datetime import datetime, timedelta

# 配置
SPREADSHEET_TOKEN = "AJ0owHChoiI6LIk0hi6ctVuCnYQ"


def _configure_console_encoding():
    """
    Windows 下默认控制台编码可能是 gbk，遇到 emoji/特殊字符会导致 print 崩溃。
    这里统一切到 UTF-8，并用 replace 保证不中断主流程。
    """
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        # 保底：不影响主流程
        pass


def normalize_text(value):
    """
    统一清洗飞书单元格文本，避免空格/不可见字符导致匹配失败。
    - 去除 BOM/零宽空格/不间断空格
    - 折叠连续空白
    - strip
    """
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\ufeff", "").replace("\u200b", "").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def format_publish_time_cell(value):
    """
    将「发布时间」列单元格值规范为展示用字符串。

    飞书/表格 API 对形如 4.10 的日期常按数字存储，读出来会变成 float 4.1，写入目标表后显示成 4.1。
    业务上多用「月 + 日/100」录入（如 4 月 10 日 -> 4.10 -> 存成 4.1），此处按该规则还原为 M.DD（如 4.10）。
    对明显为 Excel/表格序列日的较大数值（>20000）不做该解码，原样转字符串。
    """
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return normalize_text(str(value))
    if isinstance(value, (int, float)):
        x = float(value)
        if x > 20000:
            return normalize_text(str(value))
        month = int(x)
        if month < 1 or month > 12:
            return normalize_text(str(value))
        frac = x - month
        if abs(frac) < 1e-9:
            return str(month)
        day = int(round(frac * 100 + 1e-6))
        if day < 1 or day > 31:
            return normalize_text(str(value))
        return f"{month}.{day:02d}"
    return normalize_text(str(value))


def publish_time_score(text):
    """
    发布时间“可靠度”评分：用于避免 4.1 覆盖 4.10 这类信息丢失。
    分数越高越具体。
    """
    t = normalize_text(text)
    if not t or t == "未知":
        return -1

    # 完整年月日
    if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", t):
        return 30

    # 4月10日 / 4月10号
    if re.fullmatch(r"\d{1,2}月\d{1,2}(日|号)", t):
        return 25

    # 4.10 / 4-10 / 4/10
    if re.fullmatch(r"\d{1,2}[./-]\d{2}", t):
        return 20

    # 4.1 / 4-1 / 4/1
    if re.fullmatch(r"\d{1,2}[./-]\d{1}", t):
        return 10

    # 其它非空文本（比如“定稿第二天”）——保留但不用于覆盖更具体日期
    return 0


def should_update_publish_time(current_value, new_value):
    """
    是否应当用 new 覆盖 current。
    每次执行都强制用源表最新值覆盖目标表，确保发布时间始终同步。
    仅在 new 为空/未知 且 current 已有值时不覆盖。
    """
    cur = normalize_text(current_value)
    new = normalize_text(new_value)

    # 完全一致无需更新
    if cur == new:
        return False

    # new 为空/未知：不覆盖已有值
    if not new or new == "未知":
        return False

    # 有新值就更新（不再比较可靠度分数，保证每次执行同步）
    return True

# 源Sheet配置
SOURCE_SHEETS = {
    "小红书确认执行3月": {
        "sheet_id": "ZzaIyS",
        "start_row": None,
        "header_row": 2,
        "platform": "小红书",
        "month": "3月"
    },
    "小红书确认执行4月": {
        "sheet_id": "TDxBlT",
        "start_row": None,
        "header_row": 2,
        "platform": "小红书",
        "month": "4月"
    },
    "3月抖音确认执行": {
        "sheet_id": "f08bd2",
        "start_row": None,
        "header_row": 1,
        "platform": "抖音",
        "month": "3月"
    },
    "4月抖音确认执行": {
        "sheet_id": "iFO2b6",
        "start_row": None,
        "header_row": 2,
        "platform": "抖音",
        "month": "4月"
    }
}

# 目标Sheet配置 - 机器流转规划
# 写入时对应4列: A平台(小红书/抖音) B博主(昵称) C燃气类型(气源) D地址(产品邮寄地址)
TARGET_SHEETS = {
    "3月": {"sheet_id": "98jnHp"},  # 机器流转规划3月 Sheet
    "4月": {"sheet_id": "FkznaW"}  # 机器流转规划4月 Sheet
}


def run_lark_cli(cmd):
    """执行lark-cli命令"""
    result = subprocess.run(cmd, shell=True, capture_output=True)
    if result.returncode != 0:
        print(f"命令执行失败: {cmd}")
        print(f"错误: {result.stderr.decode('utf-8', errors='replace')}")
        return None
    stdout = result.stdout.decode('utf-8', errors='replace')
    return json.loads(stdout)


def _resolve_lark_cli_exe():
    """Windows 上可执行文件常为 lark-cli.cmd，需用 shutil.which 解析，否则 subprocess 报 FileNotFoundError。"""
    for name in ("lark-cli", "lark-cli.cmd"):
        p = shutil.which(name)
        if p:
            return p
    return "lark-cli"


def lark_sheets_write(target_sheet_id, range_str, values_2d):
    """
    调用 lark-cli sheets +write，不使用 shell 拼接命令行。
    避免 --values 的 JSON 中含双引号、换行、反斜杠时，在 Windows cmd 中被截断导致 invalid JSON。
    """
    payload = json.dumps(values_2d, ensure_ascii=False)
    argv = [
        _resolve_lark_cli_exe(),
        "sheets",
        "+write",
        "--spreadsheet-token",
        SPREADSHEET_TOKEN,
        "--sheet-id",
        target_sheet_id,
        "--range",
        range_str,
        "--values",
        payload,
    ]
    return subprocess.run(argv, capture_output=True)


def index_to_col(n):
    """索引转列字母"""
    result = ""
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_header_mapping(sheet_id, header_row=2):
    """读取表头，返回精确列名匹配的索引"""
    range_str = f"{sheet_id}!A{header_row}:BZ{header_row}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return {}

    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return {}

    header_row = values[0]
    mapping = {}

    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        cell_str = str(cell).strip()
        mapping[cell_str] = idx

    return mapping


def get_target_header_mapping(target_sheet_id, header_row=1):
    """
    读取目标表表头，返回列名到索引的映射

    用于写入时根据列名动态匹配目标列位置
    """
    range_str = f"{target_sheet_id}!A{header_row}:BZ{header_row}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return {}

    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return {}

    header_row = values[0]
    mapping = {}

    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        cell_str = str(cell).strip()
        mapping[cell_str] = idx

    return mapping


def read_sheet_data(sheet_id, start_row):
    """读取sheet数据"""
    range_str = f"{sheet_id}!A{start_row}:BZ{start_row + 200}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return []

    return result.get("data", {}).get("valueRange", {}).get("values", [])


def filter_and_transform(rows, col_map, platform, month):
    """
    筛选 KOL 且地址非抠图的数据，并提取目标Sheet需要的列:
    - platform: 平台 (小红书/抖音)
    - nickname: 博主昵称 (源表"昵称"列)
    - gas_type: 燃气类型 (源表"气源"列)
    - address: 地址/收件人/联系电话 (源表"产品邮寄地址"列)
    - publish_time: 发布时间 (源表"发布时间"列，无数据则为空字符串)

    筛选条件:
    1. 达人量级包含"KOL"（不筛选KOC）
    2. 产品邮寄地址不包含"抠图"
    """
    filtered = []

    kol_col = col_map.get("达人量级")
    nickname_col = col_map.get("昵称")
    gas_col = col_map.get("气源")
    addr_col = col_map.get("产品邮寄地址")
    publish_col = col_map.get("发布时间")

    print(f"  列映射: kol={kol_col}, nickname={nickname_col}, gas={gas_col}, addr={addr_col}, publish={publish_col}")

    if kol_col is None or nickname_col is None:
        print(f"  缺少必需列: kol={kol_col}, nickname={nickname_col}")
        return []

    for row in rows:
        if len(row) <= max(kol_col, nickname_col):
            continue

        kol_level = str(row[kol_col]) if row[kol_col] else ""
        address = str(row[addr_col]) if addr_col is not None and row[addr_col] else ""

        if "KOL" not in kol_level.upper():
            continue
        if "抠图" in address:
            continue

        gas_type = str(row[gas_col]) if gas_col is not None and row[gas_col] else ""
        nickname = str(row[nickname_col]) if row[nickname_col] else ""
        # 发布时间：无数据则为空字符串，后续写入时转换为"未知"（数字 4.1 -> 4.10）
        publish_time = (
            format_publish_time_cell(row[publish_col])
            if publish_col is not None and row[publish_col] not in (None, "")
            else ""
        )

        filtered.append({
            "platform": platform,
            "nickname": nickname,
            "gas_type": gas_type,
            "address": address,
            "publish_time": publish_time
        })

    return filtered


def find_published_bloggers(rows, col_map):
    """找出发布状态为"已发布"的博主"""
    kol_col = col_map.get("达人量级")
    nickname_col = col_map.get("昵称")
    publish_col = col_map.get("发布状态")

    if kol_col is None or nickname_col is None or publish_col is None:
        return set()

    published = set()
    for row in rows:
        if len(row) <= max(kol_col, nickname_col, publish_col):
            continue

        kol_level = normalize_text(row[kol_col])
        publish_status = normalize_text(row[publish_col])
        nickname = normalize_text(row[nickname_col])

        if "KOL" not in kol_level.upper():
            continue
        if "已发布" in publish_status:
            if nickname:
                published.add(nickname)

    return published


def get_target_existing_nicknames_and_last_row(target_sheet_id, nickname_col_idx):
    """
    读取目标表「博主」列：返回去重昵称集合，以及该列最后一个非空单元格所在行号（1-based）。

    注意：追加写入必须用「最后非空行号 + 1」作为起始行。
    若误用 len(昵称集合)（唯一昵称个数），当中间有空行、或存在重复昵称时，
    会算出偏小的行号，从而覆盖中间已有数据行。
    """
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        return set(), 1

    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    row_count = 1
    for sheet in sheets:
        if sheet.get("sheet_id") == target_sheet_id:
            row_count = sheet.get("grid_properties", {}).get("row_count", 1)
            break

    if row_count <= 1:
        return set(), 1

    start_col = index_to_col(nickname_col_idx)
    end_col = start_col
    range_str = f"{target_sheet_id}!{start_col}2:{end_col}{row_count}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return set(), 1

    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    existing = set()
    last_data_row = 1
    for i, row in enumerate(values):
        sheet_row = i + 2
        if row and row[0]:
            nickname = str(row[0]).strip()
            if nickname:
                existing.add(nickname)
                last_data_row = sheet_row

    return existing, last_data_row


def get_target_existing_data_by_col(target_sheet_id, nickname_col_idx):
    """
    获取目标sheet已有关键字（博主昵称），用于去重
    """
    existing, _ = get_target_existing_nicknames_and_last_row(target_sheet_id, nickname_col_idx)
    return existing


def get_target_existing_data(target_sheet_id):
    """
    获取目标sheet已有关键字（博主昵称），用于去重

    读取目标表列数据，提取"博主"或"昵称"列构建去重集合
    去重逻辑: 博主昵称已存在于目标表则跳过，避免重复写入同一博主
    """
    # 先获取目标表列名到索引的映射
    col_map = get_target_header_mapping(target_sheet_id)
    if not col_map:
        # 降级方案：按固定B列读取
        col_idx = 1
    else:
        # 优先使用"博主"，其次用"昵称"，降级到B列(索引1)
        col_idx = col_map.get("博主") or col_map.get("昵称") or 1

    return get_target_existing_data_by_col(target_sheet_id, col_idx)


def append_to_target(data_list, target_sheet_id):
    """
    追加/更新数据到目标sheet（upsert模式）

    数据写入"机器流转规划"Sheet的列:
    - 平台: 小红书/抖音
    - 博主: 博主昵称
    - 燃气类型: 燃气类型
    - 地址/收件人/联系电话: 产品邮寄地址
    - 发布时间: 发布时间（无数据显示"未知"）

    写入流程(upsert):
    1. 调用 get_target_header_mapping() 获取目标表列名到索引的映射
    2. 根据列名匹配确定目标列位置
    3. 读取目标表全部数据，构建 昵称->(行号,行数据) 的索引
    4. 遍历待写入数据：
       - 昵称已存在：更新该行字段（发布时间、燃气类型、地址等）
       - 昵称不存在：标记为新增
    5. 先逐行更新已有记录的字段
    6. 再批量追加新记录到最后非空行之后
    7. 调用 lark-cli sheets +write 写入数据
    """
    if not data_list:
        print("没有需要处理的数据")
        return

    # 获取目标表列名到索引的映射，用于动态匹配列位置
    col_map = get_target_header_mapping(target_sheet_id)
    print(f"  目标表列名映射: {col_map}")

    # 定义字段名映射关系（源字段名 -> 目标表可能列名）
    field_to_col = {
        "platform": ["平台"],
        "nickname": ["博主"],
        "gas_type": ["燃气类型"],
        "address": ["地址/收件人/联系电话"],
        "publish_time": ["发布时间"]
    }

    # 根据列名找到目标表中的实际列索引
    target_cols = {}
    for field, possible_names in field_to_col.items():
        col_idx = None
        for name in possible_names:
            if name in col_map:
                col_idx = col_map[name]
                break
        if col_idx is None:
            print(f"  警告: 目标表中未找到字段 {field} (尝试过: {possible_names})")
        target_cols[field] = col_idx

    print(f"  字段到目标列索引映射: {target_cols}")

    # 检查必需字段是否都找到
    required_fields = ["platform", "nickname"]
    for f in required_fields:
        if target_cols.get(f) is None:
            print(f"  错误: 缺少必需字段 {f}，无法写入")
            return

    # 读取目标表全部数据，构建 昵称->(行号, 原始行数据) 的映射（用于upsert更新）
    nickname_col_idx = target_cols.get("nickname")
    if nickname_col_idx is None:
        nickname_col_idx = 1

    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        print("  无法获取表格信息")
        return

    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    row_count = 1
    for sheet in sheets:
        if sheet.get("sheet_id") == target_sheet_id:
            row_count = sheet.get("grid_properties", {}).get("row_count", 1)
            break

    if row_count <= 1:
        row_count = 200

    # 读取目标表所有需要关注的列
    cols_to_read = [target_cols[f] for f in ["platform", "nickname", "gas_type", "address", "publish_time"] if target_cols.get(f) is not None]
    if not cols_to_read:
        print("  错误: 没有可读取的列")
        return

    min_col = min(cols_to_read)
    max_col = max(cols_to_read)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)
    range_str = f"{target_sheet_id}!{start_col_letter}1:{end_col_letter}{row_count}"

    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    existing_rows = {}  # {normalize(nickname): (row_num_1based, [cell_values])}
    last_data_row = 1
    if result and result.get("ok"):
        values = result.get("data", {}).get("valueRange", {}).get("values", [])
        for idx, row in enumerate(values):
            row_num = idx + 1
            if row_num == 1:
                continue  # 跳过表头
            if row and len(row) > nickname_col_idx - min_col:
                nick_val = row[nickname_col_idx - min_col]
                if nick_val:
                    nick_norm = normalize_text(nick_val)
                    if nick_norm:
                        existing_rows[nick_norm] = (row_num, row)
                        last_data_row = row_num

    print(f"  目标表已有 {len(existing_rows)} 个不重复博主，最后数据行: {last_data_row}")

    # 分类：需要更新的 和 需要新增的
    to_update = []   # [(row_num, data_dict)]
    to_append = []   # [data_dict]

    for d in data_list:
        nick_norm = normalize_text(d["nickname"])
        if nick_norm in existing_rows:
            row_num, _ = existing_rows[nick_norm]
            to_update.append((row_num, d))
        else:
            to_append.append(d)

    skip_count = len(data_list) - len(to_append)
    print(f"  更新已有 {len(to_update)} 条，新增 {len(to_append)} 条")

    # ====== 第一部分：更新已存在的行 ======
    update_count = 0
    for row_num, d in to_update:
        updated = False
        # 逐字段检查并更新（只更新有值且与当前值不同的字段）
        for field, col_idx in target_cols.items():
            if col_idx is None:
                continue
            if field == "platform":
                continue  # 平台一般不变

            new_val = ""
            if field == "publish_time":
                new_val = d["publish_time"] if d["publish_time"] else "未知"
            elif field == "gas_type":
                new_val = d["gas_type"]
            elif field == "address":
                new_val = d["address"]

            # 读取当前值
            _, orig_row = existing_rows.get(normalize_text(d["nickname"]), (row_num, []))
            col_offset = col_idx - min_col
            cur_val = normalize_text(orig_row[col_offset]) if orig_row and len(orig_row) > col_offset else ""

            norm_new = normalize_text(new_val)

            # 发布时间使用 should_update_publish_time 判断
            if field == "publish_time":
                if should_update_publish_time(cur_val, norm_new):
                    col_letter = index_to_col(col_idx)
                    range_str = f"{target_sheet_id}!{col_letter}{row_num}:{col_letter}{row_num}"
                    wr_result = lark_sheets_write(target_sheet_id, range_str, [[norm_new]])
                    if wr_result.returncode == 0:
                        resp = json.loads(wr_result.stdout.decode('utf-8', errors='replace'))
                        if resp.get("ok"):
                            update_count += 1
                            updated = True
            else:
                # 其他字段：新值非空且不同时才更新
                if norm_new and norm_new != cur_val:
                    col_letter = index_to_col(col_idx)
                    range_str = f"{target_sheet_id}!{col_letter}{row_num}:{col_letter}{row_num}"
                    wr_result = lark_sheets_write(target_sheet_id, range_str, [[norm_new]])
                    if wr_result.returncode == 0:
                        resp = json.loads(wr_result.stdout.decode('utf-8', errors='replace'))
                        if resp.get("ok"):
                            update_count += 1
                            updated = True

    if to_update:
        print(f"  已更新 {update_count} 个字段值")

    # ====== 第二部分：追加新行 ======
    if not to_append:
        return

    start_row = last_data_row + 1

    cols_to_write = [target_cols[f] for f in ["platform", "nickname", "gas_type", "address", "publish_time"] if target_cols.get(f) is not None]
    if not cols_to_write:
        print("  错误: 没有可写入的列")
        return

    min_col = min(cols_to_write)
    max_col = max(cols_to_write)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)

    values = []
    for d in to_append:
        row = []
        for col_idx in range(min_col, max_col + 1):
            if col_idx == target_cols.get("platform"):
                row.append(d["platform"])
            elif col_idx == target_cols.get("nickname"):
                row.append(d["nickname"])
            elif col_idx == target_cols.get("gas_type"):
                row.append(d["gas_type"])
            elif col_idx == target_cols.get("address"):
                row.append(d["address"])
            elif col_idx == target_cols.get("publish_time"):
                row.append(d["publish_time"] if d["publish_time"] else "未知")
            else:
                row.append("")
        values.append(row)

    range_str = f"{target_sheet_id}!{start_col_letter}{start_row}:{end_col_letter}{start_row + len(values) - 1}"

    print(f"写入 {len(values)} 行到 {target_sheet_id}...")
    result = lark_sheets_write(target_sheet_id, range_str, values)
    stdout = result.stdout.decode('utf-8', errors='replace')
    if result.returncode == 0:
        resp = json.loads(stdout)
        if resp.get("ok"):
            print(f"成功写入 {len(values)} 行")
        else:
            print(f"写入失败: {resp}")
    else:
        print(f"命令失败: {result.stderr.decode('utf-8', errors='replace')}")


def main():
    import sys
    _configure_console_encoding()

    if len(sys.argv) < 5:
        print("用法: python sync_machine_plan.py <小红书3月起始行> <小红书4月起始行> <抖音3月起始行> <抖音4月起始行> [是否更新已发布:0或1] [是否更新KOC状态:0或1] [是否更新确认执行表:0或1]")
        print("示例: python sync_machine_plan.py 3 5 2 8 1 1 1")
        sys.exit(1)

    SOURCE_SHEETS["小红书确认执行3月"]["start_row"] = int(sys.argv[1])
    SOURCE_SHEETS["小红书确认执行4月"]["start_row"] = int(sys.argv[2])
    SOURCE_SHEETS["3月抖音确认执行"]["start_row"] = int(sys.argv[3])
    SOURCE_SHEETS["4月抖音确认执行"]["start_row"] = int(sys.argv[4])

    # 是否更新已发布状态，默认1
    update_published = int(sys.argv[5]) if len(sys.argv) > 5 else 1
    # 是否更新KOC/KOL抠图状态，默认1
    update_koc_status = int(sys.argv[6]) if len(sys.argv) > 6 else 1
    # 是否更新确认执行表的"是否修改"列，默认1
    update_modify_flag = int(sys.argv[7]) if len(sys.argv) > 7 else 1

    data_3月 = []
    data_4月 = []

    for name, config in SOURCE_SHEETS.items():
        print(f"\n处理 {name}...")

        col_map = get_header_mapping(config["sheet_id"], config["header_row"])
        print(f"  可用列: {list(col_map.keys())[:10]}...")

        rows = read_sheet_data(config["sheet_id"], config["start_row"])
        print(f"  读取到 {len(rows)} 行")

        filtered = filter_and_transform(rows, col_map, config["platform"], config["month"])
        print(f"  筛选后 {len(filtered)} 行")

        if config["month"] == "3月":
            data_3月.extend(filtered)
        else:
            data_4月.extend(filtered)

    print(f"\n3月数据共 {len(data_3月)} 行")
    # 追加到机器流转规划3月 Sheet，数据写入 A平台 B博主 C燃气类型 D地址 四列
    append_to_target(data_3月, TARGET_SHEETS["3月"]["sheet_id"])

    print(f"4月数据共 {len(data_4月)} 行")
    # 追加到机器流转规划4月 Sheet，数据写入 A平台 B博主 C燃气类型 D地址 四列
    append_to_target(data_4月, TARGET_SHEETS["4月"]["sheet_id"])

    # 更新已发布状态
    if update_published:
        update_published_status()

    # 更新KOC/KOL抠图状态
    if update_koc_status:
        update_koc_status_columns()

    # 更新确认执行表的"是否修改"列
    if update_modify_flag:
        update_modify_flag_columns()

    # 检查并更新需安排流转机器状态（发布时间<10天且流转沟通情况为空）
    update_luzhu_status()

    print("\n同步完成!")


def update_published_status():
    """更新已发布状态"""
    print("\n" + "="*50)
    print("更新已发布状态...")

    # 收集所有已发布的博主
    published_3月 = set()
    published_4月 = set()
    all_kol_3月 = set()
    all_kol_4月 = set()

    for name, config in SOURCE_SHEETS.items():
        if config["start_row"] is None:
            continue

        col_map = get_header_mapping(config["sheet_id"], config["header_row"])
        rows = read_sheet_data(config["sheet_id"], config["start_row"])
        published = find_published_bloggers(rows, col_map)

        # 同时收集本月所有 KOL 昵称，用于把未发布回填为“否”
        kol_col = col_map.get("达人量级")
        nickname_col = col_map.get("昵称")
        if kol_col is not None and nickname_col is not None:
            for row in rows:
                if len(row) <= max(kol_col, nickname_col):
                    continue
                kol_level = normalize_text(row[kol_col])
                if "KOL" not in kol_level.upper():
                    continue
                nick = normalize_text(row[nickname_col])
                if not nick:
                    continue
                if config["month"] == "3月":
                    all_kol_3月.add(nick)
                else:
                    all_kol_4月.add(nick)

        if config["month"] == "3月":
            published_3月.update(published)
        else:
            published_4月.update(published)

    print(f"  3月已发布博主: {len(published_3月)} 个")
    print(f"  4月已发布博主: {len(published_4月)} 个")

    # 更新3月目标表
    if all_kol_3月:
        update_target_published(published_3月, all_kol_3月, TARGET_SHEETS["3月"]["sheet_id"])

    # 更新4月目标表
    if all_kol_4月:
        update_target_published(published_4月, all_kol_4月, TARGET_SHEETS["4月"]["sheet_id"])


def update_target_published(published_set, all_kol_set, target_sheet_id):
    """
    更新目标表的已发布状态

    根据列名动态匹配"是否已发布"列，支持多种列名格式
    """
    if not all_kol_set:
        return

    # 获取目标表列名映射
    col_map = get_target_header_mapping(target_sheet_id)
    if not col_map:
        print(f"  {target_sheet_id}: 无法获取目标表列名映射")
        return

    # 找到昵称列和是否已发布列的索引
    nickname_col = col_map.get("博主")
    published_col = col_map.get("是否已发布")

    if nickname_col is None:
        print(f"  {target_sheet_id}: 未找到昵称列（尝试过：博主）")
        return
    if published_col is None:
        print(f"  {target_sheet_id}: 未找到是否已发布列（尝试过：是否已发布）")
        return

    print(f"  {target_sheet_id}: 昵称列={nickname_col}, 是否已发布列={published_col}")

    # 获取目标表行数
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        return

    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    row_count = 1
    for sheet in sheets:
        if sheet.get("sheet_id") == target_sheet_id:
            row_count = sheet.get("grid_properties", {}).get("row_count", 1)
            break

    if row_count <= 1:
        return

    # 读取昵称列和是否已发布列数据
    min_col = min(nickname_col, published_col)
    max_col = max(nickname_col, published_col)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)
    range_str = f"{target_sheet_id}!{start_col_letter}2:{end_col_letter}{row_count}"

    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return

    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return

    # 统一规范化昵称集合，避免空格/不可见字符导致匹配失败
    published_norm = {normalize_text(x) for x in published_set if normalize_text(x)}
    all_norm = {normalize_text(x) for x in all_kol_set if normalize_text(x)}

    # 找到需要更新的行（表头是第1行，数据从第2行开始）
    rows_to_update = []
    for idx, row in enumerate(values, start=2):  # 从第2行开始
        # 根据实际列索引计算nickname和published在读取数据中的位置
        nickname_val = None
        published_val = None
        if len(row) > nickname_col - min_col:
            nickname_val = row[nickname_col - min_col]
        if len(row) > published_col - min_col:
            published_val = row[published_col - min_col]
        
        # 只处理“源表里存在的 KOL”，避免误改目标表其它历史行
        nickname_norm = normalize_text(nickname_val)
        if not nickname_norm or nickname_norm not in all_norm:
            continue

        desired = "是" if nickname_norm in published_norm else "否"
        current_val = normalize_text(published_val)
        if current_val != desired:
            rows_to_update.append((idx, desired))

    if not rows_to_update:
        print(f"  {target_sheet_id}: 没有需要更新的已发布状态")
        return

    print(f"  {target_sheet_id}: 更新 {len(rows_to_update)} 行")

    # 逐行更新是否已发布列
    published_col_letter = index_to_col(published_col)
    for row_num, value in rows_to_update:
        range_str = f"{target_sheet_id}!{published_col_letter}{row_num}:{published_col_letter}{row_num}"
        result = lark_sheets_write(target_sheet_id, range_str, [[value]])
        if result.returncode != 0:
            print(f"    行{row_num}更新失败")
        else:
            resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
            if not resp.get("ok"):
                print(f"    行{row_num}更新失败")

    print(f"  {target_sheet_id}: 已更新 {len(rows_to_update)} 行的已发布状态")


def update_koc_status_columns():
    """
    更新状态：
    1. KOC：产品邮寄地址=无需收集，气源=无需收集, 快递状态=无需配送产品, 样机情况=无样机
    2. KOL且地址包含抠图：气源=无需收集, 快递状态=无需配送产品, 样机情况=无样机
    """
    print("\n" + "="*50)
    print("更新KOC/KOL抠图状态...")

    for name, config in SOURCE_SHEETS.items():
        if config["start_row"] is None:
            continue

        print(f"\n处理 {name}...")

        col_map = get_header_mapping(config["sheet_id"], config["header_row"])

        kol_col = col_map.get("达人量级")
        addr_col = col_map.get("产品邮寄地址")
        gas_col = col_map.get("气源")
        express_col = col_map.get("产品快递状态（详情看机器流转规划表）")
        sample_col = col_map.get("样机情况（异常备注好问题）")

        if kol_col is None or addr_col is None:
            print(f"  缺少必需列，跳过")
            continue

        # 读取数据
        rows = read_sheet_data(config["sheet_id"], config["start_row"])
        print(f"  读取到 {len(rows)} 行")

        # 找到需要更新的行
        koc_rows = []  # KOC行
        kol_koutu_rows = []  # KOL但地址包含抠图的行
        for idx, row in enumerate(rows, start=config["start_row"]):
            if len(row) <= max(kol_col, addr_col):
                continue

            kol_level = str(row[kol_col]) if row[kol_col] else ""
            address = str(row[addr_col]) if row[addr_col] else ""

            # KOC
            if "KOC" in kol_level.upper():
                koc_rows.append(idx)
            # KOL且地址包含抠图
            elif "KOL" in kol_level.upper() and "抠图" in address:
                kol_koutu_rows.append(idx)

        print(f"  KOC行: {len(koc_rows)}, KOL抠图行: {len(kol_koutu_rows)}")

        updated_count = 0

        # 更新KOC行（4列）
        for row_num in koc_rows:
            updates = [
                (addr_col, "无需收集"),
                (gas_col, "无需收集"),
                (express_col, "无需配送产品"),
                (sample_col, "无样机")
            ]
            for col_idx, value in updates:
                if col_idx is None:
                    continue
                col_letter = index_to_col(col_idx)
                range_str = f"{config['sheet_id']}!{col_letter}{row_num}:{col_letter}{row_num}"
                result = lark_sheets_write(config["sheet_id"], range_str, [[value]])
                if result.returncode == 0:
                    resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
                    if resp.get("ok"):
                        updated_count += 1

        # 更新KOL抠图行（3列：气源、快递状态、样机情况）
        for row_num in kol_koutu_rows:
            updates = [
                (gas_col, "无需收集"),
                (express_col, "无需配送产品"),
                (sample_col, "无样机")
            ]
            for col_idx, value in updates:
                if col_idx is None:
                    continue
                col_letter = index_to_col(col_idx)
                range_str = f"{config['sheet_id']}!{col_letter}{row_num}:{col_letter}{row_num}"
                result = lark_sheets_write(config["sheet_id"], range_str, [[value]])
                if result.returncode == 0:
                    resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
                    if resp.get("ok"):
                        updated_count += 1

        print(f"  已更新 {updated_count} 个单元格")


def update_modify_flag_columns():
    """
    更新确认执行表的"是否修改"列
    逻辑：KOL且地址非抠图的数据，在确认执行表中将"是否修改"列标记为"是"
    """
    print("\n" + "="*50)
    print("更新确认执行表的'是否修改'列...")

    for name, config in SOURCE_SHEETS.items():
        if config["start_row"] is None:
            continue

        print(f"\n处理 {name}...")

        col_map = get_header_mapping(config["sheet_id"], config["header_row"])

        kol_col = col_map.get("达人量级")
        addr_col = col_map.get("产品邮寄地址")
        modify_col = col_map.get("是否修改")

        if kol_col is None or addr_col is None or modify_col is None:
            print(f"  缺少必需列（达人量级={kol_col}, 产品邮寄地址={addr_col}, 是否修改={modify_col}），跳过")
            continue

        # 读取数据
        rows = read_sheet_data(config["sheet_id"], config["start_row"])
        print(f"  读取到 {len(rows)} 行")

        # 找到需要更新的行：KOL且地址非抠图
        rows_to_update = []
        for idx, row in enumerate(rows, start=config["start_row"]):
            if len(row) <= max(kol_col, addr_col):
                continue

            kol_level = str(row[kol_col]) if row[kol_col] else ""
            address = str(row[addr_col]) if row[addr_col] else ""

            if "KOL" in kol_level.upper() and "抠图" not in address:
                rows_to_update.append(idx)

        print(f"  需更新 {len(rows_to_update)} 行")

        updated_count = 0
        for row_num in rows_to_update:
            col_letter = index_to_col(modify_col)
            range_str = f"{config['sheet_id']}!{col_letter}{row_num}:{col_letter}{row_num}"
            result = lark_sheets_write(config["sheet_id"], range_str, [[1]])
            if result.returncode == 0:
                resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
                if resp.get("ok"):
                    updated_count += 1

        print(f"  已更新 {updated_count} 个单元格")


def update_luzhu_status():
    """
    检查发布时间，若发布时间距今小于10天且流转沟通情况列为空，
    则更新"需安排流转机器"，字体标红

    使用飞书原生API设置单元格样式（红色字体）
    """
    print("\n" + "="*50)
    print("检查并更新需安排流转机器状态...")

    from datetime import datetime

    for name, config in SOURCE_SHEETS.items():
        if config["start_row"] is None:
            continue

        print(f"\n处理 {name}...")

        # 读取源表获取发布时间
        col_map = get_header_mapping(config["sheet_id"], config["header_row"])
        publish_col = col_map.get("发布时间")
        nickname_col = col_map.get("昵称")
        kol_col = col_map.get("达人量级")
        addr_col = col_map.get("产品邮寄地址")

        if publish_col is None:
            print(f"  源表无发布时间列，跳过")
            continue

        rows = read_sheet_data(config["sheet_id"], config["start_row"])

        # 收集需要更新的博主及其发布时间
        # 筛选KOL且地址非抠图的行
        blogger_publish_times = {}  # {normalize(nickname): publish_time_str}
        for idx, row in enumerate(rows):
            if len(row) <= max(kol_col or 0, nickname_col or 0, publish_col or 0, addr_col or 0):
                continue

            kol_level = str(row[kol_col]) if kol_col is not None and row[kol_col] else ""
            address = str(row[addr_col]) if addr_col is not None and row[addr_col] else ""

            if "KOL" not in kol_level.upper():
                continue
            if "抠图" in address:
                continue

            nickname = normalize_text(row[nickname_col]) if nickname_col is not None and row[nickname_col] else ""
            publish_time_str = (
                format_publish_time_cell(row[publish_col])
                if publish_col is not None and row[publish_col] not in (None, "")
                else ""
            )

            if nickname:
                blogger_publish_times[nickname] = publish_time_str

        print(f"  收集到 {len(blogger_publish_times)} 个博主发布时间")

        # 更新对应目标表
        if config["month"] == "3月":
            target_sheet_id = TARGET_SHEETS["3月"]["sheet_id"]
        else:
            target_sheet_id = TARGET_SHEETS["4月"]["sheet_id"]

        update_luzhu_in_target(target_sheet_id, blogger_publish_times)


def update_luzhu_in_target(target_sheet_id, blogger_publish_times):
    """
    更新目标表的流转沟通情况

    Args:
        target_sheet_id: 目标表ID
        blogger_publish_times: {博主昵称: 发布时间字符串}
    """
    if not blogger_publish_times:
        return

    from datetime import datetime

    # 获取目标表列名映射
    col_map = get_target_header_mapping(target_sheet_id)
    if not col_map:
        print(f"  {target_sheet_id}: 无法获取目标表列名映射")
        return

    # 找到昵称列、流转沟通情况列
    nickname_col = col_map.get("博主")
    luzhu_col = col_map.get("流转沟通情况")

    if nickname_col is None:
        print(f"  {target_sheet_id}: 未找到昵称列（博主）")
        return
    if luzhu_col is None:
        print(f"  {target_sheet_id}: 未找到流转沟通情况列")
        return

    print(f"  {target_sheet_id}: 昵称列={nickname_col}, 流转沟通情况列={luzhu_col}")

    # 获取目标表行数
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        return

    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    row_count = 1
    for sheet in sheets:
        if sheet.get("sheet_id") == target_sheet_id:
            row_count = sheet.get("grid_properties", {}).get("row_count", 1)
            break

    if row_count <= 1:
        return

    # 获取所有需要的列索引
    published_col = col_map.get("是否已发布")
    need_send_col = col_map.get("是否需要寄走机器")
    receive_no_col = col_map.get("流转单号（接收）")
    send_no_col = col_map.get("流转单号（寄出）")
    publish_time_col = col_map.get("发布时间")

    # 读取所有需要的列数据
    cols_to_read = [nickname_col, luzhu_col, published_col, need_send_col, receive_no_col, send_no_col, publish_time_col]
    cols_to_read = [c for c in cols_to_read if c is not None]
    if not cols_to_read:
        print(f"  {target_sheet_id}: 没有可读取的列")
        return

    min_col = min(cols_to_read)
    max_col = max(cols_to_read)
    start_col_letter = index_to_col(min_col)
    end_col_letter = index_to_col(max_col)
    range_str = f"{target_sheet_id}!{start_col_letter}1:{end_col_letter}{row_count}"

    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {target_sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)

    if not result or not result.get("ok"):
        return

    values = result.get("data", {}).get("valueRange", {}).get("values", [])
    if not values:
        return

    header_row = values[0] if values else []

    today = datetime.now()
    rows_to_update = []  # [(row_num, col_letter, text)]
    publish_rows_to_update = []  # [(row_num, col_letter, new_value)]

    # 检查每一行
    for idx, row in enumerate(values):
        row_num = idx + 1  # 行号从1开始
        if row_num == 1:
            continue  # 跳过表头

        # 获取昵称
        nickname_val = None
        if len(row) > nickname_col - min_col:
            nickname_val = row[nickname_col - min_col]
        if not nickname_val:
            continue
        nickname = normalize_text(nickname_val)

        # 检查是否在需要更新的博主列表中
        if nickname not in blogger_publish_times:
            continue

        publish_time_str = blogger_publish_times.get(nickname, "")

        # 获取目标表发布时间当前值
        current_publish_time = ""
        if publish_time_col is not None and len(row) > publish_time_col - min_col:
            raw_pt = row[publish_time_col - min_col]
            current_publish_time = (
                format_publish_time_cell(raw_pt)
                if raw_pt not in (None, "")
                else ""
            )
        
        # 计算新的发布时间值（无数据则为"未知"）
        new_publish_time = publish_time_str if publish_time_str else "未知"
        
        # 如果发布时间不一致，添加到更新列表
        if should_update_publish_time(current_publish_time, new_publish_time):
            publish_rows_to_update.append((row_num, index_to_col(publish_time_col), new_publish_time))
            print(f"    博主 {nickname} 发布时间不一致，目标表: '{current_publish_time}' -> 新值: '{new_publish_time}'")

        # 获取流转沟通情况当前值
        luzhu_val = ""
        if len(row) > luzhu_col - min_col:
            luzhu_val = row[luzhu_col - min_col]
        luzhu_val = str(luzhu_val).strip() if luzhu_val else ""

        # 如果流转沟通情况已有内容，跳过
        if luzhu_val:
            continue

        # 检查是否满足"需沟通寄走机器"的条件
        # 条件：是否已发布=是 + 是否需要寄走机器=是 + 流转单号（接收）有数据 + 流转单号（寄出）=空
        is_published = ""
        if published_col is not None and len(row) > published_col - min_col:
            is_published = str(row[published_col - min_col]).strip() if row[published_col - min_col] else ""
        
        need_send = ""
        if need_send_col is not None and len(row) > need_send_col - min_col:
            need_send = str(row[need_send_col - min_col]).strip() if row[need_send_col - min_col] else ""
        
        receive_no = ""
        if receive_no_col is not None and len(row) > receive_no_col - min_col:
            receive_no = str(row[receive_no_col - min_col]).strip() if row[receive_no_col - min_col] else ""
        
        send_no = ""
        if send_no_col is not None and len(row) > send_no_col - min_col:
            send_no = str(row[send_no_col - min_col]).strip() if row[send_no_col - min_col] else ""

        # 检查是否满足"需沟通寄走机器"条件
        if is_published == "是" and need_send == "是" and receive_no and not send_no:
            rows_to_update.append((row_num, index_to_col(luzhu_col), "需沟通寄走机器"))
            print(f"    博主 {nickname} 满足需沟通寄走机器条件")
            continue

        # 检查是否满足"需安排流转机器"条件（发布时间<10天 且 流转单号接收为空）
        if not publish_time_str:
            continue

        # 流转单号（接收）有值时不需要安排流转机器
        if receive_no:
            continue

        try:
            # 尝试解析发布时间（支持多种格式）
            publish_time = None
            for fmt in ["%Y/%m/%d", "%Y-%m-%d", "%m/%d", "%m-%d", "%m.%d", "%m.%d日", "%m月%d日", "%m月%d号"]:
                try:
                    publish_time = datetime.strptime(publish_time_str, fmt)
                    # 如果解析的年份小于1000，加上当前年份
                    if publish_time.year < 1000:
                        publish_time = publish_time.replace(year=today.year)
                    break
                except ValueError:
                    continue

            if publish_time is None:
                print(f"    博主 {nickname} 发布时间格式无法解析: {publish_time_str}")
                continue

            # 检查是否小于10天
            days_diff = (publish_time - today).days
            # 如果发布时间在10天以内（包括负数表示已过期）
            if abs(days_diff) < 10 or days_diff < 10:
                rows_to_update.append((row_num, index_to_col(luzhu_col), "需安排流转机器"))
                print(f"    博主 {nickname} 发布时间 {publish_time_str} 距今 {days_diff} 天，需安排流转机器")
        except Exception as e:
            print(f"    博主 {nickname} 发布时间处理异常: {e}")
            continue

    # 更新发布时间列
    if publish_rows_to_update:
        print(f"  {target_sheet_id}: 需要更新 {len(publish_rows_to_update)} 行的发布时间")
        for row_num, col_letter, new_value in publish_rows_to_update:
            range_str = f"{target_sheet_id}!{col_letter}{row_num}:{col_letter}{row_num}"
            result = lark_sheets_write(target_sheet_id, range_str, [[new_value]])
            if result.returncode != 0:
                print(f"    行{row_num}发布时间写入失败")
                continue

            resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
            if not resp.get("ok"):
                print(f"    行{row_num}发布时间写入失败")
        print(f"  {target_sheet_id}: 已更新 {len(publish_rows_to_update)} 行的发布时间")

    # 更新流转沟通情况列
    if not rows_to_update:
        print(f"  {target_sheet_id}: 没有需要更新流转沟通情况的行")
        return

    print(f"  {target_sheet_id}: 需要更新 {len(rows_to_update)} 行的流转沟通情况")

    # 更新每一行：写入文本并标红
    # 注意：lark-cli sheets +write 不支持设置字体颜色，需要使用原生API
    # 这里先写入文本，然后调用原生API设置红色字体

    for row_num, col_letter, text in rows_to_update:
        # 写入文本
        range_str = f"{target_sheet_id}!{col_letter}{row_num}:{col_letter}{row_num}"
        result = lark_sheets_write(target_sheet_id, range_str, [[text]])
        if result.returncode != 0:
            print(f"    行{row_num}写入失败")
            continue

        resp = json.loads(result.stdout.decode('utf-8', errors='replace'))
        if not resp.get("ok"):
            print(f"    行{row_num}写入失败")
            continue

        # 设置红色字体（使用飞书原生API）
        set_cell_red_font(target_sheet_id, row_num, col_letter)

    print(f"  {target_sheet_id}: 已更新 {len(rows_to_update)} 行的流转沟通情况并标红")


def set_cell_red_font(sheet_id, row, col_letter):
    """
    使用飞书原生API设置单元格字体为红色（加粗+红色）

    Args:
        sheet_id: sheet ID
        row: 行号
        col_letter: 列字母
    """
    import os

    # 优先尝试通过 lark-cli 获取 tenant_access_token
    token = None

    # 方式1：通过 lark-cli 命令获取token
    try:
        exe = _resolve_lark_cli_exe()
        r = subprocess.run(
            [exe, "auth", "token"],
            capture_output=True,
            timeout=10,
        )
        if r.returncode == 0:
            out = r.stdout.decode("utf-8", errors="replace").strip()
            token_data = json.loads(out)
            token = token_data.get("data", {}).get("tenant_access_token") or token_data.get("tenant_access_token")
    except Exception:
        pass

    # 方式2：从配置文件读取
    if not token:
        config_path = os.path.expanduser("~/.lark-cli/config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r") as f:
                    config = json.load(f)
                token = (
                    config.get("access_token")
                    or config.get("tenant_access_token")
                    or config.get("token")
                )
            except Exception:
                pass

    if not token:
        print(f"    无法获取access_token，跳过字体标红")
        return

    # 飞书原生API v3 设置单元格样式（POST）
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/{sheet_id}/style"

    range_str = f"{sheet_id}!{col_letter}{row}:{col_letter}{row}"
    payload = {
        "appendStyle": {
            "range": range_str,
            "style": {
                "font": {
                    "bold": True,
                    "color": "#FF0000"
                }
            }
        }
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    try:
        import urllib.request
        req = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as response:
            result = response.read().decode("utf-8")
            resp_data = json.loads(result)
            code = resp_data.get("code")
            if code == 0:
                print(f"    行{row}字体已标红(加粗红色)")
            else:
                print(f"    行{row}字体标红失败(code={code}): {resp_data.get('msg', 'unknown')}")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"    行{row}字体标红HTTP错误({e.code}): {body[:200]}")
    except Exception as e:
        print(f"    行{row}字体标红异常: {e}")


if __name__ == "__main__":
    main()
