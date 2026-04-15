#!/usr/bin/env python3
"""
诊断脚本：检查特定博主的筛选和发布状态
"""

import json
import re
import shutil
import subprocess
import sys

SPREADSHEET_TOKEN = "ScEkwBLZDizM85kvRhpcnO7Lnkr"

def normalize_text(value):
    """统一清洗文本"""
    if value is None:
        return ""
    if isinstance(value, list):
        urls = []
        for item in value:
            if isinstance(item, dict) and "link" in item:
                urls.append(item["link"])
            elif isinstance(item, dict) and "text" in item:
                urls.append(item["text"])
        if urls:
            text = ", ".join(urls)
        else:
            text = str(value)
    else:
        text = str(value)
    text = text.replace("\ufeff", "").replace("\u200b", "").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def run_lark_cli(cmd):
    """执行 lark-cli 命令"""
    result = subprocess.run(cmd, shell=True, capture_output=True)
    if result.returncode != 0:
        print(f"命令执行失败: {cmd}")
        print(f"错误: {result.stderr.decode('utf-8', errors='replace')}")
        return None
    stdout = result.stdout.decode('utf-8', errors='replace')
    return json.loads(stdout)


def get_header_mapping(sheet_id, header_row=2):
    """读取表头映射"""
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


def get_sheet_row_count(sheet_id):
    """获取sheet行数"""
    cmd = f'lark-cli sheets +info --spreadsheet-token {SPREADSHEET_TOKEN}'
    result = run_lark_cli(cmd)
    if not result:
        return 1
    sheets = result.get("data", {}).get("sheets", {}).get("sheets", [])
    for sheet in sheets:
        if sheet.get("sheet_id") == sheet_id:
            return sheet.get("grid_properties", {}).get("row_count", 1)
    return 1


def read_sheet_data(sheet_id, start_row):
    """读取sheet数据"""
    total_rows = get_sheet_row_count(sheet_id)
    end_row = max(start_row + 300, total_rows)
    range_str = f"{sheet_id}!A{start_row}:BZ{end_row}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)
    if not result or not result.get("ok"):
        return []
    return result.get("data", {}).get("valueRange", {}).get("values", [])


def check_filter_conditions(blogger_name, sheet_id, sheet_name, header_row=2, start_row=3):
    """检查博主是否满足筛选条件"""
    print(f"\n{'='*60}")
    print(f"检查博主: {blogger_name}")
    print(f"{'='*60}")
    
    col_map = get_header_mapping(sheet_id, header_row)
    
    kol_col = col_map.get("达人量级")
    addr_col = col_map.get("产品邮寄地址")
    publish_col = col_map.get("发布状态")
    nickname_col = col_map.get("昵称")
    
    if nickname_col is None:
        print("未找到'昵称'列")
        return
    
    rows = read_sheet_data(sheet_id, start_row)
    
    for idx, row in enumerate(rows):
        if len(row) <= nickname_col:
            continue
        nick = normalize_text(row[nickname_col])
        if nick == blogger_name:
            row_num = start_row + idx
            print(f"找到博主 '{blogger_name}' (第 {row_num} 行)")
            
            kol_level = normalize_text(row[kol_col]) if kol_col is not None and kol_col < len(row) else ""
            address = normalize_text(row[addr_col]) if addr_col is not None and addr_col < len(row) else ""
            publish_status = normalize_text(row[publish_col]) if publish_col is not None and publish_col < len(row) else ""
            
            print(f"\n达人量级: '{kol_level}'")
            print(f"产品邮寄地址: '{address}'")
            print(f"发布状态: '{publish_status}'")
            
            print(f"\n筛选条件检查:")
            print(f"  1. 达人量级包含'KOL': {'✓' if 'KOL' in kol_level.upper() else '✗'}")
            print(f"  2. 地址不包含'抠图'或'扣图': {'✓' if '抠图' not in address and '扣图' not in address else '✗'}")
            print(f"  3. 地址不是'自有': {'✓' if address != '自有' else '✗'}")
            
            is_kol = "KOL" in kol_level.upper()
            is_not_koutu = "抠图" not in address and "扣图" not in address
            is_not_ziyou = address != "自有"
            
            if is_kol and is_not_koutu and is_not_ziyou:
                print(f"\n✓ 该博主满足所有筛选条件，应该同步到目标表")
            else:
                print(f"\n✗ 该博主不满足筛选条件，不应该同步到目标表")
            
            return
    
    print(f"未找到博主 '{blogger_name}'")


def main():
    bloggers = ["梅子家电攻略", "卷卷电器攻略"]
    
    # 检查小红书确认执行表
    sheet_id = "TDxBlT"
    
    for blogger in bloggers:
        check_filter_conditions(blogger, sheet_id, "小红书确认执行4月", header_row=2, start_row=3)


if __name__ == "__main__":
    main()
