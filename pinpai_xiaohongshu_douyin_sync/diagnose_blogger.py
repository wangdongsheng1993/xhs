#!/usr/bin/env python3
"""
诊断脚本：检查特定博主的同步状态
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
    end_row = max(start_row + 200, total_rows)
    range_str = f"{sheet_id}!A{start_row}:BZ{end_row}"
    cmd = f'lark-cli sheets +read --spreadsheet-token {SPREADSHEET_TOKEN} --sheet-id {sheet_id} --range "{range_str}"'
    result = run_lark_cli(cmd)
    if not result or not result.get("ok"):
        return []
    return result.get("data", {}).get("valueRange", {}).get("values", [])


def find_blogger_in_sheet(sheet_id, sheet_name, blogger_name, header_row=2, start_row=3, nickname_col_name="昵称"):
    """在指定sheet中查找博主"""
    print(f"\n{'='*60}")
    print(f"检查表: {sheet_name}")
    print(f"{'='*60}")
    
    col_map = get_header_mapping(sheet_id, header_row)
    print(f"列名: {list(col_map.keys())[:15]}")
    
    nickname_col = col_map.get(nickname_col_name)
    if nickname_col is None:
        print(f"未找到'{nickname_col_name}'列")
        return None
    
    rows = read_sheet_data(sheet_id, start_row)
    print(f"读取到 {len(rows)} 行数据")
    
    for idx, row in enumerate(rows):
        if len(row) <= nickname_col:
            continue
        nick = normalize_text(row[nickname_col])
        if nick == blogger_name:
            print(f"\n✓ 找到博主 '{blogger_name}' (第 {start_row + idx} 行)")
            
            result = {"row_num": start_row + idx, "data": {}}
            
            for col_name, col_idx in col_map.items():
                if col_idx < len(row):
                    result["data"][col_name] = normalize_text(row[col_idx])
                else:
                    result["data"][col_name] = ""
            
            return result
    
    print(f"\n✗ 未找到博主 '{blogger_name}'")
    return None


def check_sync_status(erhe_data, confirm_data):
    """检查同步状态"""
    print(f"\n{'='*60}")
    print("同步状态检查")
    print(f"{'='*60}")
    
    fields_to_check = ["ID", "主页链接", "合作形式"]
    
    for field in fields_to_check:
        erhe_val = erhe_data["data"].get(field, "") if erhe_data else ""
        confirm_val = confirm_data["data"].get(field, "") if confirm_data else ""
        
        print(f"\n字段: {field}")
        print(f"  二核表: '{erhe_val}'")
        print(f"  确认执行表: '{confirm_val}'")
        
        if erhe_val and not confirm_val:
            print(f"  ⚠ 应该同步（目标为空，源有值）")
        elif erhe_val == confirm_val:
            print(f"  ✓ 已同步")
        elif confirm_val:
            print(f"  ⚠ 目标已有值，根据规则不会覆盖")
        else:
            print(f"  - 源为空，无法同步")
    
    # 检查平台价格
    erhe_price = erhe_data["data"].get("平台价格", "") if erhe_data else ""
    confirm_price = confirm_data["data"].get("平台价（平台裸价）", "") if confirm_data else ""
    
    print(f"\n字段: 平台价格")
    print(f"  二核表平台价格: '{erhe_price}'")
    print(f"  确认执行表平台价（平台裸价）: '{confirm_price}'")
    
    if erhe_price and not confirm_price:
        print(f"  ⚠ 应该同步（目标为空，源有值）")
    elif erhe_price == confirm_price:
        print(f"  ✓ 已同步")
    elif confirm_price:
        print(f"  ⚠ 目标已有值，根据规则不会覆盖")


def main():
    blogger_name = sys.argv[1] if len(sys.argv) > 1 else "梅花鹿鹿"
    
    print(f"\n{'='*60}")
    print(f"诊断博主: {blogger_name}")
    print(f"{'='*60}")
    
    # 检查抖音二核表
    erhe_sheet_id = "jRY56A"
    erhe_data = find_blogger_in_sheet(
        erhe_sheet_id,
        "抖音二核表",
        blogger_name,
        header_row=2,
        start_row=3
    )
    
    if erhe_data:
        print(f"\n关键字段:")
        print(f"  昵称: {erhe_data['data'].get('昵称', '')}")
        print(f"  ID: {erhe_data['data'].get('ID', '')}")
        print(f"  主页链接: {erhe_data['data'].get('主页链接', '')}")
        print(f"  合作形式: {erhe_data['data'].get('合作形式', '')}")
        print(f"  平台价格: {erhe_data['data'].get('平台价格', '')}")
    
    # 检查4月抖音确认执行表
    confirm_sheet_id = "iFO2b6"
    confirm_data = find_blogger_in_sheet(
        confirm_sheet_id, 
        "4月抖音确认执行", 
        blogger_name,
        header_row=2,
        start_row=3
    )
    
    if confirm_data:
        print(f"\n关键字段:")
        print(f"  昵称: {confirm_data['data'].get('昵称', '')}")
        print(f"  达人量级: {confirm_data['data'].get('达人量级', '')}")
        print(f"  ID: {confirm_data['data'].get('ID', '')}")
        print(f"  主页链接: {confirm_data['data'].get('主页链接', '')}")
        print(f"  合作形式: {confirm_data['data'].get('合作形式', '')}")
        print(f"  平台价（平台裸价）: {confirm_data['data'].get('平台价（平台裸价）', '')}")
        print(f"  气源: {confirm_data['data'].get('气源', '')}")
        print(f"  产品邮寄地址: {confirm_data['data'].get('产品邮寄地址', '')}")
    
    # 检查同步状态
    check_sync_status(erhe_data, confirm_data)


if __name__ == "__main__":
    main()
