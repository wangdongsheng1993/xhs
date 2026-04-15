#!/usr/bin/env python3
"""
搜索二核表中包含特定关键词的博主
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


def search_blogger(sheet_id, sheet_name, keyword, header_row=2, start_row=3):
    """搜索包含关键词的博主"""
    print(f"\n{'='*60}")
    print(f"在 {sheet_name} 中搜索: '{keyword}'")
    print(f"{'='*60}")
    
    col_map = get_header_mapping(sheet_id, header_row)
    nickname_col = col_map.get("昵称")
    
    if nickname_col is None:
        print("未找到'昵称'列")
        return []
    
    rows = read_sheet_data(sheet_id, start_row)
    print(f"读取到 {len(rows)} 行数据")
    
    matches = []
    for idx, row in enumerate(rows):
        if len(row) <= nickname_col:
            continue
        nick = normalize_text(row[nickname_col])
        if keyword in nick:
            matches.append((start_row + idx, nick, row, col_map))
    
    return matches


def main():
    keyword = sys.argv[1] if len(sys.argv) > 1 else "梅花"
    
    # 搜索抖音二核表
    erhe_sheet_id = "jRY56A"
    erhe_matches = search_blogger(erhe_sheet_id, "抖音二核表", keyword, header_row=2, start_row=3)
    
    if erhe_matches:
        print(f"\n找到 {len(erhe_matches)} 个匹配:")
        for row_num, nick, row, col_map in erhe_matches:
            print(f"\n  第 {row_num} 行: {nick}")
            id_col = col_map.get("ID")
            link_col = col_map.get("主页链接")
            form_col = col_map.get("合作形式")
            price_col = col_map.get("平台价格")
            if id_col is not None and id_col < len(row):
                print(f"    ID: {normalize_text(row[id_col])}")
            if link_col is not None and link_col < len(row):
                print(f"    主页链接: {normalize_text(row[link_col])}")
            if form_col is not None and form_col < len(row):
                print(f"    合作形式: {normalize_text(row[form_col])}")
            if price_col is not None and price_col < len(row):
                print(f"    平台价格: {normalize_text(row[price_col])}")
    else:
        print(f"\n未找到包含 '{keyword}' 的博主")
    
    # 同时搜索确认执行表
    confirm_sheet_id = "iFO2b6"
    confirm_matches = search_blogger(confirm_sheet_id, "4月抖音确认执行", keyword, header_row=2, start_row=3)
    
    if confirm_matches:
        print(f"\n找到 {len(confirm_matches)} 个匹配:")
        for row_num, nick, row, col_map in confirm_matches:
            print(f"\n  第 {row_num} 行: {nick}")
            id_col = col_map.get("ID")
            link_col = col_map.get("主页链接")
            form_col = col_map.get("合作形式")
            price_col = col_map.get("平台价（平台裸价）")
            kol_col = col_map.get("达人量级")
            if kol_col is not None and kol_col < len(row):
                print(f"    达人量级: {normalize_text(row[kol_col])}")
            if id_col is not None and id_col < len(row):
                print(f"    ID: {normalize_text(row[id_col])}")
            if link_col is not None and link_col < len(row):
                print(f"    主页链接: {normalize_text(row[link_col])}")
            if form_col is not None and form_col < len(row):
                print(f"    合作形式: {normalize_text(row[form_col])}")
            if price_col is not None and price_col < len(row):
                print(f"    平台价（平台裸价）: {normalize_text(row[price_col])}")
    else:
        print(f"\n未找到包含 '{keyword}' 的博主")


if __name__ == "__main__":
    main()
