#!/usr/bin/env python3
"""
从日志文件中提取已更新的笔记URL，并更新到CSV文件
"""
import re
import csv

LOG_FILE = "/mnt/c/code_20251212/AI/xhs/cvs_handle/logs/fix_urls_20260512_215959_worker0.log"
INPUT_CSV = "/mnt/c/code_20251212/AI/xhs/cvs_handle/xhs_export_excel20265_洗碗机-05.11-1041_831751206_1778549825836.csv"
OUTPUT_CSV = "/mnt/c/code_20251212/AI/xhs/cvs_handle/xhs_export_excel20265_洗碗机-05.11-1041_831751206_1778549825836_结果.csv"

def parse_log_for_updates(log_file):
    """从日志文件中提取所有成功更新的URL"""
    updates = {}  # {row_number: url}

    with open(log_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配模式: [行号] 第 X 行: 标题
    # 例如: [1/1041] 第 2 行: 别买错了！普通洗碗机vs消毒洗碗机真不一样
    # 然后某处会有:    - 已更新: https://www.xiaohongshu.com/explore/...
    #
    # 或在同一行: [1/1041] 第 2 行: ... - 已更新: URL

    # 先找到所有 "已更新:" 后面跟URL的行（可能在多行后的缩进行）
    # 提取行号和URL
    # 匹配 "[序号/总数] 第 行号 行:" 后面跟 "- 已更新: URL"

    # 分两步：1) 找到所有行号信息 2) 找到所有URL信息
    # 行号模式
    index_pattern = r'\[(\d+)/(\d+)\] 第 (\d+) 行:'
    # URL模式（缩进的已更新行）
    url_pattern = r'^\s*- 已更新: (https?://\S+)'

    # 找到所有"已更新:"的URL及其上下文
    url_matches = re.findall(url_pattern, content, re.MULTILINE | re.IGNORECASE)

    # 找到所有行号索引
    index_matches = re.findall(index_pattern, content)

    # 按出现顺序配对：最近的已更新URL对应最近的行号
    # 由于日志是顺序的，我们可以按顺序建立映射
    # 实际模式：[序号/总数] 第 行号 行: 标题  ...  - 已更新: URL
    # 提取每个"已更新:"前的行号上下文

    # 更简单的方法：按顺序扫描，每遇到"[X/总数] 第 Y 行:"时记录Y，
    # 每遇到"- 已更新: URL"时用最近记录的Y

    updates = {}
    current_row = None

    for line in content.split('\n'):
        # 匹配行号行
        row_match = re.search(r'\[(\d+)/(\d+)\] 第 (\d+) 行:', line)
        if row_match:
            current_row = int(row_match.group(3))

        # 匹配已更新行
        url_match = re.search(r'- 已更新: (https?://\S+)', line)
        if url_match and current_row is not None:
            url = url_match.group(1).rstrip()
            updates[current_row] = url

    print(f"从日志中提取到 {len(updates)} 条成功更新的URL")
    return updates

def update_csv(input_csv, output_csv, updates):
    """更新CSV文件的笔记官方地址列"""
    rows = []

    with open(input_csv, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        header = next(reader)
        rows.append(header)

        # 找到"笔记官方地址"列的索引
        try:
            url_col_index = header.index('笔记官方地址')
        except ValueError:
            print(f"错误：找不到'笔记官方地址'列")
            print(f"表头: {header}")
            return

        print(f"表头共 {len(header)} 列")
        print(f"'笔记官方地址' 在第 {url_col_index + 1} 列")

        # 处理数据行
        for row_idx, row in enumerate(reader, start=2):  # 从第2行开始（跳过表头）
            if row_idx in updates:
                row[url_col_index] = updates[row_idx]

            rows.append(row)

    # 写入结果文件
    with open(output_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(rows)

    print(f"已保存到: {output_csv}")
    print(f"共写入 {len(rows)} 行（含表头）")

def main():
    print("=" * 60)
    print("从日志提取更新URL并更新CSV")
    print("=" * 60)
    print()

    # 1. 从日志提取更新
    updates = parse_log_for_updates(LOG_FILE)
    print(f"\n前10条更新:")
    for i, (row, url) in enumerate(sorted(updates.items())[:10]):
        print(f"  行{row}: {url[:80]}...")

    # 2. 更新CSV
    update_csv(INPUT_CSV, OUTPUT_CSV, updates)

    print("\n完成!")

if __name__ == "__main__":
    main()