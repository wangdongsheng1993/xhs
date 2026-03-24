import argparse
import os
import subprocess
import sys


DEFAULT_BRAND_INPUT = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表 副本.xlsx"
DEFAULT_BRAND_OUTPUT = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"

DEFAULT_ECOM_INPUT = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表 副本 (1).xlsx"
DEFAULT_ECOM_OUTPUT = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表 副本 (1)_结果.xlsx"
DEFAULT_ECOM_FALLBACK_OUTPUT = r"c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表 副本 (1)_结果_自动保存.xlsx"


def parse_rows_spec(spec):
    """把 326-329,338 这种写法展开成具体行号列表。"""
    rows = []
    for part in str(spec or "").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if end < start:
                start, end = end, start
            rows.extend(range(start, end + 1))
        else:
            rows.append(int(part))
    return rows


def resolve_mode(mode_text):
    """把常见别名统一映射到 brand / ecommerce。"""
    text = str(mode_text or "").strip().lower()
    if text in {"brand", "品牌", "小红书品牌-kol", "品牌sheet", "品牌表"}:
        return "brand"
    if text in {"ecommerce", "电商", "小红书电商-kol", "电商sheet", "电商表"}:
        return "ecommerce"
    raise ValueError(f"无法识别的 sheet/mode: {mode_text}")


def build_mode_config(mode):
    """根据 mode 返回对应脚本和默认文件路径。"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if mode == "brand":
        return {
            "script": os.path.join(base_dir, "xhs_extractor.py"),
            "sheet_name": "小红书品牌-KOL",
            "excel_path": DEFAULT_BRAND_INPUT,
            "output_path": DEFAULT_BRAND_OUTPUT,
            "fallback_output_path": "",
        }
    return {
        "script": os.path.join(base_dir, "xhs_ecommerce_extractor.py"),
        "sheet_name": "小红书电商-KOL",
        "excel_path": DEFAULT_ECOM_INPUT,
        "output_path": DEFAULT_ECOM_OUTPUT,
        "fallback_output_path": DEFAULT_ECOM_FALLBACK_OUTPUT,
    }


def main():
    parser = argparse.ArgumentParser(description="统一入口：按 sheet 和行号调用品牌/电商 Excel 抓取脚本。")
    parser.add_argument("mode", help="brand / ecommerce，或直接写 品牌 / 电商")
    parser.add_argument("rows", help="要处理的行号，如 338-340 或 326,327,329")
    parser.add_argument("--excel", dest="excel_path", default="", help="自定义输入 Excel 路径")
    parser.add_argument("--output", dest="output_path", default="", help="自定义输出 Excel 路径")
    parser.add_argument("--verbose", action="store_true", help="开启详细调试日志")
    args = parser.parse_args()

    mode = resolve_mode(args.mode)
    rows = parse_rows_spec(args.rows)
    config = build_mode_config(mode)

    excel_path = args.excel_path.strip() or config["excel_path"]
    output_path = args.output_path.strip() or config["output_path"]

    print(f"模式: {mode}")
    print(f"Sheet: {config['sheet_name']}")
    print(f"输入文件: {excel_path}")
    print(f"输出文件: {output_path}")
    print(f"处理行: {rows}")

    env = os.environ.copy()
    env["XHS_DEBUG_ROW"] = args.rows.strip()
    env["XHS_DEBUG_NAME"] = ""
    env["XHS_DEBUG_MAX_ROWS"] = ""
    env["XHS_DEBUG_VERBOSE"] = "1" if args.verbose else "0"
    env["XHS_EXCEL_PATH"] = excel_path
    env["XHS_OUTPUT_PATH"] = output_path
    env["XHS_SHEET_NAME"] = config["sheet_name"]
    if config["fallback_output_path"]:
        env["XHS_FALLBACK_OUTPUT_PATH"] = config["fallback_output_path"]

    result = subprocess.run([sys.executable, config["script"]], env=env, check=False)
    if result.returncode != 0:
        raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
