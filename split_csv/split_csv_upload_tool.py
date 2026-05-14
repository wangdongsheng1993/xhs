import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable


CATEGORY_KEYWORDS = ("油烟机", "蒸烤", "洗碗机")
COMMERCIAL_COLUMN = "是否商业笔记"
TITLE_COLUMN = "笔记标题"
NO_TITLE_VALUE = "笔记暂未设置标题"
VISIBLE_SHEET_COLUMNS = {"笔记官方地址", "笔记标题", "主页链接"}


def subprocess_no_window_kwargs() -> dict:
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return {
        "creationflags": subprocess.CREATE_NO_WINDOW,
        "startupinfo": startupinfo,
    }


@dataclass
class SplitFile:
    index: int
    path: Path
    row_count: int
    sha256: str
    feishu_url: str = ""
    file_token: str = ""
    upload_status: str = "未上传"
    remark: str = ""


class LarkCliError(RuntimeError):
    def __init__(self, command: list[str], returncode: int, output: str, payload: dict | None = None):
        self.command = command
        self.returncode = returncode
        self.output = output
        self.payload = payload or {}
        super().__init__(f"命令执行失败({returncode}): {' '.join(command)}\n{output}")


def default_logger(message: str) -> None:
    print(message, flush=True)


def detect_category(source_path: Path) -> str:
    name = source_path.name
    for keyword in CATEGORY_KEYWORDS:
        if keyword in name:
            return keyword
    return "未分类"


def normalize_yyyymmdd(year: int, month: int, day: int) -> str:
    return f"{year:04d}{month:02d}{day:02d}"


def detect_date_from_name(source_path: Path, now: datetime | None = None) -> str:
    now = now or datetime.now()
    name = source_path.stem

    match = re.search(r"(20\d{2})[年._/-]?(0?[1-9]|1[0-2])[月._/-]?(0?[1-9]|[12]\d|3[01])日?", name)
    if match:
        return normalize_yyyymmdd(int(match.group(1)), int(match.group(2)), int(match.group(3)))

    match = re.search(r"(?<!\d)(0?[1-9]|1[0-2])[._-](0?[1-9]|[12]\d|3[01])(?!\d)", name)
    if match:
        return normalize_yyyymmdd(now.year, int(match.group(1)), int(match.group(2)))

    return now.strftime("%Y%m%d")


def read_csv_rows(source_path: Path) -> tuple[list[str], list[dict[str, str]]]:
    last_error = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with source_path.open("r", encoding=encoding, newline="") as file:
                reader = csv.DictReader(file)
                fieldnames = reader.fieldnames or []
                rows = list(reader)
            return fieldnames, rows
        except UnicodeDecodeError as exc:
            last_error = exc
    raise RuntimeError(f"无法识别 CSV 编码: {source_path}") from last_error


def filter_rows(fieldnames: list[str], rows: list[dict[str, str]]) -> list[dict[str, str]]:
    missing = [column for column in (COMMERCIAL_COLUMN, TITLE_COLUMN) if column not in fieldnames]
    if missing:
        raise ValueError(f"源文件缺少字段: {', '.join(missing)}")

    return [
        row
        for row in rows
        if (row.get(COMMERCIAL_COLUMN) or "").strip() == "否"
        and (row.get(TITLE_COLUMN) or "").strip() != NO_TITLE_VALUE
    ]


def split_rows(rows: list[dict[str, str]], chunk_size: int, min_last_rows: int) -> list[list[dict[str, str]]]:
    if chunk_size <= 0:
        raise ValueError("每个分片行数必须大于 0")
    if min_last_rows < 0:
        raise ValueError("最后分片最少行数不能小于 0")

    chunks = [rows[index : index + chunk_size] for index in range(0, len(rows), chunk_size)]
    if len(chunks) > 1 and len(chunks[-1]) < min_last_rows:
        chunks[-2].extend(chunks[-1])
        chunks.pop()
    return chunks


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_csv(path: Path, fieldnames: list[str], rows: Iterable[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def prepare_split_files(
    source_path: Path,
    output_root: Path,
    chunk_size: int,
    min_last_rows: int,
    folder_date: str | None = None,
    run_date: str | None = None,
    logger: Callable[[str], None] = default_logger,
) -> tuple[str, str, Path, Path, list[SplitFile]]:
    source_path = source_path.resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"源文件不存在: {source_path}")

    now = datetime.now()
    category = detect_category(source_path)
    folder_date = folder_date or detect_date_from_name(source_path, now)
    run_date = run_date or now.strftime("%Y%m%d")
    output_dir = output_root / category / folder_date
    output_dir.mkdir(parents=True, exist_ok=True)

    fieldnames, rows = read_csv_rows(source_path)
    filtered = filter_rows(fieldnames, rows)
    full_output = output_dir / f"{category}_{folder_date}_筛选结果.csv"
    write_csv(full_output, fieldnames, filtered)

    pattern = f"{category}_{run_date}_[0-9][0-9].csv"
    for stale in output_dir.glob(pattern):
        stale.unlink()

    split_files: list[SplitFile] = []
    for index, chunk in enumerate(split_rows(filtered, chunk_size, min_last_rows), 1):
        split_path = output_dir / f"{category}_{run_date}_{index:02d}.csv"
        write_csv(split_path, fieldnames, chunk)
        split_files.append(
            SplitFile(
                index=index,
                path=split_path,
                row_count=len(chunk),
                sha256=file_sha256(split_path),
            )
        )

    logger(f"源文件行数: {len(rows)}")
    logger(f"筛选后行数: {len(filtered)}")
    logger(f"输出目录: {output_dir}")
    for split_file in split_files:
        logger(f"分片 {split_file.path.name}: {split_file.row_count} 行")

    return category, folder_date, output_dir, full_output, split_files


def parse_first_json(text: str) -> dict:
    decoder = json.JSONDecoder()
    for index, char in enumerate(text):
        if char != "{":
            continue
        try:
            data, _ = decoder.raw_decode(text[index:])
            return data
        except json.JSONDecodeError:
            continue
    raise ValueError(f"命令输出中没有 JSON: {text[:500]}")


def run_command(command: list[str], logger: Callable[[str], None] = default_logger, cwd: Path | None = None) -> dict:
    if cwd:
        logger(f"$ cd {cwd}")
    logger("$ " + " ".join(command))
    completed = subprocess.run(
        command,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        cwd=cwd,
        **subprocess_no_window_kwargs(),
    )
    output = (completed.stdout or "") + (completed.stderr or "")
    payload = None
    try:
        payload = parse_first_json(output)
    except ValueError:
        pass
    if completed.returncode != 0:
        raise LarkCliError(command, completed.returncode, output, payload)
    if payload is None:
        raise ValueError(f"命令输出中没有 JSON: {output[:500]}")
    return payload


def without_as_user(command: list[str]) -> list[str]:
    cleaned: list[str] = []
    index = 0
    while index < len(command):
        if command[index] == "--as" and index + 1 < len(command) and command[index + 1] == "user":
            index += 2
            continue
        cleaned.append(command[index])
        index += 1
    return cleaned


def first_value_by_keys(data: object, keys: set[str]) -> str:
    if isinstance(data, dict):
        for key, value in data.items():
            if key in keys and isinstance(value, str) and value:
                return value
        for value in data.values():
            found = first_value_by_keys(value, keys)
            if found:
                return found
    elif isinstance(data, list):
        for item in data:
            found = first_value_by_keys(item, keys)
            if found:
                return found
    return ""


def build_feishu_url(token: str, remote_type: str) -> str:
    if not token:
        return ""
    if remote_type == "sheet":
        return f"https://www.feishu.cn/sheets/{token}"
    return f"https://www.feishu.cn/file/{token}"


def token_from_url(url: str, remote_type: str) -> str:
    if not url:
        return ""
    if remote_type == "sheet":
        match = re.search(r"/sheets/([^/?#]+)", url)
    else:
        match = re.search(r"/file/([^/?#]+)", url)
    return match.group(1) if match else ""


def remote_info_from_response(data: dict, remote_type: str) -> tuple[str, str]:
    body = data.get("data", data)
    url = first_value_by_keys(body, {"url"})
    token = token_from_url(url, remote_type)
    if not token:
        if remote_type == "sheet":
            token = first_value_by_keys(body, {"spreadsheet_token", "spreadsheetToken", "token"})
        else:
            token = first_value_by_keys(body, {"file_token", "token"})
    if not url:
        url = build_feishu_url(token, remote_type)
    return token, url


def is_missing_retrieve_permission(exc: LarkCliError) -> bool:
    error = exc.payload.get("error", {})
    message = f"{error.get('message', '')}\n{error.get('hint', '')}\n{exc.output}"
    return "space:document:retrieve" in message or "insufficient permissions" in message


def is_missing_sheet_permission(exc: LarkCliError) -> bool:
    error = exc.payload.get("error", {})
    message = f"{error.get('message', '')}\n{error.get('hint', '')}\n{exc.output}"
    return "sheets:spreadsheet:read" in message or "sheets:spreadsheet:write_only" in message


def require_lark_cli() -> str:
    executable = shutil.which("lark-cli")
    if not executable:
        raise RuntimeError("未找到 lark-cli，请先安装并确认 lark-cli 在 PATH 中。")
    return executable


def list_drive_children(lark_cli: str, folder_token: str, logger: Callable[[str], None]) -> list[dict]:
    params = {"page_size": 200}
    if folder_token:
        params["folder_token"] = folder_token
    try:
        data = run_command(
            [
                lark_cli,
                "drive",
                "files",
                "list",
                "--as",
                "user",
                "--page-all",
                "--params",
                json.dumps(params, ensure_ascii=False),
            ],
            logger,
        )
    except LarkCliError as exc:
        if not is_missing_retrieve_permission(exc):
            raise
        raise RuntimeError(
            "飞书读取目录权限不足，无法复用已有目录或跳过同名文件。\n"
            '请在运行 GUI 的同一个 Windows 环境中执行：lark-cli auth login --scope "space:document:retrieve"\n'
            "授权完成后重新运行本工具。"
        ) from exc
    return data.get("data", {}).get("files", []) or data.get("files", []) or []


def get_existing_child(children: list[dict], name: str, expected_type: str | None = None) -> dict | None:
    for child in children:
        if child.get("name") != name:
            continue
        if expected_type and child.get("type") != expected_type:
            continue
        return child
    return None


def get_or_create_folder(
    lark_cli: str,
    name: str,
    parent_token: str,
    logger: Callable[[str], None] = default_logger,
) -> tuple[str, str]:
    children = list_drive_children(lark_cli, parent_token, logger)
    existing = get_existing_child(children, name, "folder")
    if existing:
        token = existing.get("token", "")
        url = existing.get("url", "")
        logger(f"复用飞书文件夹: {name} {url or token}")
        return token, url

    data_payload = {"folder_token": parent_token or "", "name": name}
    command = [
        lark_cli,
        "drive",
        "files",
        "create_folder",
        "--as",
        "user",
        "--data",
        json.dumps(data_payload, ensure_ascii=False),
    ]
    try:
        data = run_command(command, logger)
    except LarkCliError as exc:
        if "unknown flag: --as" in exc.output:
            logger("当前 lark-cli 版本的建目录命令不支持 --as，去掉该参数后重试。")
            data = run_command(without_as_user(command), logger)
        elif "--yes" in exc.output:
            data = run_command(command + ["--yes"], logger)
        else:
            raise
    if not isinstance(data, dict):
        raise RuntimeError(f"创建文件夹失败，返回值异常: {data}")
    folder = data.get("data", {})
    token = folder.get("folder_token") or folder.get("token")
    if not token:
        raise RuntimeError(f"创建文件夹失败，未返回 token: {data}")
    return token, folder.get("url", "")


def get_first_sheet_id(lark_cli: str, spreadsheet_token: str, logger: Callable[[str], None]) -> str:
    try:
        data = run_command(
            [
                lark_cli,
                "sheets",
                "+info",
                "--as",
                "user",
                "--spreadsheet-token",
                spreadsheet_token,
            ],
            logger,
        )
    except LarkCliError as exc:
        if not is_missing_sheet_permission(exc):
            raise
        raise RuntimeError(
            "飞书表格权限不足，无法隐藏列。\n"
            '请在 GUI 中点击“补充飞书读取授权”，或执行：lark-cli auth login --scope "space:document:retrieve sheets:spreadsheet:read sheets:spreadsheet:write_only"\n'
            "授权完成后重新运行本工具。"
        ) from exc
    sheet_id = first_value_by_keys(data.get("data", data), {"sheet_id", "sheetId"})
    if not sheet_id:
        raise RuntimeError(f"无法从表格信息中获取 sheet_id: {data}")
    return sheet_id


def get_hidden_column_ranges(csv_path: Path) -> list[tuple[int, int]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file)
        headers = next(reader, [])

    hidden_indexes = [
        index
        for index, header in enumerate(headers, 1)
        if header.strip() not in VISIBLE_SHEET_COLUMNS
    ]
    ranges: list[tuple[int, int]] = []
    if not hidden_indexes:
        return ranges

    start = previous = hidden_indexes[0]
    for index in hidden_indexes[1:]:
        if index == previous + 1:
            previous = index
            continue
        ranges.append((start, previous))
        start = previous = index
    ranges.append((start, previous))
    return ranges


def hide_unneeded_sheet_columns(
    lark_cli: str,
    spreadsheet_token: str,
    csv_path: Path,
    logger: Callable[[str], None],
) -> None:
    if not spreadsheet_token:
        logger("未获取到电子表格 token，跳过隐藏列。")
        return

    sheet_id = get_first_sheet_id(lark_cli, spreadsheet_token, logger)
    hidden_ranges = get_hidden_column_ranges(csv_path)
    if not hidden_ranges:
        logger("没有需要隐藏的列。")
        return

    for start_index, end_index in hidden_ranges:
        update_sheet_dimension(
            lark_cli,
            spreadsheet_token,
            sheet_id,
            start_index,
            end_index,
            logger,
        )
        logger(f"已隐藏列: {start_index}-{end_index}")


def should_fallback_dimension_api(exc: LarkCliError) -> bool:
    text = exc.output
    return (
        "unknown flag: --spreadsheet-token" in text
        or "unknown command" in text.lower()
        or "Use \"lark-cli sheets [command] --help\"" in text
    )


def update_sheet_dimension(
    lark_cli: str,
    spreadsheet_token: str,
    sheet_id: str,
    start_index: int,
    end_index: int,
    logger: Callable[[str], None],
) -> None:
    command = [
        lark_cli,
        "sheets",
        "+update-dimension",
        "--as",
        "user",
        "--spreadsheet-token",
        spreadsheet_token,
        "--sheet-id",
        sheet_id,
        "--dimension",
        "COLUMNS",
        "--start-index",
        str(start_index),
        "--end-index",
        str(end_index),
    ]
    try:
        run_command(command, logger)
        return
    except LarkCliError as exc:
        if "unknown flag: --as" in exc.output:
            logger("当前 lark-cli 版本的表格命令不支持 --as，去掉该参数后重试。")
            try:
                run_command(without_as_user(command), logger)
                return
            except LarkCliError as retry_exc:
                if should_fallback_dimension_api(retry_exc):
                    logger("当前 lark-cli 版本没有 sheets +update-dimension，改用底层 API 隐藏列。")
                    update_sheet_dimension_by_api(
                        lark_cli,
                        spreadsheet_token,
                        sheet_id,
                        start_index,
                        end_index,
                        logger,
                    )
                    return
                raise
        if should_fallback_dimension_api(exc):
            logger("当前 lark-cli 版本没有 sheets +update-dimension，改用底层 API 隐藏列。")
            update_sheet_dimension_by_api(
                lark_cli,
                spreadsheet_token,
                sheet_id,
                start_index,
                end_index,
                logger,
            )
            return
        if is_missing_sheet_permission(exc):
            raise RuntimeError(
                "飞书表格权限不足，无法隐藏列。\n"
                '请在 GUI 中点击“补充飞书读取授权”，或执行：lark-cli auth login --scope "space:document:retrieve sheets:spreadsheet:read sheets:spreadsheet:write_only"\n'
                "授权完成后重新运行本工具。"
            ) from exc
        raise


def update_sheet_dimension_by_api(
    lark_cli: str,
    spreadsheet_token: str,
    sheet_id: str,
    start_index: int,
    end_index: int,
    logger: Callable[[str], None],
) -> None:
    body = {
        "dimension": {
            "sheetId": sheet_id,
            "majorDimension": "COLUMNS",
            "startIndex": start_index,
            "endIndex": end_index,
        },
        "dimensionProperties": {
            "visible": False,
        },
    }
    command = [
        lark_cli,
        "api",
        "PUT",
        f"/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range",
        "--as",
        "user",
        "--data",
        json.dumps(body, ensure_ascii=False),
    ]
    try:
        run_command(command, logger)
    except LarkCliError as exc:
        if "unknown flag: --as" in exc.output:
            logger("当前 lark-cli 版本的 API 命令不支持 --as，去掉该参数后重试。")
            run_command(without_as_user(command), logger)
            return
        if is_missing_sheet_permission(exc):
            raise RuntimeError(
                "飞书表格权限不足，无法隐藏列。\n"
                '请在 GUI 中点击“补充飞书读取授权”，或执行：lark-cli auth login --scope "space:document:retrieve sheets:spreadsheet:read sheets:spreadsheet:write_only"\n'
                "授权完成后重新运行本工具。"
            ) from exc
        raise


def upload_split_files(
    split_files: list[SplitFile],
    category: str,
    folder_date: str,
    skip_existing: bool = True,
    import_as_sheet: bool = True,
    logger: Callable[[str], None] = default_logger,
) -> tuple[str, list[SplitFile]]:
    lark_cli = require_lark_cli()
    category_token, _ = get_or_create_folder(lark_cli, category, "", logger)
    date_token, date_url = get_or_create_folder(lark_cli, folder_date, category_token, logger)

    existing_files = list_drive_children(lark_cli, date_token, logger)
    if import_as_sheet:
        existing_by_name = {
            item.get("name"): item
            for item in existing_files
            if item.get("type") in {"sheet", "spreadsheet"}
        }
    else:
        existing_by_name = {item.get("name"): item for item in existing_files if item.get("type") == "file"}

    for split_file in split_files:
        remote_name = split_file.path.stem if import_as_sheet else split_file.path.name
        remote_type = "sheet" if import_as_sheet else "file"
        existing = existing_by_name.get(remote_name)
        if existing and skip_existing:
            split_file.upload_status = "已存在，跳过上传"
            split_file.file_token = existing.get("token", "")
            split_file.feishu_url = existing.get("url", "") or build_feishu_url(split_file.file_token, remote_type)
            logger(f"跳过已存在文件: {remote_name}")
            if import_as_sheet:
                hide_unneeded_sheet_columns(lark_cli, split_file.file_token, split_file.path, logger)
                split_file.upload_status = "已存在，已检查隐藏列"
            continue

        if import_as_sheet:
            upload_command = [
                lark_cli,
                "drive",
                "+import",
                "--as",
                "user",
                "--folder-token",
                date_token,
                "--file",
                split_file.path.name,
                "--name",
                remote_name,
                "--type",
                "sheet",
            ]
        else:
            upload_command = [
                lark_cli,
                "drive",
                "+upload",
                "--as",
                "user",
                "--folder-token",
                date_token,
                "--file",
                split_file.path.name,
                "--name",
                remote_name,
            ]
        try:
            data = run_command(upload_command, logger, cwd=split_file.path.parent)
        except LarkCliError as exc:
            if "unknown flag: --as" not in exc.output:
                raise
            logger("当前 lark-cli 版本的命令不支持 --as，去掉该参数后重试。")
            data = run_command(without_as_user(upload_command), logger, cwd=split_file.path.parent)
        split_file.file_token, split_file.feishu_url = remote_info_from_response(data, remote_type)
        split_file.upload_status = "已导入为电子表格" if import_as_sheet else "已上传"
        logger(f"{split_file.upload_status}: {remote_name} {split_file.feishu_url}")
        if import_as_sheet:
            hide_unneeded_sheet_columns(lark_cli, split_file.file_token, split_file.path, logger)

    return date_url, split_files


def write_tracking_files(
    output_dir: Path,
    category: str,
    folder_date: str,
    remote_folder_url: str,
    split_files: list[SplitFile],
) -> tuple[Path, Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = output_dir / f"上传跟踪_{category}_{folder_date}_{timestamp}.csv"
    md_path = output_dir / f"上传跟踪_{category}_{folder_date}_{timestamp}.md"

    fieldnames = [
        "分片名称",
        "本地路径",
        "行数",
        "sha256",
        "飞书地址",
        "飞书文件token",
        "上传状态",
        "备注",
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for split_file in split_files:
            writer.writerow(
                {
                    "分片名称": split_file.path.name,
                    "本地路径": str(split_file.path),
                    "行数": split_file.row_count,
                    "sha256": split_file.sha256,
                    "飞书地址": split_file.feishu_url,
                    "飞书文件token": split_file.file_token,
                    "上传状态": split_file.upload_status,
                    "备注": split_file.remark,
                }
            )

    lines = [
        f"# {category}/{folder_date} 上传跟踪",
        "",
        f"- 远端目录: {remote_folder_url or '(未上传)'}",
        f"- 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "| 分片名称 | 行数 | 上传状态 | 飞书地址 | 备注 |",
        "| --- | ---: | --- | --- | --- |",
    ]
    for split_file in split_files:
        url_text = split_file.feishu_url or ""
        lines.append(
            f"| {split_file.path.name} | {split_file.row_count} | {split_file.upload_status} | {url_text} | {split_file.remark} |"
        )
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return csv_path, md_path


def run_pipeline(
    source: Path,
    output_root: Path,
    upload: bool,
    folder_date: str | None = None,
    chunk_size: int = 200,
    min_last_rows: int = 100,
    skip_existing: bool = True,
    import_as_sheet: bool = True,
    logger: Callable[[str], None] = default_logger,
) -> tuple[Path, Path, Path, list[SplitFile]]:
    category, detected_date, output_dir, _full_output, split_files = prepare_split_files(
        source,
        output_root,
        chunk_size=chunk_size,
        min_last_rows=min_last_rows,
        folder_date=folder_date,
        logger=logger,
    )
    remote_folder_url = ""
    if upload:
        remote_folder_url, split_files = upload_split_files(
            split_files,
            category,
            detected_date,
            skip_existing=skip_existing,
            import_as_sheet=import_as_sheet,
            logger=logger,
        )
    tracking_csv, tracking_md = write_tracking_files(output_dir, category, detected_date, remote_folder_url, split_files)
    logger(f"跟踪 CSV: {tracking_csv}")
    logger(f"跟踪 MD: {tracking_md}")
    return output_dir, tracking_csv, tracking_md, split_files


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="xhs CSV 筛选、分片并上传飞书")
    parser.add_argument("source", help="源 CSV 文件")
    parser.add_argument("--output-root", default=".", help="输出根目录，默认当前目录")
    parser.add_argument("--date", help="目录日期 YYYYMMDD；不填则从文件名解析，解析不到用今天")
    parser.add_argument("--chunk-size", type=int, default=200, help="每个分片的目标行数")
    parser.add_argument("--min-last-rows", type=int, default=100, help="最后一个分片不足该行数时合并到倒数第二个")
    parser.add_argument("--no-upload", action="store_true", help="只生成本地分片，不上传飞书")
    parser.add_argument("--allow-duplicate-upload", action="store_true", help="远端同名文件已存在时仍继续上传")
    parser.add_argument("--upload-csv-file", action="store_true", help="上传为普通 CSV 文件，不导入为飞书电子表格")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        run_pipeline(
            source=Path(args.source),
            output_root=Path(args.output_root),
            upload=not args.no_upload,
            folder_date=args.date,
            chunk_size=args.chunk_size,
            min_last_rows=args.min_last_rows,
            skip_existing=not args.allow_duplicate_upload,
            import_as_sheet=not args.upload_csv_file,
        )
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
