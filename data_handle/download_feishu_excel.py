import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.parse
import urllib.request


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "feishu_download_config.json")
DEFAULT_OUTPUT_NAME = "【内部深演智能】老板电器C5 提号表.xlsx"
DEFAULT_OUTPUT_PATH = os.path.join(BASE_DIR, DEFAULT_OUTPUT_NAME)
DEFAULT_FEISHU_SOURCE = "https://mv21kbvltn.feishu.cn/wiki/PzBaw9C66iZtRFkPvfbcVCDqnBb"
FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"
AUTH_ERROR_MARKERS = ("need_user_authorization", "not configured")
EXPORT_TIMEOUT_SECONDS = int(os.getenv("XHS_FEISHU_EXPORT_TIMEOUT_SECONDS", "600"))
EXPORT_POLL_INTERVAL_SECONDS = int(os.getenv("XHS_FEISHU_EXPORT_POLL_INTERVAL_SECONDS", "5"))


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(source, output_path):
    data = {
        "source": source,
        "output_path": output_path,
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def resolve_lark_cli():
    for name in ("lark-cli.cmd", "lark-cli"):
        path = shutil.which(name)
        if path:
            return [path]

    for name in ("npx.cmd", "npx"):
        path = shutil.which(name)
        if path:
            return [path, "-y", "@larksuite/cli"]

    return []


def format_command(command):
    return " ".join(str(part) for part in command)


def run_command(command, allow_fail=False):
    result = subprocess.run(
        command,
        cwd=BASE_DIR,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    output = result.stdout or ""
    if result.returncode != 0 and not allow_fail:
        raise RuntimeError(
            "命令执行失败:\n"
            + format_command(command)
            + "\n\n"
            + output.strip()
        )
    return result.returncode, output


def cli_output_path(output_path):
    relative_path = os.path.relpath(output_path, BASE_DIR)
    if relative_path.startswith("..") or os.path.isabs(relative_path):
        raise RuntimeError(
            "lark-cli 要求输出路径必须在 data_handle 目录内，请选择 data_handle 下的 xlsx 路径。"
        )
    return "." + os.sep + relative_path


def extract_token_from_url(source, path_name):
    match = re.search(rf"/{re.escape(path_name)}/([A-Za-z0-9]+)", source)
    return match.group(1) if match else ""


def looks_like_token(source):
    return bool(re.fullmatch(r"[A-Za-z0-9_-]{12,}", source or ""))


def parse_lark_cli_json(output):
    text = (output or "").strip()
    if not text:
        return {}

    try:
        return json.loads(text)
    except Exception:
        pass

    # lark-cli may print human text before/after JSON in some versions.
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return json.loads(text[start : end + 1])
    return {}


def run_lark_cli_json(command):
    return_code, output = run_command(command, allow_fail=True)
    if return_code != 0:
        raise RuntimeError(
            "lark-cli 命令失败:\n"
            + format_command(command)
            + "\n\n"
            + output.strip()
        )
    try:
        return parse_lark_cli_json(output)
    except Exception as exc:
        raise RuntimeError(f"无法解析 lark-cli JSON 输出: {exc}\n\n{output.strip()}")


def ensure_lark_cli_ok(payload, action):
    if payload.get("ok") is False:
        raise RuntimeError(f"{action}失败:\n{json.dumps(payload, ensure_ascii=False, indent=2)}")
    if payload.get("code") not in (None, 0):
        raise RuntimeError(f"{action}失败:\n{json.dumps(payload, ensure_ascii=False, indent=2)}")


def extract_obj_token(payload):
    node = (
        payload.get("data", {}).get("node")
        or payload.get("node")
        or payload.get("data", {})
    )
    if not isinstance(node, dict):
        return ""

    obj_type = node.get("obj_type", "")
    if obj_type and obj_type != "sheet":
        raise RuntimeError(f"当前 wiki 节点不是飞书表格，obj_type={obj_type}")
    return node.get("obj_token", "") or node.get("spreadsheet_token", "")


def resolve_wiki_token_with_lark_cli(lark_cli, wiki_token):
    command = [
        *lark_cli,
        "api",
        "GET",
        "/open-apis/wiki/v2/spaces/get_node",
        "--as",
        "user",
        "--params",
        json.dumps({"token": wiki_token}, ensure_ascii=False),
        "--format",
        "json",
    ]
    return_code, output = run_command(command, allow_fail=True)
    if return_code != 0:
        return "", output

    payload = parse_lark_cli_json(output)
    return extract_obj_token(payload), output


def feishu_request(method, path, token="", params=None, body=None):
    url = f"{FEISHU_BASE_URL}{path}"
    if params:
        url = f"{url}?{urllib.parse.urlencode(params, doseq=True)}"

    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"

    data = None
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")

    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(request, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))

    if payload.get("code") not in (0, None):
        raise RuntimeError(f"飞书接口返回异常: {payload}")
    return payload.get("data", payload)


def resolve_wiki_token_with_env(wiki_token):
    app_id = os.getenv("FEISHU_APP_ID", "").strip()
    app_secret = os.getenv("FEISHU_APP_SECRET", "").strip()
    if not app_id or not app_secret:
        return ""

    token_data = feishu_request(
        "POST",
        "/auth/v3/tenant_access_token/internal",
        body={"app_id": app_id, "app_secret": app_secret},
    )
    tenant_access_token = token_data.get("tenant_access_token", "")
    if not tenant_access_token:
        return ""

    node_data = feishu_request(
        "GET",
        "/wiki/v2/spaces/get_node",
        token=tenant_access_token,
        params={"token": wiki_token},
    )
    return extract_obj_token({"data": node_data})


def resolve_spreadsheet_token(source, spreadsheet_token, wiki_token, lark_cli):
    if spreadsheet_token:
        return spreadsheet_token

    source = (source or "").strip()
    if source:
        sheet_token = extract_token_from_url(source, "sheets")
        if sheet_token:
            return sheet_token

        wiki_token = wiki_token or extract_token_from_url(source, "wiki")
        if not wiki_token and looks_like_token(source):
            return source

    if not wiki_token:
        raise RuntimeError("缺少飞书表格链接、wiki 链接、wiki token 或 spreadsheet token。")

    token, cli_output = resolve_wiki_token_with_lark_cli(lark_cli, wiki_token)
    if token:
        return token

    token = resolve_wiki_token_with_env(wiki_token)
    if token:
        return token

    raise RuntimeError(
        "无法把 wiki token 解析成 spreadsheet token。\n"
        "请确认 lark-cli 已登录且有 wiki/sheets 权限，或设置 FEISHU_APP_ID、FEISHU_APP_SECRET。\n\n"
        f"lark-cli 输出:\n{cli_output.strip()}"
    )


def normalize_output_path(output_path):
    output_path = (output_path or DEFAULT_OUTPUT_PATH).strip()
    if not output_path:
        output_path = DEFAULT_OUTPUT_PATH
    if os.path.isdir(output_path):
        output_path = os.path.join(output_path, DEFAULT_OUTPUT_NAME)
    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"
    return os.path.abspath(output_path)


def is_feishu_url(source):
    return bool(re.match(r"https?://", source or ""))


def export_spreadsheet_url(lark_cli, source_url, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    command = [
        *lark_cli,
        "sheets",
        "+export",
        "--url",
        source_url,
        "--file-extension",
        "xlsx",
        "--output-path",
        cli_output_path(output_path),
    ]
    return_code, output = run_command(command, allow_fail=True)
    if return_code == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
        return True, output
    return False, output


def is_auth_error(output):
    text = output or ""
    return any(marker in text for marker in AUTH_ERROR_MARKERS)


def build_auth_error_message(command_output):
    return (
        "飞书 CLI 还没有完成配置或用户授权。\n"
        "请回到启动器点击「飞书授权」，按日志里的验证链接完成授权后再下载。\n\n"
        f"lark-cli 输出:\n{command_output.strip()}"
    )


def export_spreadsheet(lark_cli, spreadsheet_token, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    if os.getenv("XHS_FEISHU_USE_CLI_EXPORT", "").strip() != "1":
        print("使用长轮询导出任务，避免 CLI 默认等待超时...", flush=True)
        export_spreadsheet_with_long_poll(lark_cli, spreadsheet_token, output_path)
        return

    command = [
        *lark_cli,
        "sheets",
        "+export",
        "--spreadsheet-token",
        spreadsheet_token,
        "--file-extension",
        "xlsx",
        "--output-path",
        cli_output_path(output_path),
    ]
    return_code, output = run_command(command, allow_fail=True)
    if return_code != 0:
        if "export task timed out" in output:
            print("CLI 导出等待超时，改用长轮询导出任务...", flush=True)
            export_spreadsheet_with_long_poll(lark_cli, spreadsheet_token, output_path)
            return
        raise RuntimeError("飞书表格导出失败。\n\n命令输出:\n" + output.strip())

    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(f"导出命令已结束，但没有生成有效文件: {output_path}")


def export_spreadsheet_with_long_poll(lark_cli, spreadsheet_token, output_path):
    create_payload = run_lark_cli_json(
        [
            *lark_cli,
            "api",
            "POST",
            "/open-apis/drive/v1/export_tasks",
            "--as",
            "user",
            "--data",
            json.dumps(
                {
                    "file_extension": "xlsx",
                    "token": spreadsheet_token,
                    "type": "sheet",
                },
                ensure_ascii=False,
            ),
            "--format",
            "json",
        ]
    )
    ensure_lark_cli_ok(create_payload, "创建飞书导出任务")
    ticket = (
        create_payload.get("data", {}).get("ticket")
        or create_payload.get("ticket")
        or ""
    )
    if not ticket:
        raise RuntimeError(
            "创建飞书导出任务成功，但没有返回 ticket:\n"
            + json.dumps(create_payload, ensure_ascii=False, indent=2)
        )

    print(f"导出任务已创建，ticket={ticket}", flush=True)
    deadline = time.time() + max(EXPORT_TIMEOUT_SECONDS, 60)
    last_payload = {}
    file_token = ""
    while time.time() < deadline:
        task_payload = run_lark_cli_json(
            [
                *lark_cli,
                "drive",
                "+task_result",
                "--as",
                "user",
                "--scenario",
                "export",
                "--ticket",
                ticket,
                "--file-token",
                spreadsheet_token,
            ]
        )
        ensure_lark_cli_ok(task_payload, "查询飞书导出任务")
        last_payload = task_payload
        data = task_payload.get("data") or {}
        result = data.get("result") or data or task_payload.get("result") or {}
        job_status = result.get("job_status")
        file_token = result.get("file_token") or ""
        if file_token:
            print("导出任务完成，开始下载文件...", flush=True)
            break

        if result.get("failed") or job_status not in (None, 0, 1, 2):
            raise RuntimeError(
                "飞书导出任务失败:\n"
                + json.dumps(task_payload, ensure_ascii=False, indent=2)
            )

        print("导出任务仍在处理中，继续等待...", flush=True)
        time.sleep(max(EXPORT_POLL_INTERVAL_SECONDS, 1))

    if not file_token:
        raise RuntimeError(
            "飞书导出任务长轮询超时。\n最后一次任务状态:\n"
            + json.dumps(last_payload, ensure_ascii=False, indent=2)
        )

    download_code, download_output = run_command(
        [
            *lark_cli,
            "drive",
            "+export-download",
            "--as",
            "user",
            "--file-token",
            file_token,
            "--file-name",
            os.path.basename(output_path),
            "--output-dir",
            ".",
            "--overwrite",
        ],
        allow_fail=True,
    )
    if download_code != 0:
        raise RuntimeError("下载飞书导出文件失败:\n" + download_output.strip())
    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(
            "导出文件下载命令已结束，但没有生成有效文件:\n"
            + output_path
            + "\n\nlark-cli 输出:\n"
            + download_output.strip()
        )


def build_parser():
    parser = argparse.ArgumentParser(description="从飞书表格自动导出 Excel 到 data_handle。")
    parser.add_argument("--source", "--url", dest="source", default="", help="飞书 wiki/sheets 链接，或 spreadsheet token")
    parser.add_argument("--spreadsheet-token", default="", help="直接指定飞书表格 token")
    parser.add_argument("--wiki-token", default="", help="直接指定飞书 wiki token")
    parser.add_argument("--output", default="", help="输出 xlsx 路径，默认保存到 data_handle")
    parser.add_argument("--remember", action="store_true", help="保存本次 source/output 到本地配置，后续可直接运行")
    return parser


def main():
    config = load_config()
    parser = build_parser()
    args = parser.parse_args()

    source = (
        args.source.strip()
        or os.getenv("XHS_FEISHU_SOURCE", "").strip()
        or os.getenv("XHS_FEISHU_URL", "").strip()
        or str(config.get("source", "")).strip()
        or DEFAULT_FEISHU_SOURCE
    )
    output_path = normalize_output_path(
        args.output.strip()
        or os.getenv("XHS_FEISHU_EXCEL_OUTPUT", "").strip()
        or str(config.get("output_path", "")).strip()
        or DEFAULT_OUTPUT_PATH
    )

    prompted_source = False
    if not source and not args.spreadsheet_token and not args.wiki_token:
        source = input("请输入飞书 wiki/sheets 链接或 spreadsheet token: ").strip()
        prompted_source = True

    lark_cli = resolve_lark_cli()
    if not lark_cli:
        raise RuntimeError(
            "未找到 lark-cli 或 npx。请先安装 Node.js，然后登录飞书 CLI:\n"
            "  npm install -g @larksuite/cli\n"
            "  lark-cli config init --new\n"
            "  lark-cli auth login --recommend --domain sheets,wiki,drive"
        )

    print(f"lark-cli: {format_command(lark_cli)}")
    print(f"输出文件: {output_path}")
    print("开始导出飞书表格为 xlsx...")

    spreadsheet_token = resolve_spreadsheet_token(
        source=source,
        spreadsheet_token=args.spreadsheet_token.strip(),
        wiki_token=args.wiki_token.strip(),
        lark_cli=lark_cli,
    )

    export_spreadsheet(lark_cli, spreadsheet_token, output_path)
    print(f"下载完成: {output_path}")

    if source and (args.remember or prompted_source):
        save_config(source, output_path)
        print(f"已保存本地配置: {CONFIG_PATH}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1)
