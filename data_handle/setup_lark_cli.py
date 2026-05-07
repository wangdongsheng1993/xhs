import shutil
import subprocess
import sys


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


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


def run_capture(command):
    result = subprocess.run(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    output = result.stdout or ""
    if output.strip():
        print(output.strip(), flush=True)
    return result.returncode, output


def run_stream(command):
    print("\n执行命令:", format_command(command), flush=True)
    process = subprocess.Popen(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        bufsize=1,
    )
    assert process.stdout is not None
    for line in process.stdout:
        print(line, end="", flush=True)
    return process.wait()


def main():
    lark_cli = resolve_lark_cli()
    if not lark_cli:
        raise RuntimeError(
            "未找到 lark-cli 或 npx。请先安装 Node.js，再执行 npm install -g @larksuite/cli。"
        )

    print(f"lark-cli: {format_command(lark_cli)}", flush=True)
    print("检查飞书 CLI 配置状态...", flush=True)
    status_code, status_output = run_capture([*lark_cli, "auth", "status"])

    if status_code != 0 and "not configured" in status_output:
        print("\n当前还没有配置飞书 CLI，开始创建/绑定应用。", flush=True)
        print("日志里出现验证链接后，请复制到浏览器打开并完成确认。", flush=True)
        init_code = run_stream([*lark_cli, "config", "init", "--new"])
        if init_code != 0:
            raise RuntimeError(f"飞书 CLI 配置失败，退出码: {init_code}")
    else:
        print("\n已检测到飞书 CLI 配置，跳过 config init。", flush=True)

    print("\n开始用户授权。", flush=True)
    print("日志里出现验证链接后，请复制到浏览器打开并完成授权。", flush=True)
    login_code = run_stream(
        [
            *lark_cli,
            "auth",
            "login",
            "--recommend",
            "--domain",
            "sheets,wiki,drive",
        ]
    )
    if login_code != 0:
        raise RuntimeError(f"飞书用户授权失败，退出码: {login_code}")

    print("\n授权完成，正在确认状态...", flush=True)
    final_code, _final_output = run_capture([*lark_cli, "auth", "status"])
    if final_code != 0:
        raise RuntimeError(f"授权状态检查失败，退出码: {final_code}")
    print("\n飞书 CLI 已可用。", flush=True)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr, flush=True)
        raise SystemExit(1)
