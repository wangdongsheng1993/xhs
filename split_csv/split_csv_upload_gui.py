import queue
import shutil
import subprocess
import os
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from split_csv_upload_tool import (
    detect_date_from_name,
    run_pipeline,
)
from xhs_note_url_dom_updater import DEFAULT_PORT, DEFAULT_PROFILE_DIR, launch_browser, update_note_urls
from xhs_note_url_dom_updater import format_duration


BASE_DIR = Path(__file__).resolve().parent
AUTH_SCOPE = "space:document:retrieve sheets:spreadsheet:read sheets:spreadsheet:write_only"
UPDATED_DIR_NAME = "地址已更新"


def subprocess_no_window_kwargs() -> dict:
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    return {
        "creationflags": subprocess.CREATE_NO_WINDOW,
        "startupinfo": startupinfo,
    }


class CsvSplitUploadApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("xhs CSV 筛选分片 / 飞书上传 / 笔记地址更新")
        self.root.geometry("980x780")
        self.root.minsize(860, 680)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.auth_worker: threading.Thread | None = None
        self.xhs_worker: threading.Thread | None = None
        self.update_stop_event = threading.Event()
        self.update_continue_event = threading.Event()

        self.source_var = tk.StringVar()
        self.update_target_var = tk.StringVar()
        self.output_root_var = tk.StringVar(value=str(BASE_DIR))
        self.date_var = tk.StringVar()
        self.chunk_size_var = tk.StringVar(value="200")
        self.min_last_rows_var = tk.StringVar(value="100")
        self.upload_var = tk.BooleanVar(value=True)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.import_as_sheet_var = tk.BooleanVar(value=True)
        self.update_interval_var = tk.StringVar(value="20")

        self._build_ui()
        self.root.after(150, self._drain_log_queue)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        form = ttk.Frame(self.root, padding=16)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        title = ttk.Label(form, text="xhs CSV 处理工具", font=("Microsoft YaHei UI", 15, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(form, text="源 CSV").grid(row=1, column=0, sticky="w", pady=(16, 6))
        ttk.Entry(form, textvariable=self.source_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(16, 6))
        ttk.Button(form, text="浏览", command=self._pick_source).grid(row=1, column=2, pady=(16, 6))

        ttk.Label(form, text="输出根目录").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_root_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(form, text="浏览", command=self._pick_output_root).grid(row=2, column=2, pady=6)

        options = ttk.LabelFrame(self.root, text="筛选分片 / 飞书上传", padding=16)
        options.grid(row=1, column=0, sticky="ew", padx=16)
        options.columnconfigure(5, weight=1)

        ttk.Label(options, text="目录日期").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.date_var, width=12).grid(row=0, column=1, sticky="w", padx=(8, 20))
        ttk.Label(options, text="YYYYMMDD；留空按文件名解析").grid(row=0, column=2, sticky="w", padx=(0, 20))

        ttk.Label(options, text="每片行数").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(options, textvariable=self.chunk_size_var, width=8).grid(row=1, column=1, sticky="w", padx=(8, 20), pady=(12, 0))

        ttk.Label(options, text="最后分片最少行数").grid(row=1, column=2, sticky="w", pady=(12, 0))
        ttk.Entry(options, textvariable=self.min_last_rows_var, width=8).grid(row=1, column=3, sticky="w", padx=(8, 20), pady=(12, 0))

        checks = ttk.Frame(options)
        checks.grid(row=2, column=1, columnspan=5, sticky="w", pady=(12, 0))
        ttk.Checkbutton(checks, text="上传到飞书", variable=self.upload_var).grid(row=0, column=0, padx=(0, 20))
        ttk.Checkbutton(checks, text="飞书同名文件已存在时跳过", variable=self.skip_existing_var).grid(row=0, column=1)
        ttk.Checkbutton(checks, text="转为飞书电子表格", variable=self.import_as_sheet_var).grid(row=0, column=2, padx=(20, 0))

        auth_hint = ttk.Label(
            options,
            text=f'复用目录、跳过同名文件、隐藏表格列需要授权：lark-cli auth login --scope "{AUTH_SCOPE}"',
            foreground="gray",
        )
        auth_hint.grid(row=3, column=1, columnspan=5, sticky="w", pady=(12, 0))

        upload_buttons = ttk.Frame(options)
        upload_buttons.grid(row=4, column=1, columnspan=5, sticky="e", pady=(12, 0))
        self.auth_button = ttk.Button(upload_buttons, text="补充飞书读取授权", command=self._start_auth)
        self.auth_button.grid(row=0, column=0, padx=(0, 10))
        self.run_button = ttk.Button(upload_buttons, text="开始筛选分片/上传", command=self._start)
        self.run_button.grid(row=0, column=1)

        xhs_tools = ttk.LabelFrame(self.root, text="XHS 笔记地址", padding=16)
        xhs_tools.grid(row=2, column=0, sticky="ew", padx=16, pady=(12, 0))
        xhs_tools.columnconfigure(1, weight=1)

        self.open_xhs_button = ttk.Button(
            xhs_tools,
            text="打开/连接xhs浏览器",
            command=self._open_xhs_browser,
        )
        self.open_xhs_button.grid(row=0, column=0, sticky="w")
        ttk.Label(
            xhs_tools,
            text="先登录打开的浏览器；更新分支可处理单个 CSV 或整个分片目录。",
            foreground="gray",
        ).grid(row=0, column=1, columnspan=4, sticky="w", padx=(12, 0))

        ttk.Label(xhs_tools, text="更新目标").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(xhs_tools, textvariable=self.update_target_var).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(8, 8),
            pady=(12, 0),
        )
        self.pick_update_csv_button = ttk.Button(xhs_tools, text="选CSV", command=self._pick_update_csv)
        self.pick_update_csv_button.grid(row=1, column=2, sticky="w", pady=(12, 0))
        self.pick_update_folder_button = ttk.Button(xhs_tools, text="选文件夹", command=self._pick_update_folder)
        self.pick_update_folder_button.grid(row=1, column=3, sticky="w", padx=(8, 0), pady=(12, 0))
        self.update_note_url_button = ttk.Button(
            xhs_tools,
            text="更新笔记官方地址",
            command=self._update_note_urls,
        )
        self.update_note_url_button.grid(row=1, column=4, sticky="w", padx=(8, 0), pady=(12, 0))

        ttk.Label(xhs_tools, text="基础间隔秒数").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(xhs_tools, textvariable=self.update_interval_var, width=8).grid(
            row=2,
            column=1,
            sticky="w",
            padx=(8, 0),
            pady=(10, 0),
        )
        self.stop_update_button = ttk.Button(
            xhs_tools,
            text="停止更新",
            command=self._stop_update_note_urls,
            state="disabled",
        )
        self.stop_update_button.grid(row=2, column=3, sticky="w", padx=(8, 0), pady=(10, 0))
        self.continue_update_button = ttk.Button(
            xhs_tools,
            text="继续当前条",
            command=self._continue_current_note_url,
            state="disabled",
        )
        self.continue_update_button.grid(row=2, column=4, sticky="w", padx=(8, 0), pady=(10, 0))
        ttk.Label(
            xhs_tools,
            text=(
                f"更新结果写入 {UPDATED_DIR_NAME} 子目录；若已存在同名更新结果，会按最后一个有效 token 的下一行续跑，"
                "不会回头补之前失败行；"
                "主页访问会按基础间隔到 +20 秒随机等待；遇到登录/验证码页或请求频繁页时，会静默挂起当前条，"
                "不会再自动访问页面；你处理完成后点击“继续当前条”，超时才记失败。"
            ),
            foreground="gray",
            wraplength=360,
        ).grid(row=3, column=1, columnspan=4, sticky="w", padx=(8, 0), pady=(8, 0))

        self.log_text = ScrolledText(self.root, height=18, font=("Consolas", 10))
        self.log_text.grid(row=3, column=0, sticky="nsew", padx=16, pady=(12, 0))

    def _pick_source(self) -> None:
        path = filedialog.askopenfilename(
            title="选择源 CSV",
            initialdir=str(BASE_DIR),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self.source_var.set(path)
        if not self.date_var.get().strip():
            try:
                self.date_var.set(detect_date_from_name(Path(path)))
            except Exception:
                pass

    def _pick_output_root(self) -> None:
        path = filedialog.askdirectory(title="选择输出根目录", initialdir=self.output_root_var.get() or str(BASE_DIR))
        if path:
            self.output_root_var.set(path)

    def _pick_update_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要更新的分片 CSV",
            initialdir=str(BASE_DIR),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.update_target_var.set(path)

    def _pick_update_folder(self) -> None:
        path = filedialog.askdirectory(title="选择包含分片 CSV 的文件夹", initialdir=self.output_root_var.get() or str(BASE_DIR))
        if path:
            self.update_target_var.set(path)

    def _append_log(self, message: str) -> None:
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def _logger(self, message: str) -> None:
        self.log_queue.put(message)

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)
        self.root.after(150, self._drain_log_queue)

    def _validate(self) -> tuple[Path, Path, str | None, int, int, bool, bool, bool] | None:
        source_text = self.source_var.get().strip()
        if not source_text:
            messagebox.showerror("缺少源文件", "请选择源 CSV 文件。")
            return None

        source = Path(source_text)
        if not source.exists():
            messagebox.showerror("源文件不存在", str(source))
            return None

        output_root = Path(self.output_root_var.get().strip() or str(BASE_DIR))

        date_text = self.date_var.get().strip()
        if date_text and (len(date_text) != 8 or not date_text.isdigit()):
            messagebox.showerror("目录日期格式错误", "目录日期请填写 YYYYMMDD，例如 20260514。")
            return None

        try:
            chunk_size = int(self.chunk_size_var.get().strip())
            min_last_rows = int(self.min_last_rows_var.get().strip())
        except ValueError:
            messagebox.showerror("数字格式错误", "每片行数和最后分片最少行数必须是整数。")
            return None

        return (
            source,
            output_root,
            date_text or None,
            chunk_size,
            min_last_rows,
            self.upload_var.get(),
            self.skip_existing_var.get(),
            self.import_as_sheet_var.get(),
        )

    def _validate_update_targets(self) -> list[Path] | None:
        target_text = self.update_target_var.get().strip()
        if not target_text:
            messagebox.showerror("缺少更新目标", "请选择要更新的单个 CSV，或包含分片 CSV 的文件夹。")
            return None

        target = Path(target_text)
        if not target.exists():
            messagebox.showerror("更新目标不存在", str(target))
            return None

        if target.is_file():
            if target.suffix.lower() != ".csv":
                messagebox.showerror("文件类型错误", "请选择 CSV 文件。")
                return None
            if not self._is_update_csv_candidate(target):
                messagebox.showerror("文件不可更新", "请选择原始分片 CSV，不要选择跟踪、失败、筛选结果或已更新产物。")
                return None
            return [target]

        csv_files = [path for path in sorted(target.glob("*.csv")) if self._is_update_csv_candidate(path)]
        if not csv_files:
            messagebox.showerror("没有可更新的 CSV", "该文件夹下没有可处理的分片 CSV。")
            return None
        return csv_files

    def _validate_update_interval(self) -> int | None:
        interval_text = self.update_interval_var.get().strip()
        try:
            interval_sec = int(interval_text)
        except ValueError:
            messagebox.showerror("间隔格式错误", "基础间隔秒数必须是整数。")
            return None
        if interval_sec < 0:
            messagebox.showerror("间隔格式错误", "基础间隔秒数不能小于 0。")
            return None
        return interval_sec

    def _is_update_csv_candidate(self, path: Path) -> bool:
        name = path.name
        if not name.lower().endswith(".csv"):
            return False
        if UPDATED_DIR_NAME in path.parts:
            return False
        excluded_keywords = ("上传跟踪_", "失败数据", "筛选结果", "笔记地址已更新")
        return not any(keyword in name for keyword in excluded_keywords)

    def _updated_output_path_for(self, source: Path) -> Path:
        return source.parent / UPDATED_DIR_NAME / source.name

    def _is_processing(self) -> bool:
        return bool(
            (self.worker and self.worker.is_alive())
            or (self.xhs_worker and self.xhs_worker.is_alive())
        )

    def _set_xhs_buttons_state(self, state: str) -> None:
        self.open_xhs_button.configure(state=state)
        self.pick_update_csv_button.configure(state=state)
        self.pick_update_folder_button.configure(state=state)
        self.update_note_url_button.configure(state=state)
        self.stop_update_button.configure(state="disabled")
        self.continue_update_button.configure(state="disabled")

    def _stop_update_note_urls(self) -> None:
        if not (self.xhs_worker and self.xhs_worker.is_alive()):
            return
        self.update_stop_event.set()
        self.stop_update_button.configure(state="disabled")
        self.continue_update_button.configure(state="disabled")
        self._append_log("已请求停止更新；当前正在处理的页面结束后会保存已完成数据。")

    def _continue_current_note_url(self) -> None:
        if not (self.xhs_worker and self.xhs_worker.is_alive()):
            return
        self.update_continue_event.set()
        self.continue_update_button.configure(state="disabled")
        self._append_log("已请求继续当前条；脚本将基于当前浏览器页面状态恢复处理。")

    def _open_xhs_browser(self) -> None:
        if self._is_processing():
            messagebox.showinfo("正在运行", "当前任务还在执行。")
            return

        self._set_xhs_buttons_state("disabled")
        self._append_log("")
        self._append_log("打开/连接 xhs 浏览器...")

        def worker() -> None:
            try:
                launch_browser(DEFAULT_PORT, DEFAULT_PROFILE_DIR, logger=self._logger)
                self._logger("请在打开的浏览器中登录 xhs；登录后可回到 GUI 执行更新。")
                self.root.after(0, lambda: messagebox.showinfo("浏览器已就绪", "请在打开的浏览器中登录 xhs。"))
            except Exception as exc:
                error_message = str(exc)
                self._logger(f"ERROR: {error_message}")
                self.root.after(0, lambda: messagebox.showerror("打开浏览器失败", error_message))
            finally:
                self.root.after(0, lambda: self._set_xhs_buttons_state("normal"))

        self.xhs_worker = threading.Thread(target=worker, daemon=True)
        self.xhs_worker.start()

    def _update_note_urls(self) -> None:
        targets = self._validate_update_targets()
        if targets is None:
            return
        interval_sec = self._validate_update_interval()
        if interval_sec is None:
            return

        if self._is_processing():
            messagebox.showinfo("正在运行", "当前任务还在执行。")
            return

        self.update_stop_event.clear()
        self.update_continue_event.clear()
        self._set_xhs_buttons_state("disabled")
        self.stop_update_button.configure(state="normal")
        self.continue_update_button.configure(state="normal")
        self.run_button.configure(state="disabled")
        self._append_log("")
        self._append_log("开始更新笔记官方地址...")

        def worker() -> None:
            try:
                batch_started_at = time.monotonic()
                failed_files = []
                output_files = []
                for index, target in enumerate(targets, start=1):
                    csv_started_at = time.monotonic()
                    self._logger(f"开始更新 CSV {index}/{len(targets)}: {target}")
                    output_csv = self._updated_output_path_for(target)
                    self._logger(f"更新结果输出到: {output_csv}")
                    if output_csv.exists():
                        self._logger("检测到同名更新结果，自动续跑未完成部分。")
                    saved_csv, failed_csv = update_note_urls(
                        source=target,
                        output=output_csv,
                        port=DEFAULT_PORT,
                        profile_dir=DEFAULT_PROFILE_DIR,
                        skip_existing_xsec=True,
                        interval_sec=interval_sec,
                        max_scrolls=0,
                        logger=self._logger,
                        stop_event=self.update_stop_event,
                        continue_event=self.update_continue_event,
                    )
                    output_files.append(saved_csv)
                    if failed_csv:
                        failed_files.append(failed_csv)
                    self._logger(f"CSV {index}/{len(targets)} 耗时: {format_duration(time.monotonic() - csv_started_at)}")
                    if self.update_stop_event.is_set():
                        self._logger("停止请求已生效，不再处理后续 CSV。")
                        break

                output_dir = output_files[0].parent if output_files else targets[0].parent / UPDATED_DIR_NAME
                self.root.after(0, lambda path=output_dir: self.update_target_var.set(str(path)))
                processed_count = len(output_files)
                status_text = "已停止" if self.update_stop_event.is_set() else "更新完成"
                batch_duration = format_duration(time.monotonic() - batch_started_at)
                self._logger(f"{status_text}，共处理 {processed_count}/{len(targets)} 个 CSV，总耗时: {batch_duration}。")
                self._logger(f"更新结果目录: {output_dir}")
                for failed_csv in failed_files:
                    self._logger(f"失败数据 CSV: {failed_csv}")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        status_text,
                        f"已处理 {processed_count}/{len(targets)} 个 CSV。\n总耗时: {batch_duration}\n\n更新结果目录:\n{output_dir}",
                    ),
                )
            except Exception as exc:
                error_message = str(exc)
                self._logger(f"ERROR: {error_message}")
                self.root.after(0, lambda: messagebox.showerror("更新失败", error_message))
            finally:
                self.root.after(0, lambda: self._set_xhs_buttons_state("normal"))
                self.root.after(0, lambda: self.run_button.configure(state="normal"))
                self.root.after(0, lambda: self.stop_update_button.configure(state="disabled"))
                self.root.after(0, lambda: self.continue_update_button.configure(state="disabled"))

        self.xhs_worker = threading.Thread(target=worker, daemon=True)
        self.xhs_worker.start()

    def _start_auth(self) -> None:
        if self.auth_worker and self.auth_worker.is_alive():
            messagebox.showinfo("正在授权", "飞书授权任务还在执行。")
            return

        lark_cli = shutil.which("lark-cli")
        if not lark_cli:
            messagebox.showerror("未找到 lark-cli", "未找到 lark-cli，请先安装并确认它在 PATH 中。")
            return

        self.auth_button.configure(state="disabled")
        self._append_log("")
        self._append_log("开始补充飞书读取/表格授权...")
        self._append_log(f'命令: lark-cli auth login --scope "{AUTH_SCOPE}"')

        def worker() -> None:
            try:
                process = subprocess.Popen(
                    [lark_cli, "auth", "login", "--scope", AUTH_SCOPE],
                    cwd=str(BASE_DIR),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    **subprocess_no_window_kwargs(),
                )
                assert process.stdout is not None
                for line in process.stdout:
                    self._logger(line.rstrip())
                returncode = process.wait()
                if returncode == 0:
                    self._logger("飞书授权完成。现在可以复用目录、跳过同名文件并隐藏表格列。")
                    self.root.after(0, lambda: messagebox.showinfo("授权完成", "飞书授权完成。"))
                else:
                    self._logger(f"ERROR: 飞书授权命令退出码 {returncode}")
                    self.root.after(0, lambda: messagebox.showerror("授权失败", f"飞书授权命令退出码 {returncode}"))
            except Exception as exc:
                error_message = str(exc)
                self._logger(f"ERROR: {error_message}")
                self.root.after(0, lambda: messagebox.showerror("授权失败", error_message))
            finally:
                self.root.after(0, lambda: self.auth_button.configure(state="normal"))

        self.auth_worker = threading.Thread(target=worker, daemon=True)
        self.auth_worker.start()

    def _start(self) -> None:
        params = self._validate()
        if params is None:
            return

        if self._is_processing():
            messagebox.showinfo("正在运行", "当前任务还在执行。")
            return

        self.run_button.configure(state="disabled")
        self._set_xhs_buttons_state("disabled")
        self._append_log("")
        self._append_log("开始处理...")

        def worker() -> None:
            try:
                source, output_root, folder_date, chunk_size, min_last_rows, upload, skip_existing, import_as_sheet = params
                output_dir, tracking_csv, tracking_md, _split_files = run_pipeline(
                    source=source,
                    output_root=output_root,
                    upload=upload,
                    folder_date=folder_date,
                    chunk_size=chunk_size,
                    min_last_rows=min_last_rows,
                    skip_existing=skip_existing,
                    import_as_sheet=import_as_sheet,
                    logger=self._logger,
                )
                self._logger(f"任务完成，本地目录: {output_dir}")
                self._logger(f"跟踪清单: {tracking_csv}")
                self._logger(f"跟踪文档: {tracking_md}")
                self.root.after(0, lambda path=output_dir: self.update_target_var.set(str(path)))
                self.root.after(0, lambda: messagebox.showinfo("完成", f"任务完成。\n\n输出目录:\n{output_dir}"))
            except Exception as exc:
                error_message = str(exc)
                self._logger(f"ERROR: {error_message}")
                self.root.after(0, lambda: messagebox.showerror("处理失败", error_message))
            finally:
                self.root.after(0, lambda: self.run_button.configure(state="normal"))
                self.root.after(0, lambda: self._set_xhs_buttons_state("normal"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()


def main() -> None:
    root = tk.Tk()
    app = CsvSplitUploadApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
