import queue
import shutil
import subprocess
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from split_csv_upload_tool import detect_date_from_name, run_pipeline


BASE_DIR = Path(__file__).resolve().parent
AUTH_SCOPE = "space:document:retrieve sheets:spreadsheet:read sheets:spreadsheet:write_only"


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
        self.root.title("xhs CSV 分片上传飞书")
        self.root.geometry("900x680")
        self.root.minsize(760, 560)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.auth_worker: threading.Thread | None = None

        self.source_var = tk.StringVar()
        self.output_root_var = tk.StringVar(value=str(BASE_DIR))
        self.date_var = tk.StringVar()
        self.chunk_size_var = tk.StringVar(value="200")
        self.min_last_rows_var = tk.StringVar(value="100")
        self.upload_var = tk.BooleanVar(value=True)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.import_as_sheet_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.root.after(150, self._drain_log_queue)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        form = ttk.Frame(self.root, padding=16)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        title = ttk.Label(form, text="xhs CSV 筛选、分片并上传飞书", font=("Microsoft YaHei UI", 15, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(form, text="源 CSV").grid(row=1, column=0, sticky="w", pady=(16, 6))
        ttk.Entry(form, textvariable=self.source_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(16, 6))
        ttk.Button(form, text="浏览", command=self._pick_source).grid(row=1, column=2, pady=(16, 6))

        ttk.Label(form, text="输出根目录").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_root_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(form, text="浏览", command=self._pick_output_root).grid(row=2, column=2, pady=6)

        options = ttk.LabelFrame(self.root, text="规则", padding=16)
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

        buttons = ttk.Frame(self.root, padding=(16, 12))
        buttons.grid(row=3, column=0, sticky="ew")
        buttons.columnconfigure(0, weight=1)
        self.auth_button = ttk.Button(buttons, text="补充飞书读取授权", command=self._start_auth)
        self.auth_button.grid(row=0, column=1, padx=(0, 10))
        self.run_button = ttk.Button(buttons, text="开始处理", command=self._start)
        self.run_button.grid(row=0, column=2)

        self.log_text = ScrolledText(self.root, height=18, font=("Consolas", 10))
        self.log_text.grid(row=2, column=0, sticky="nsew", padx=16, pady=(12, 0))

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

        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", "当前任务还在执行。")
            return

        self.run_button.configure(state="disabled")
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
                self.root.after(0, lambda: messagebox.showinfo("完成", f"任务完成。\n\n输出目录:\n{output_dir}"))
            except Exception as exc:
                error_message = str(exc)
                self._logger(f"ERROR: {error_message}")
                self.root.after(0, lambda: messagebox.showerror("处理失败", error_message))
            finally:
                self.root.after(0, lambda: self.run_button.configure(state="normal"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()


def main() -> None:
    root = tk.Tk()
    app = CsvSplitUploadApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
