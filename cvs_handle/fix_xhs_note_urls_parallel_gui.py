import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, "fix_xhs_note_urls_parallel.py")
RETRY_SCRIPT_PATH = os.path.join(BASE_DIR, "retry_failed_parallel.py")
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "小红书笔记列表_规范Excel版.xlsx")
XLSX_FILE_TYPES = [("Excel/CSV files", "*.xlsx *.csv"), ("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
DEFAULT_TARGET_DATE = (datetime.now() - timedelta(days=3)).strftime("%m.%d")


def resolve_cli_python():
    current = sys.executable
    folder = os.path.dirname(current)
    name = os.path.basename(current).lower()
    if name == "pythonw.exe":
        candidate = os.path.join(folder, "python.exe")
        if os.path.exists(candidate):
            return candidate
    return current


def kill_process_tree(pid):
    if sys.platform == "win32":
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                timeout=10,
            )
        except Exception:
            pass


class ParallelNoteUrlFixerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书笔记官方地址修复 - 并行版")
        self.root.geometry("920x800")
        self.root.minsize(800, 600)

        self.process = None
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.excel_path_var = tk.StringVar(value=DEFAULT_EXCEL_PATH)
        self.output_path_var = tk.StringVar(value=DEFAULT_EXCEL_PATH)
        self.sheet_var = tk.StringVar(value="小红书笔记列表")
        self.rows_var = tk.StringVar()
        self.workers_var = tk.StringVar(value="1")
        self.per_item_timeout_var = tk.StringVar(value="20")
        self.max_scrolls_var = tk.StringVar(value="15")
        self.login_wait_var = tk.StringVar(value="50")
        self.only_empty_var = tk.BooleanVar(value=False)
        self.headless_var = tk.BooleanVar(value=False)
        self.skip_no_title_var = tk.BooleanVar(value=True)
        self.target_date_var = tk.StringVar(value=DEFAULT_TARGET_DATE)
        self.batch_size_var = tk.StringVar(value="30")
        self.batch_interval_var = tk.StringVar(value="30")
        self.sessions_var = tk.StringVar(value="")
        self.session_mode_var = tk.StringVar(value="rotate")

        self.retry_csv_var = tk.StringVar()
        self.retry_failed_var = tk.StringVar()
        self.retry_workers_var = tk.StringVar(value="1")
        self.retry_per_item_timeout_var = tk.StringVar(value="20")
        self.retry_max_scrolls_var = tk.StringVar(value="15")
        self.retry_login_wait_var = tk.StringVar(value="50")
        self.retry_skip_no_title_var = tk.BooleanVar(value=True)
        self.retry_target_date_var = tk.StringVar(value=DEFAULT_TARGET_DATE)
        self.retry_session_mode_var = tk.StringVar(value="rotate")

        self._build_ui()
        self.root.after(150, self._drain_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        form = ttk.Frame(self.root, padding=16)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        title = ttk.Label(form, text="小红书笔记官方地址修复 - 并行版", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(form, text="Excel/CSV").grid(row=1, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(form, textvariable=self.excel_path_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(14, 6))
        ttk.Button(form, text="浏览", command=self._pick_excel).grid(row=1, column=2, sticky="ew", pady=(14, 6))

        ttk.Label(form, text="输出").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_path_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(form, text="浏览", command=self._pick_output).grid(row=2, column=2, sticky="ew", pady=6)

        settings = ttk.LabelFrame(self.root, text="并行设置", padding=16)
        settings.grid(row=1, column=0, sticky="ew", padx=16)
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Sheet").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.sheet_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(settings, text="行号").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(settings, textvariable=self.rows_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(12, 0))
        ttk.Label(settings, text="留空处理全部；支持 2 / 2-20 / 2,5,9-12").grid(
            row=2, column=1, sticky="w", padx=(8, 0), pady=(4, 0)
        )

        row1 = ttk.Frame(settings)
        row1.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(12, 0))

        ttk.Label(row1, text="并行进程").grid(row=0, column=0, sticky="w")
        ttk.Entry(row1, textvariable=self.workers_var, width=5).grid(row=0, column=1, padx=(6, 4))
        ttk.Label(row1, text="个").grid(row=0, column=2, padx=(0, 16))

        ttk.Label(row1, text="单条超时").grid(row=0, column=3, sticky="w")
        ttk.Entry(row1, textvariable=self.per_item_timeout_var, width=5).grid(row=0, column=4, padx=(6, 4))
        ttk.Label(row1, text="秒").grid(row=0, column=5, padx=(0, 16))

        ttk.Label(row1, text="最多滚动").grid(row=0, column=6, sticky="w")
        ttk.Entry(row1, textvariable=self.max_scrolls_var, width=5).grid(row=0, column=7, padx=(6, 4))
        ttk.Label(row1, text="次").grid(row=0, column=8, padx=(0, 16))

        ttk.Label(row1, text="登录等待").grid(row=0, column=9, sticky="w")
        ttk.Entry(row1, textvariable=self.login_wait_var, width=5).grid(row=0, column=10, padx=(6, 4))
        ttk.Label(row1, text="秒").grid(row=0, column=11)

        row2 = ttk.Frame(settings)
        row2.grid(row=4, column=1, sticky="w", padx=(8, 0), pady=(12, 0))

        ttk.Checkbutton(row2, text="只处理空地址", variable=self.only_empty_var).grid(row=0, column=0, padx=(0, 16))
        ttk.Checkbutton(row2, text="无头模式", variable=self.headless_var).grid(row=0, column=1, padx=(0, 16))
        ttk.Checkbutton(row2, text="跳过无标题笔记（推荐）", variable=self.skip_no_title_var).grid(row=0, column=2, padx=(0, 16))
        
        ttk.Label(row2, text="目标日期").grid(row=0, column=3, sticky="w")
        ttk.Entry(row2, textvariable=self.target_date_var, width=10).grid(row=0, column=4, padx=(6, 0))
        ttk.Label(row2, text="（如 5-7）", foreground="gray").grid(row=0, column=5, padx=(4, 0))

        row3 = ttk.Frame(settings)
        row3.grid(row=5, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        
        ttk.Label(row3, text="每").grid(row=0, column=0)
        ttk.Entry(row3, textvariable=self.batch_size_var, width=5).grid(row=0, column=1, padx=(4, 0))
        ttk.Label(row3, text="条休息").grid(row=0, column=2, padx=(0, 0))
        ttk.Entry(row3, textvariable=self.batch_interval_var, width=5).grid(row=0, column=3, padx=(4, 0))
        ttk.Label(row3, text="秒（防风控，0=不休息）", foreground="gray").grid(row=0, column=4, padx=(4, 0))

        row4 = ttk.Frame(settings)
        row4.grid(row=6, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        
        ttk.Label(row4, text="多账号Session").grid(row=0, column=0, sticky="w")
        ttk.Entry(row4, textvariable=self.sessions_var, width=40).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(row4, text="管理Session", command=self._open_session_manager).grid(row=0, column=2, padx=(6, 0))

        row5 = ttk.Frame(settings)
        row5.grid(row=7, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        
        ttk.Label(row5, text="账号模式").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(row5, text="轮询(所有任务共享所有账号)", variable=self.session_mode_var, value="rotate").grid(row=0, column=1, padx=(6, 0))
        ttk.Radiobutton(row5, text="绑定(每个任务分配账号组)", variable=self.session_mode_var, value="bind").grid(row=0, column=2, padx=(16, 0))

        hint = ttk.Label(settings, text="提示：轮询模式-所有任务共享所有账号；绑定模式-账号平均分配给各任务，任务内轮询。", foreground="gray")
        hint.grid(row=8, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        retry_frame = ttk.LabelFrame(self.root, text="并行重试失败数据", padding=16)
        retry_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(8, 0))
        retry_frame.columnconfigure(1, weight=1)

        ttk.Label(retry_frame, text="结果CSV").grid(row=0, column=0, sticky="w")
        ttk.Entry(retry_frame, textvariable=self.retry_csv_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(retry_frame, text="浏览", command=self._pick_retry_csv).grid(row=0, column=2, sticky="ew")

        ttk.Label(retry_frame, text="失败数据").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(retry_frame, textvariable=self.retry_failed_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))
        ttk.Button(retry_frame, text="浏览", command=self._pick_retry_failed).grid(row=1, column=2, sticky="ew", pady=(8, 0))

        retry_row = ttk.Frame(retry_frame)
        retry_row.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        ttk.Label(retry_row, text="并行进程").grid(row=0, column=0, sticky="w")
        ttk.Entry(retry_row, textvariable=self.retry_workers_var, width=5).grid(row=0, column=1, padx=(6, 4))
        ttk.Label(retry_row, text="个").grid(row=0, column=2, padx=(0, 16))

        ttk.Label(retry_row, text="单条超时").grid(row=0, column=3, sticky="w")
        ttk.Entry(retry_row, textvariable=self.retry_per_item_timeout_var, width=5).grid(row=0, column=4, padx=(6, 4))
        ttk.Label(retry_row, text="秒").grid(row=0, column=5, padx=(0, 16))

        ttk.Label(retry_row, text="最多滚动").grid(row=0, column=6, sticky="w")
        ttk.Entry(retry_row, textvariable=self.retry_max_scrolls_var, width=5).grid(row=0, column=7, padx=(6, 4))
        ttk.Label(retry_row, text="次").grid(row=0, column=8, padx=(0, 16))

        ttk.Label(retry_row, text="登录等待").grid(row=0, column=9, sticky="w")
        ttk.Entry(retry_row, textvariable=self.retry_login_wait_var, width=5).grid(row=0, column=10, padx=(6, 4))
        ttk.Label(retry_row, text="秒").grid(row=0, column=11)

        retry_row2 = ttk.Frame(retry_frame)
        retry_row2.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(4, 0))

        ttk.Checkbutton(retry_row2, text="跳过无标题笔记（推荐）", variable=self.retry_skip_no_title_var).grid(row=0, column=0, padx=(0, 16))
        
        ttk.Label(retry_row2, text="目标日期").grid(row=0, column=1, sticky="w")
        ttk.Entry(retry_row2, textvariable=self.retry_target_date_var, width=10).grid(row=0, column=2, padx=(6, 0))
        ttk.Label(retry_row2, text="（如 5-7，检测到更早日期时提前停止）", foreground="gray").grid(row=0, column=3, padx=(4, 0))

        retry_hint = ttk.Label(retry_frame, text="选择之前执行的结果CSV和对应的失败数据CSV，对失败行重新查找并更新回结果CSV。", foreground="gray")
        retry_hint.grid(row=4, column=1, sticky="w", padx=(8, 0), pady=(4, 0))

        self.retry_button = ttk.Button(retry_frame, text="并行重试失败数据", command=self._start_retry)
        self.retry_button.grid(row=4, column=2, sticky="e", pady=(4, 0))

        actions = ttk.Frame(self.root, padding=(16, 12))
        actions.grid(row=4, column=0, sticky="ew")
        actions.columnconfigure(2, weight=1)

        self.run_button = ttk.Button(actions, text="开始并行处理", command=self._start_run)
        self.run_button.grid(row=0, column=0, sticky="w")
        self.stop_button = ttk.Button(actions, text="停止", command=self._stop_run, state="disabled")
        self.stop_button.grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Button(actions, text="打开文件夹", command=self._open_folder).grid(row=0, column=3, sticky="e")

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=16)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel/CSV",
            initialdir=BASE_DIR,
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.excel_path_var.set(path)
        folder = os.path.dirname(path)
        name = os.path.splitext(os.path.basename(path))[0]
        output_path = os.path.join(folder, f"{name}_结果.csv")
        self.output_path_var.set(output_path)

    def _pick_output(self):
        current = self.output_path_var.get().strip() or self.excel_path_var.get().strip() or DEFAULT_EXCEL_PATH
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            initialdir=os.path.dirname(current) if current else BASE_DIR,
            initialfile=os.path.basename(current) if current else "",
            defaultextension=".csv",
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.output_path_var.set(path)

    def _pick_retry_csv(self):
        path = filedialog.askopenfilename(
            title="选择结果CSV",
            initialdir=BASE_DIR,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.retry_csv_var.set(path)

    def _pick_retry_failed(self):
        path = filedialog.askopenfilename(
            title="选择失败数据CSV",
            initialdir=BASE_DIR,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.retry_failed_var.set(path)

    def _open_session_manager(self):
        session_manager_path = os.path.join(BASE_DIR, "session_manager_gui.py")
        if not os.path.exists(session_manager_path):
            messagebox.showerror("错误", f"Session管理器不存在: {session_manager_path}")
            return
        
        python_exe = resolve_cli_python()
        try:
            subprocess.Popen([python_exe, session_manager_path])
        except Exception as e:
            messagebox.showerror("错误", f"启动Session管理器失败: {e}")

    def _validate_form(self):
        excel_path = self.excel_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        workers = self.workers_var.get().strip()
        per_item_timeout = self.per_item_timeout_var.get().strip()
        max_scrolls = self.max_scrolls_var.get().strip()
        login_wait = self.login_wait_var.get().strip()

        if not excel_path:
            raise ValueError("请选择 Excel/CSV。")
        if not os.path.exists(excel_path):
            raise ValueError("文件不存在。")
        if not output_path:
            raise ValueError("请填写输出路径。")
        if not sheet_name:
            raise ValueError("请填写 Sheet 名。")
        if not workers.isdigit() or int(workers) < 1:
            raise ValueError("并行进程数必须是正整数。")
        if not per_item_timeout.isdigit() or int(per_item_timeout) < 1:
            raise ValueError("单条超时必须是正整数。")
        if not max_scrolls.isdigit():
            raise ValueError("最多滚动次数必须是整数。")
        if not login_wait.isdigit():
            raise ValueError("登录等待必须是整数秒。")
        return excel_path, output_path, sheet_name, workers, per_item_timeout, max_scrolls, login_wait

    def _validate_retry_form(self):
        retry_csv = self.retry_csv_var.get().strip()
        retry_failed = self.retry_failed_var.get().strip()
        retry_workers = self.retry_workers_var.get().strip()
        retry_per_item_timeout = self.retry_per_item_timeout_var.get().strip()
        retry_max_scrolls = self.retry_max_scrolls_var.get().strip()
        retry_login_wait = self.retry_login_wait_var.get().strip()

        if not retry_csv:
            raise ValueError("请选择结果CSV。")
        if not os.path.exists(retry_csv):
            raise ValueError("结果CSV文件不存在。")
        if not retry_failed:
            raise ValueError("请选择失败数据CSV。")
        if not os.path.exists(retry_failed):
            raise ValueError("失败数据CSV文件不存在。")
        if not retry_workers.isdigit() or int(retry_workers) < 1:
            raise ValueError("并行进程数必须是正整数。")
        if not retry_per_item_timeout.isdigit() or int(retry_per_item_timeout) < 1:
            raise ValueError("单条超时必须是正整数。")
        if not retry_max_scrolls.isdigit():
            raise ValueError("最多滚动次数必须是整数。")
        if not retry_login_wait.isdigit():
            raise ValueError("登录等待必须是整数秒。")
        return retry_csv, retry_failed, retry_workers, retry_per_item_timeout, retry_max_scrolls, retry_login_wait

    def _build_command(self):
        excel_path, output_path, sheet_name, workers, per_item_timeout, max_scrolls, login_wait = self._validate_form()
        command = [
            resolve_cli_python(),
            "-u",
            SCRIPT_PATH,
            "--excel",
            excel_path,
            "--output",
            output_path,
            "--sheet",
            sheet_name,
            "--workers",
            workers,
            "--per-item-timeout",
            per_item_timeout,
            "--max-scrolls",
            max_scrolls,
            "--login-wait",
            login_wait,
        ]
        rows = self.rows_var.get().strip()
        if rows:
            command.extend(["--rows", rows])
        if self.only_empty_var.get():
            command.append("--only-empty")
        if self.headless_var.get():
            command.append("--headless")
        
        if not self.skip_no_title_var.get():
            command.append("--no-skip-no-title")
        
        target_date = self.target_date_var.get().strip()
        if target_date:
            command.extend(["--target-date", target_date])
        
        batch_size = self.batch_size_var.get().strip()
        if batch_size:
            command.extend(["--batch-size", batch_size])
        batch_interval = self.batch_interval_var.get().strip()
        if batch_interval:
            command.extend(["--batch-interval", batch_interval])
        
        sessions = self.sessions_var.get().strip()
        if sessions:
            command.extend(["--sessions", sessions])
        
        session_mode = self.session_mode_var.get().strip()
        if session_mode:
            command.extend(["--session-mode", session_mode])
        
        return command

    def _build_retry_command(self):
        retry_csv, retry_failed, retry_workers, retry_per_item_timeout, retry_max_scrolls, retry_login_wait = self._validate_retry_form()
        command = [
            resolve_cli_python(),
            "-u",
            RETRY_SCRIPT_PATH,
            "--result-csv",
            retry_csv,
            "--failed-csv",
            retry_failed,
            "--workers",
            retry_workers,
            "--per-item-timeout",
            retry_per_item_timeout,
            "--max-scrolls",
            retry_max_scrolls,
            "--login-wait",
            retry_login_wait,
        ]
        if not self.retry_skip_no_title_var.get():
            command.append("--no-skip-no-title")
        
        target_date = self.retry_target_date_var.get().strip()
        if target_date:
            command.extend(["--target-date", target_date])
        
        batch_size = self.batch_size_var.get().strip()
        if batch_size:
            command.extend(["--batch-size", batch_size])
        batch_interval = self.batch_interval_var.get().strip()
        if batch_interval:
            command.extend(["--batch-interval", batch_interval])
        
        sessions = self.sessions_var.get().strip()
        if sessions:
            command.extend(["--sessions", sessions])
        
        session_mode = self.session_mode_var.get().strip()
        if session_mode:
            command.extend(["--session-mode", session_mode])
        
        return command

    def _start_run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        try:
            command = self._build_command()
        except ValueError as exc:
            messagebox.showerror("参数不完整", str(exc))
            return

        if os.path.abspath(self.excel_path_var.get().strip()) == os.path.abspath(self.output_path_var.get().strip()):
            confirmed = messagebox.askyesno(
                "确认覆盖",
                "输出路径与输入文件相同。脚本会先自动备份原文件，再更新原文件。\n是否继续？",
            )
            if not confirmed:
                return

        self._clear_log()
        self._append_log("命令：\n")
        self._append_log(" ".join(f'"{part}"' if " " in part else part for part in command) + "\n\n")
        self._set_running(True)

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["PYTHONUNBUFFERED"] = "1"

        self.worker_thread = threading.Thread(
            target=self._run_command_worker,
            args=(command, env),
            daemon=True,
        )
        self.worker_thread.start()

    def _start_retry(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        try:
            command = self._build_retry_command()
        except ValueError as exc:
            messagebox.showerror("参数不完整", str(exc))
            return

        self._clear_log()
        self._append_log("并行重试失败数据命令：\n")
        self._append_log(" ".join(f'"{part}"' if " " in part else part for part in command) + "\n\n")
        self._set_running(True)

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["PYTHONUNBUFFERED"] = "1"

        self.worker_thread = threading.Thread(
            target=self._run_command_worker,
            args=(command, env),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_command_worker(self, command, env):
        try:
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
            self.process = subprocess.Popen(
                command,
                cwd=BASE_DIR,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creationflags,
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.log_queue.put(("log", line))
            return_code = self.process.wait()
            self.log_queue.put(("done", return_code))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))
        finally:
            self.process = None

    def _stop_run(self):
        if self.process and self.process.poll() is None:
            kill_process_tree(self.process.pid)
            self._append_log("\n已停止所有进程。\n")

    def _set_running(self, running):
        self.run_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.run_button.configure(text="处理中..." if running else "开始并行处理")
        self.retry_button.configure(state="disabled" if running else "normal")

    def _append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _drain_log_queue(self):
        while True:
            try:
                event_type, payload = self.log_queue.get_nowait()
            except queue.Empty:
                break

            if event_type == "log":
                self._append_log(payload)
            elif event_type == "done":
                self._set_running(False)
                if payload == 0:
                    self._append_log("\n运行完成。\n")
                    messagebox.showinfo("完成", "处理完成。")
                else:
                    self._append_log(f"\n运行失败，退出码：{payload}\n")
                    messagebox.showerror("失败", f"运行失败，退出码：{payload}")
            elif event_type == "error":
                self._set_running(False)
                self._append_log(f"\n启动失败：{payload}\n")
                messagebox.showerror("启动失败", payload)

        self.root.after(150, self._drain_log_queue)

    def _open_folder(self):
        target = self.output_path_var.get().strip() or self.excel_path_var.get().strip() or BASE_DIR
        folder = os.path.dirname(target) if os.path.splitext(target)[1] else target
        if not os.path.isdir(folder):
            folder = BASE_DIR
        os.startfile(folder)

    def _on_close(self):
        if self.process and self.process.poll() is None:
            confirmed = messagebox.askyesno("确认退出", "脚本还在运行，确定关闭窗口吗？")
            if not confirmed:
                return
            kill_process_tree(self.process.pid)
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    ParallelNoteUrlFixerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
