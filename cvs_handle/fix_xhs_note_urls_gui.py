import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, "fix_xhs_note_urls.py")
DEFAULT_EXCEL_PATH = os.path.join(BASE_DIR, "小红书笔记列表_规范Excel版.xlsx")
XLSX_FILE_TYPES = [("Excel files", "*.xlsx"), ("All files", "*.*")]


def resolve_cli_python():
    current = sys.executable
    folder = os.path.dirname(current)
    name = os.path.basename(current).lower()
    if name == "pythonw.exe":
        candidate = os.path.join(folder, "python.exe")
        if os.path.exists(candidate):
            return candidate
    return current


class NoteUrlFixerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书笔记官方地址修复")
        self.root.geometry("900x640")
        self.root.minsize(780, 560)

        self.process = None
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.excel_path_var = tk.StringVar(value=DEFAULT_EXCEL_PATH)
        self.output_path_var = tk.StringVar(value=DEFAULT_EXCEL_PATH)
        self.sheet_var = tk.StringVar(value="小红书笔记列表")
        self.rows_var = tk.StringVar()
        self.login_wait_var = tk.StringVar(value="20")
        self.max_scrolls_var = tk.StringVar(value="28")
        self.only_empty_var = tk.BooleanVar(value=False)
        self.headless_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.root.after(150, self._drain_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        form = ttk.Frame(self.root, padding=16)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        title = ttk.Label(form, text="小红书笔记官方地址修复", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(form, text="Excel").grid(row=1, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(form, textvariable=self.excel_path_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(14, 6))
        ttk.Button(form, text="浏览", command=self._pick_excel).grid(row=1, column=2, sticky="ew", pady=(14, 6))

        ttk.Label(form, text="输出").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_path_var).grid(row=2, column=1, sticky="ew", padx=8, pady=6)
        ttk.Button(form, text="浏览", command=self._pick_output).grid(row=2, column=2, sticky="ew", pady=6)

        settings = ttk.LabelFrame(self.root, text="执行设置", padding=16)
        settings.grid(row=1, column=0, sticky="ew", padx=16)
        settings.columnconfigure(1, weight=1)

        ttk.Label(settings, text="Sheet").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.sheet_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(settings, text="行号").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(settings, textvariable=self.rows_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(12, 0))
        ttk.Label(settings, text="留空处理全部；支持 2 / 2-20 / 2,5,9-12").grid(
            row=2, column=1, sticky="w", padx=(8, 0), pady=(4, 0)
        )

        option_row = ttk.Frame(settings)
        option_row.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(12, 0))

        ttk.Label(option_row, text="登录等待").grid(row=0, column=0, sticky="w")
        ttk.Entry(option_row, textvariable=self.login_wait_var, width=7).grid(row=0, column=1, padx=(6, 4))
        ttk.Label(option_row, text="秒").grid(row=0, column=2, padx=(0, 14))

        ttk.Label(option_row, text="最多滚动").grid(row=0, column=3, sticky="w")
        ttk.Entry(option_row, textvariable=self.max_scrolls_var, width=7).grid(row=0, column=4, padx=(6, 4))
        ttk.Label(option_row, text="次").grid(row=0, column=5, padx=(0, 14))

        ttk.Checkbutton(option_row, text="只处理空地址", variable=self.only_empty_var).grid(row=0, column=6, padx=(0, 14))
        ttk.Checkbutton(option_row, text="无头模式", variable=self.headless_var).grid(row=0, column=7)

        actions = ttk.Frame(self.root, padding=(16, 12))
        actions.grid(row=3, column=0, sticky="ew")
        actions.columnconfigure(2, weight=1)

        self.run_button = ttk.Button(actions, text="开始更新", command=self._start_run)
        self.run_button.grid(row=0, column=0, sticky="w")
        self.stop_button = ttk.Button(actions, text="停止", command=self._stop_run, state="disabled")
        self.stop_button.grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Button(actions, text="打开文件夹", command=self._open_folder).grid(row=0, column=3, sticky="e")

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=16)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=16, pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel",
            initialdir=BASE_DIR,
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.excel_path_var.set(path)
        self.output_path_var.set(path)

    def _pick_output(self):
        current = self.output_path_var.get().strip() or self.excel_path_var.get().strip() or DEFAULT_EXCEL_PATH
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel",
            initialdir=os.path.dirname(current) if current else BASE_DIR,
            initialfile=os.path.basename(current) if current else "",
            defaultextension=".xlsx",
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.output_path_var.set(path)

    def _validate_form(self):
        excel_path = self.excel_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        login_wait = self.login_wait_var.get().strip()
        max_scrolls = self.max_scrolls_var.get().strip()

        if not excel_path:
            raise ValueError("请选择 Excel。")
        if not os.path.exists(excel_path):
            raise ValueError("Excel 文件不存在。")
        if not output_path:
            raise ValueError("请填写输出路径。")
        if not sheet_name:
            raise ValueError("请填写 Sheet 名。")
        if not login_wait.isdigit():
            raise ValueError("登录等待必须是整数秒。")
        if not max_scrolls.isdigit():
            raise ValueError("最多滚动次数必须是整数。")
        return excel_path, output_path, sheet_name, login_wait, max_scrolls

    def _build_command(self):
        excel_path, output_path, sheet_name, login_wait, max_scrolls = self._validate_form()
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
            "--login-wait",
            login_wait,
            "--max-scrolls",
            max_scrolls,
        ]
        rows = self.rows_var.get().strip()
        if rows:
            command.extend(["--rows", rows])
        if self.only_empty_var.get():
            command.append("--only-empty")
        if self.headless_var.get():
            command.append("--headless")
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
                "输出路径与输入 Excel 相同。脚本会先自动备份原文件，再更新原文件。\n是否继续？",
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

    def _run_command_worker(self, command, env):
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
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
            self.process.terminate()
            self._append_log("\n已请求停止。\n")

    def _set_running(self, running):
        self.run_button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.run_button.configure(text="更新中..." if running else "开始更新")

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
                    messagebox.showinfo("完成", "笔记官方地址更新完成。")
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
            self.process.terminate()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    NoteUrlFixerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
