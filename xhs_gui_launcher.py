import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNNER_PATH = os.path.join(BASE_DIR, "xhs_excel_runner.py")

MODE_OPTIONS = [
    ("品牌", "brand"),
    ("电商", "ecommerce"),
    ("KOC", "koc"),
]

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


def derive_output_path(input_path):
    if not input_path:
        return ""

    folder, filename = os.path.split(input_path)
    name, ext = os.path.splitext(filename)
    if not ext:
        ext = ".xlsx"

    if name.endswith("_结果"):
        return input_path

    return os.path.join(folder, f"{name}_结果{ext}")


def get_default_excel_path():
    candidates = []
    for name in os.listdir(BASE_DIR):
        if not name.lower().endswith(".xlsx"):
            continue
        if name.startswith("tmp_"):
            continue
        candidates.append(name)

    if not candidates:
        return ""

    candidates.sort(key=lambda item: os.path.getmtime(os.path.join(BASE_DIR, item)), reverse=True)
    return os.path.join(BASE_DIR, candidates[0])


class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书提号工具")
        self.root.geometry("920x700")
        self.root.minsize(820, 620)

        self.process = None
        self.worker_thread = None
        self.log_queue = queue.Queue()
        self.auto_output = True

        self.input_path_var = tk.StringVar(value=get_default_excel_path())
        self.output_path_var = tk.StringVar(value=derive_output_path(self.input_path_var.get()))
        self.mode_var = tk.StringVar(value="brand")
        self.rows_var = tk.StringVar()
        self.verbose_var = tk.BooleanVar(value=False)
        self.login_wait_var = tk.StringVar(value="10")

        self._build_ui()
        self._bind_events()
        self.root.after(150, self._drain_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        top = ttk.Frame(self.root, padding=16)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(1, weight=1)

        title = ttk.Label(top, text="本地提号启动器", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        desc = ttk.Label(
            top,
            text="不需要后端服务。选本地 Excel、选 sheet、填行号，直接运行现有抓取脚本。",
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))

        ttk.Label(top, text="输入 Excel").grid(row=2, column=0, sticky="w", pady=6)
        self.input_entry = ttk.Entry(top, textvariable=self.input_path_var)
        self.input_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=6)
        ttk.Button(top, text="浏览", command=self._pick_input_file).grid(row=2, column=2, sticky="ew", pady=6)

        ttk.Label(top, text="输出 Excel").grid(row=3, column=0, sticky="w", pady=6)
        self.output_entry = ttk.Entry(top, textvariable=self.output_path_var)
        self.output_entry.grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=6)
        ttk.Button(top, text="浏览", command=self._pick_output_file).grid(row=3, column=2, sticky="ew", pady=6)

        mode_frame = ttk.LabelFrame(self.root, text="执行设置", padding=16)
        mode_frame.grid(row=1, column=0, sticky="ew", padx=16)
        mode_frame.columnconfigure(1, weight=1)

        ttk.Label(mode_frame, text="Sheet").grid(row=0, column=0, sticky="w")
        mode_box = ttk.Frame(mode_frame)
        mode_box.grid(row=0, column=1, sticky="w")
        for index, (label, value) in enumerate(MODE_OPTIONS):
            ttk.Radiobutton(mode_box, text=label, value=value, variable=self.mode_var).grid(
                row=0, column=index, padx=(0, 16), sticky="w"
            )

        ttk.Label(mode_frame, text="行号").grid(row=1, column=0, sticky="w", pady=(12, 0))
        self.rows_entry = ttk.Entry(mode_frame, textvariable=self.rows_var)
        self.rows_entry.grid(row=1, column=1, sticky="ew", pady=(12, 0))

        rows_tip = ttk.Label(mode_frame, text="支持 34 / 34-40 / 34,36,40-45")
        rows_tip.grid(row=2, column=1, sticky="w", pady=(4, 0))

        ttk.Label(mode_frame, text="登录等待").grid(row=3, column=0, sticky="w", pady=(12, 0))
        login_row = ttk.Frame(mode_frame)
        login_row.grid(row=3, column=1, sticky="w", pady=(12, 0))
        ttk.Entry(login_row, textvariable=self.login_wait_var, width=8).grid(row=0, column=0, sticky="w")
        ttk.Label(login_row, text="秒").grid(row=0, column=1, padx=(6, 12))
        ttk.Checkbutton(login_row, text="详细日志", variable=self.verbose_var).grid(row=0, column=2, sticky="w")

        action_bar = ttk.Frame(self.root, padding=(16, 12))
        action_bar.grid(row=2, column=0, sticky="ew")
        action_bar.columnconfigure(1, weight=1)

        self.run_button = ttk.Button(action_bar, text="开始运行", command=self._start_run)
        self.run_button.grid(row=0, column=0, sticky="w")

        ttk.Button(action_bar, text="打开所在文件夹", command=self._open_output_folder).grid(
            row=0, column=2, sticky="e"
        )

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=16)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 16))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

    def _bind_events(self):
        self.input_path_var.trace_add("write", self._handle_input_path_change)
        self.output_entry.bind("<KeyRelease>", self._mark_output_manual)
        self.output_entry.bind("<FocusIn>", self._mark_output_manual)

    def _handle_input_path_change(self, *_args):
        if self.auto_output:
            self.output_path_var.set(derive_output_path(self.input_path_var.get()))

    def _mark_output_manual(self, _event=None):
        self.auto_output = False

    def _pick_input_file(self):
        path = filedialog.askopenfilename(
            title="选择输入 Excel",
            initialdir=BASE_DIR,
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.auto_output = True
        self.input_path_var.set(path)

    def _pick_output_file(self):
        current = self.output_path_var.get().strip() or derive_output_path(self.input_path_var.get())
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel",
            initialdir=os.path.dirname(current) if current else BASE_DIR,
            initialfile=os.path.basename(current) if current else "",
            defaultextension=".xlsx",
            filetypes=XLSX_FILE_TYPES,
        )
        if not path:
            return
        self.auto_output = False
        self.output_path_var.set(path)

    def _append_log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _log(self, text):
        self.log_queue.put(("log", text))

    def _set_running(self, is_running):
        state = "disabled" if is_running else "normal"
        self.run_button.configure(state=state)
        if is_running:
            self.run_button.configure(text="运行中...")
        else:
            self.run_button.configure(text="开始运行")

    def _validate_form(self):
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        rows = self.rows_var.get().strip()
        login_wait = self.login_wait_var.get().strip()

        if not input_path:
            raise ValueError("请选择输入 Excel。")
        if not os.path.exists(input_path):
            raise ValueError("输入 Excel 不存在。")
        if not output_path:
            raise ValueError("请填写输出 Excel 路径。")
        if not rows:
            raise ValueError("请填写行号。")
        if not login_wait.isdigit():
            raise ValueError("登录等待时间必须是整数秒。")
        return input_path, output_path, rows, login_wait

    def _build_command(self, input_path, output_path, rows):
        command = [
            resolve_cli_python(),
            "-u",
            RUNNER_PATH,
            self.mode_var.get(),
            rows,
            "--excel",
            input_path,
            "--output",
            output_path,
        ]
        if self.verbose_var.get():
            command.append("--verbose")
        return command

    def _start_run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        try:
            input_path, output_path, rows, login_wait = self._validate_form()
        except ValueError as exc:
            messagebox.showerror("参数不完整", str(exc))
            return

        if os.path.abspath(input_path) == os.path.abspath(output_path):
            confirm = messagebox.askyesno(
                "确认覆盖",
                "输入文件和输出文件是同一个路径。\n这会直接覆盖当前文件，是否继续？",
            )
            if not confirm:
                return

        env = os.environ.copy()
        env["XHS_LOGIN_WAIT_SECONDS"] = login_wait
        env["XHS_REQUIRE_ENTER_CONFIRM"] = "0"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["PYTHONUNBUFFERED"] = "1"

        command = self._build_command(input_path, output_path, rows)

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self._set_running(True)
        self._append_log("命令：\n")
        self._append_log(" ".join(f'"{part}"' if " " in part else part for part in command) + "\n\n")

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
                self._log(line)

            return_code = self.process.wait()
            self.log_queue.put(("done", return_code))
        except Exception as exc:
            self.log_queue.put(("error", str(exc)))
        finally:
            self.process = None

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
                    messagebox.showinfo("完成", "脚本执行完成。")
                else:
                    self._append_log(f"\n运行失败，退出码：{payload}\n")
                    messagebox.showerror("失败", f"脚本执行失败，退出码：{payload}")
            elif event_type == "error":
                self._set_running(False)
                self._append_log(f"\n启动失败：{payload}\n")
                messagebox.showerror("启动失败", payload)

        self.root.after(150, self._drain_log_queue)

    def _open_output_folder(self):
        output_path = self.output_path_var.get().strip()
        target = os.path.dirname(output_path) if output_path else BASE_DIR
        if not os.path.isdir(target):
            target = BASE_DIR
        os.startfile(target)

    def _on_close(self):
        if self.process and self.process.poll() is None:
            confirm = messagebox.askyesno("确认退出", "脚本还在运行，确定关闭窗口吗？")
            if not confirm:
                return
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
