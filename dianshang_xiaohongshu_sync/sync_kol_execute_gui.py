#!/usr/bin/env python3
"""
小红书KOL执行表同步工具 - GUI启动器
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, "sync_kol_execute.py")


def resolve_cli_python():
    current = sys.executable
    folder = os.path.dirname(current)
    name = os.path.basename(current).lower()
    if name == "pythonw.exe":
        candidate = os.path.join(folder, "python.exe")
        if os.path.exists(candidate):
            return candidate
    return current


class SyncKolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("小红书KOL执行表同步工具")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        self.process = None
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.token_var = tk.StringVar(value="")
        self.token_mar_var = tk.StringVar(value="")
        self.kol_var = tk.BooleanVar(value=True)
        self.koc_var = tk.BooleanVar(value=True)
        self.flow_var = tk.BooleanVar(value=True)

        self._build_ui()
        self.root.after(150, self._drain_log_queue)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(5, weight=1)

        top = ttk.Frame(self.root, padding=16)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(1, weight=1)

        title = ttk.Label(top, text="小红书KOL执行表同步", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        desc = ttk.Label(
            top,
            text="从二核-KOL/KOC同步到确认执行表，再按月份分流到机器流转统计",
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))

        token_frame = ttk.LabelFrame(self.root, text="文档Token设置（可选）", padding=16)
        token_frame.grid(row=1, column=0, sticky="ew", padx=16)
        token_frame.columnconfigure(1, weight=1)

        ttk.Label(token_frame, text="SPREADSHEET_TOKEN:", font=("Microsoft YaHei UI", 10)).grid(
            row=0, column=0, sticky="w", pady=4
        )
        ttk.Entry(token_frame, textvariable=self.token_var, width=50).grid(
            row=0, column=1, sticky="ew", padx=(8, 0)
        )

        ttk.Label(token_frame, text="TOKEN_MAR（3月文档）:", font=("Microsoft YaHei UI", 10)).grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Entry(token_frame, textvariable=self.token_mar_var, width=50).grid(
            row=1, column=1, sticky="ew", padx=(8, 0)
        )

        ttk.Label(
            token_frame,
            text="提示：留空则使用脚本中的默认值",
            foreground="gray",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        task_frame = ttk.LabelFrame(self.root, text="同步任务选择", padding=16)
        task_frame.grid(row=2, column=0, sticky="ew", padx=16)

        ttk.Checkbutton(
            task_frame,
            text="KOL任务: 二核-KOL → 确认执行-KOL-4月/5月/6月",
            variable=self.kol_var,
        ).grid(row=0, column=0, sticky="w", pady=4)

        ttk.Checkbutton(
            task_frame,
            text="KOC任务: 二核-KOC → 确认执行-KOC-4月/5月/6月",
            variable=self.koc_var,
        ).grid(row=1, column=0, sticky="w", pady=4)

        ttk.Checkbutton(
            task_frame,
            text="流转任务: 确认执行表 → 机器流转统计（按月份自动分流）",
            variable=self.flow_var,
        ).grid(row=2, column=0, sticky="w", pady=4)

        mapping_frame = ttk.LabelFrame(self.root, text="说明", padding=12)
        mapping_frame.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 0))

        mapping_text = (
            "流转任务会根据确认执行表的月份自动分流到对应的目标表：\n"
            "  3月 → 机器流转统计-3月特殊, 4月 → 机器流转统计3-4月, 5月 → 机器流转统计5月\n"
            "地址列支持富文本格式自动提取纯文本内容"
        )
        ttk.Label(mapping_frame, text=mapping_text, wraplength=700, foreground="gray", font=("Microsoft YaHei UI", 9)).grid(
            row=0, column=0, sticky="w"
        )

        action_bar = ttk.Frame(self.root, padding=(16, 12))
        action_bar.grid(row=4, column=0, sticky="ew")
        action_bar.columnconfigure(1, weight=1)

        self.run_button = ttk.Button(action_bar, text="开始同步", command=self._start_run)
        self.run_button.grid(row=0, column=0, sticky="w")

        ttk.Button(action_bar, text="打开所在文件夹", command=self._open_folder).grid(
            row=0, column=2, sticky="e"
        )

        log_frame = ttk.LabelFrame(self.root, text="运行日志", padding=16)
        log_frame.grid(row=5, column=0, sticky="nsew", padx=16, pady=(0, 16))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, wrap="word", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.configure(state="disabled")

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
            self.run_button.configure(text="同步中...")
        else:
            self.run_button.configure(text="开始同步")

    def _start_run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        tasks = ""
        if self.kol_var.get():
            tasks += "1"
        if self.koc_var.get():
            tasks += "2"
        if self.flow_var.get():
            tasks += "3"

        if not tasks:
            messagebox.showwarning("提示", "请至少选择一个同步任务")
            return

        python_exec = resolve_cli_python()
        command = [python_exec, "-u", SCRIPT_PATH]
        
        token_value = self.token_var.get().strip()
        if token_value:
            command.extend(["--token", token_value])
        
        token_mar_value = self.token_mar_var.get().strip()
        if token_mar_value:
            command.extend(["--token-mar", token_mar_value])
        
        command.append(tasks)

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self._set_running(True)
        self._append_log("命令：\n")
        self._append_log(" ".join(f'"{part}"' if " " in part else part for part in command) + "\n\n")

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
                    self._append_log("\n同步完成。\n")
                    messagebox.showinfo("完成", "数据同步完成！")
                else:
                    self._append_log(f"\n同步失败，退出码：{payload}\n")
                    messagebox.showerror("失败", f"同步失败，退出码：{payload}")
            elif event_type == "error":
                self._set_running(False)
                self._append_log(f"\n启动失败：{payload}\n")
                messagebox.showerror("启动失败", payload)

        self.root.after(150, self._drain_log_queue)

    def _open_folder(self):
        os.startfile(BASE_DIR)

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
    SyncKolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
