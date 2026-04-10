#!/usr/bin/env python3
"""
飞书机器流转规划同步工具 - GUI启动器
"""

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(BASE_DIR, "sync_machine_plan.py")
SPREADSHEET_TOKEN = "VjIrs07udhzM85tkinZcX6wBnte"


def resolve_cli_python():
    current = sys.executable
    folder = os.path.dirname(current)
    name = os.path.basename(current).lower()
    if name == "pythonw.exe":
        candidate = os.path.join(folder, "python.exe")
        if os.path.exists(candidate):
            return candidate
    return current


class LarkSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("飞书机器流转规划同步工具")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)

        self.process = None
        self.worker_thread = None
        self.log_queue = queue.Queue()

        self.xhs_3_row_var = tk.StringVar(value="3")
        self.xhs_4_row_var = tk.StringVar(value="3")
        self.douyin_3_row_var = tk.StringVar(value="2")
        self.douyin_4_row_var = tk.StringVar(value="3")
        self.update_published_var = tk.BooleanVar(value=True)
        self.update_koc_status_var = tk.BooleanVar(value=True)
        self.update_modify_flag_var = tk.BooleanVar(value=True)

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

        title = ttk.Label(top, text="飞书机器流转规划同步", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        desc = ttk.Label(
            top,
            text="从4个确认执行sheet读取数据，筛选KOL后同步到机器流转规划表",
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))

        input_frame = ttk.LabelFrame(self.root, text="起始行设置（每次运行前填写）", padding=16)
        input_frame.grid(row=1, column=0, sticky="ew", padx=16)
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

        ttk.Label(input_frame, text="小红书确认执行3月", font=("Microsoft YaHei UI", 10)).grid(
            row=0, column=0, sticky="w", pady=8
        )
        ttk.Entry(input_frame, textvariable=self.xhs_3_row_var, width=10).grid(
            row=0, column=1, sticky="w", padx=(8, 24)
        )

        ttk.Label(input_frame, text="小红书确认执行4月", font=("Microsoft YaHei UI", 10)).grid(
            row=0, column=2, sticky="w", pady=8
        )
        ttk.Entry(input_frame, textvariable=self.xhs_4_row_var, width=10).grid(
            row=0, column=3, sticky="w", padx=(8, 0)
        )

        ttk.Label(input_frame, text="3月抖音确认执行", font=("Microsoft YaHei UI", 10)).grid(
            row=1, column=0, sticky="w", pady=8
        )
        ttk.Entry(input_frame, textvariable=self.douyin_3_row_var, width=10).grid(
            row=1, column=1, sticky="w", padx=(8, 24)
        )

        ttk.Label(input_frame, text="4月抖音确认执行", font=("Microsoft YaHei UI", 10)).grid(
            row=1, column=2, sticky="w", pady=8
        )
        ttk.Entry(input_frame, textvariable=self.douyin_4_row_var, width=10).grid(
            row=1, column=3, sticky="w", padx=(8, 0)
        )

        tip_label = ttk.Label(
            input_frame,
            text="提示：起始行是数据开始的行号，不是表头行。从表头下一行开始填。",
            foreground="gray",
        )
        tip_label.grid(row=2, column=0, columnspan=4, sticky="w", pady=(8, 0))

        self.update_published_check = ttk.Checkbutton(
            input_frame,
            text='更新已发布状态（同步时将"已发布"的博主标记到"是否已发布"列）',
            variable=self.update_published_var
        )
        self.update_published_check.grid(row=3, column=0, columnspan=4, sticky="w", pady=(12, 0))

        self.update_koc_status_check = ttk.Checkbutton(
            input_frame,
            text='KOC或KOL地址为抠图时，更新气源为"无需收集"、快递状态为"无需配送产品"、样机情况为"无样机"',
            variable=self.update_koc_status_var
        )
        self.update_koc_status_check.grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))

        self.update_modify_flag_check = ttk.Checkbutton(
            input_frame,
            text='更新确认执行表的"是否修改"列（KOL且地址非抠图的数据标记为"是"）',
            variable=self.update_modify_flag_var
        )
        self.update_modify_flag_check.grid(row=5, column=0, columnspan=4, sticky="w", pady=(8, 0))

        action_bar = ttk.Frame(self.root, padding=(16, 12))
        action_bar.grid(row=2, column=0, sticky="ew")
        action_bar.columnconfigure(1, weight=1)

        self.run_button = ttk.Button(action_bar, text="开始同步", command=self._start_run)
        self.run_button.grid(row=0, column=0, sticky="w")

        ttk.Button(action_bar, text="打开所在文件夹", command=self._open_folder).grid(
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
        pass

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

    def _validate_form(self):
        try:
            xhs_3 = int(self.xhs_3_row_var.get().strip())
            xhs_4 = int(self.xhs_4_row_var.get().strip())
            dy_3 = int(self.douyin_3_row_var.get().strip())
            dy_4 = int(self.douyin_4_row_var.get().strip())
        except ValueError:
            raise ValueError("起始行必须是数字")

        if xhs_3 < 1 or xhs_4 < 1 or dy_3 < 1 or dy_4 < 1:
            raise ValueError("起始行必须大于0")

        return xhs_3, xhs_4, dy_3, dy_4

    def _start_run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return

        try:
            xhs_3, xhs_4, dy_3, dy_4 = self._validate_form()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        update_published = "1" if self.update_published_var.get() else "0"
        update_koc_status = "1" if self.update_koc_status_var.get() else "0"
        update_modify_flag = "1" if self.update_modify_flag_var.get() else "0"

        python_exec = resolve_cli_python()
        command = [
            python_exec,
            "-u",
            SCRIPT_PATH,
            str(xhs_3),
            str(xhs_4),
            str(dy_3),
            str(dy_4),
            update_published,
            update_koc_status,
            update_modify_flag
        ]

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
    LarkSyncApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
