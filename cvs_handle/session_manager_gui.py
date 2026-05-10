import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BASE_DIR)
MAIN_GUI_PATH = os.path.join(BASE_DIR, "fix_xhs_note_urls_parallel_gui.py")


class SessionManagerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("XHS Session Manager")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        self.sessions = []
        self.log_queue = []
        self.process = None
        
        self._build_ui()
        self._load_existing_sessions()
        self.root.after(100, self._drain_log_queue)
    
    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        
        form = ttk.Frame(self.root, padding=16)
        form.grid(row=0, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)
        
        title = ttk.Label(form, text="Multi-Account Session Manager", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")
        
        hint = ttk.Label(form, text="Tip: Initialize multiple accounts to rotate and avoid rate limits", foreground="gray")
        hint.grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
        
        row1 = ttk.Frame(form)
        row1.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        
        ttk.Label(row1, text="Account Number:").grid(row=0, column=0, sticky="w")
        self.account_var = tk.StringVar(value="1")
        self.account_entry = ttk.Entry(row1, textvariable=self.account_var, width=10)
        self.account_entry.grid(row=0, column=1, padx=(8, 0))
        
        ttk.Label(row1, text="Wait Time (sec):").grid(row=0, column=2, padx=(16, 0))
        self.wait_var = tk.StringVar(value="120")
        ttk.Entry(row1, textvariable=self.wait_var, width=8).grid(row=0, column=3, padx=(8, 0))
        
        self.init_btn = ttk.Button(row1, text="Init Session", command=self._init_session)
        self.init_btn.grid(row=0, column=4, padx=(16, 0))
        
        self.stop_btn = ttk.Button(row1, text="Stop", command=self._stop_process, state="disabled")
        self.stop_btn.grid(row=0, column=5, padx=(8, 0))
        
        list_frame = ttk.LabelFrame(self.root, text="Existing Sessions", padding=8)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8, 0))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        columns = ("account", "path", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=6)
        self.tree.heading("account", text="Account")
        self.tree.heading("path", text="Session Path")
        self.tree.heading("status", text="Status")
        self.tree.column("account", width=80)
        self.tree.column("path", width=400)
        self.tree.column("status", width=100)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        
        ttk.Button(btn_frame, text="Refresh", command=self._load_existing_sessions).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(btn_frame, text="Verify Account", command=self._open_folder).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(btn_frame, text="Delete Selected", command=self._delete_selected).grid(row=0, column=2)
        
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=8)
        log_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(8, 16))
        log_frame.columnconfigure(0, weight=1)
        
        self.log_text = ScrolledText(log_frame, height=8, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="ew")
        
        output_frame = ttk.Frame(self.root, padding=(16, 0, 16, 16))
        output_frame.grid(row=3, column=0, sticky="ew")
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Session String for GUI:").grid(row=0, column=0, sticky="w")
        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var, state="readonly").grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Button(output_frame, text="Copy", command=self._copy_to_clipboard).grid(row=0, column=2, padx=(8, 0))
        
        run_frame = ttk.Frame(self.root, padding=(16, 0, 16, 16))
        run_frame.grid(row=4, column=0, sticky="ew")
        
        run_btn = ttk.Button(run_frame, text="Run Main Program (Parallel Processing)", command=self._launch_main_gui)
        run_btn.grid(row=0, column=0)
        
        ttk.Label(run_frame, text="<- Launch the main program after initializing sessions", foreground="gray").grid(row=0, column=1, padx=(8, 0))
    
    def _log(self, msg):
        self.log_queue.append(msg)
    
    def _drain_log_queue(self):
        while self.log_queue:
            msg = self.log_queue.pop(0)
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(100, self._drain_log_queue)
    
    def _load_existing_sessions(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.sessions = []
        
        for name in os.listdir(REPO_DIR):
            if name.startswith("browser_session"):
                full_path = os.path.join(REPO_DIR, name)
                if os.path.isdir(full_path):
                    parts = name.split("_")
                    if len(parts) >= 3 and parts[-1].isdigit():
                        account_num = parts[-1]
                    else:
                        account_num = "?"
                    
                    has_data = self._check_session_data(full_path)
                    status = "Ready" if has_data else "Empty"
                    
                    self.tree.insert("", "end", values=(account_num, full_path, status))
                    if has_data:
                        self.sessions.append(name)
        
        self._update_output_string()
    
    def _check_session_data(self, path):
        for item in os.listdir(path):
            if item.startswith("parallel_"):
                continue
            item_path = os.path.join(path, item)
            if os.path.isdir(item_path):
                return True
        return False
    
    def _update_output_string(self):
        if self.sessions:
            self.output_var.set(",".join(self.sessions))
        else:
            self.output_var.set("")
    
    def _init_session(self):
        try:
            account = int(self.account_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid account number")
            return
        
        wait_time = 120
        try:
            wait_time = int(self.wait_var.get().strip())
        except ValueError:
            pass
        
        session_dir = os.path.join(REPO_DIR, f"browser_session_{account}")
        
        self.init_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._log(f"Initializing Account {account}...")
        self._log(f"Session dir: {session_dir}")
        
        def run_init():
            try:
                from playwright.sync_api import sync_playwright
                
                if not os.path.exists(session_dir):
                    os.makedirs(session_dir, exist_ok=True)
                
                self._log("Opening browser, please login in the browser window...")
                
                with sync_playwright() as p:
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=session_dir,
                        headless=False,
                        slow_mo=80,
                        viewport={"width": 1440, "height": 1000},
                    )
                    page = context.pages[0] if context.pages else context.new_page()
                    
                    page.goto("https://www.xiaohongshu.com/explore", wait_until="domcontentloaded", timeout=45000)
                    
                    import time
                    deadline = time.time() + wait_time
                    logged_in = False
                    
                    while time.time() < deadline:
                        page.wait_for_timeout(2000)
                        try:
                            is_login_required = page.evaluate(
                                """
                                () => {
                                    const body = document.body ? (document.body.innerText || '') : '';
                                    const modal = document.querySelector('.login-modal, [class*="login-modal"], .login-container');
                                    return Boolean(
                                        modal
                                        || body.includes('登录即可查看')
                                        || (body.includes('手机号登录') && body.includes('获取验证码'))
                                    );
                                }
                                """
                            )
                            if not is_login_required:
                                logged_in = True
                                break
                        except Exception:
                            pass
                        
                        remaining = int(deadline - time.time())
                        if remaining > 0 and remaining % 10 == 0:
                            self._log(f"Waiting for login... {remaining} sec remaining")
                    
                    context.close()
                    
                    if logged_in:
                        self._log(f"Account {account} login success!")
                        self.root.after(0, self._load_existing_sessions)
                    else:
                        self._log(f"Account {account} login timeout")
                
            except Exception as e:
                self._log(f"Error: {e}")
            finally:
                self.root.after(0, self._init_done)
        
        threading.Thread(target=run_init, daemon=True).start()
    
    def _init_done(self):
        self.init_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
    
    def _stop_process(self):
        self._log("Stop requested (close the browser window to stop)")
    
    def _open_folder(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a session to verify")
            return
        
        item = self.tree.item(selected[0])
        path = item["values"][1]
        account = item["values"][0]
        
        if not os.path.exists(path):
            messagebox.showerror("Error", f"Session path not found: {path}")
            return
        
        self._log(f"Opening browser for Account {account}...")
        
        def open_browser():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=path,
                        headless=False,
                        viewport={"width": 1200, "height": 800},
                    )
                    page = context.pages[0] if context.pages else context.new_page()
                    page.goto("https://www.xiaohongshu.com/user/profile/self", wait_until="domcontentloaded", timeout=30000)
                    self._log(f"Browser opened for Account {account}. Close the browser when done.")
                    page.wait_for_timeout(300000)
                    context.close()
            except Exception as e:
                self._log(f"Error: {e}")
        
        threading.Thread(target=open_browser, daemon=True).start()
    
    def _delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a session to delete")
            return
        
        item = self.tree.item(selected[0])
        path = item["values"][1]
        account = item["values"][0]
        
        if messagebox.askyesno("Confirm", f"Delete session for Account {account}?\n{path}"):
            import shutil
            try:
                shutil.rmtree(path)
                self._log(f"Deleted: {path}")
                self._load_existing_sessions()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete: {e}")
    
    def _copy_to_clipboard(self):
        text = self.output_var.get()
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("Copied", "Session string copied to clipboard")
    
    def _launch_main_gui(self):
        if not os.path.exists(MAIN_GUI_PATH):
            messagebox.showerror("Error", f"Main GUI not found: {MAIN_GUI_PATH}")
            return
        
        python_exe = sys.executable
        if python_exe.endswith("pythonw.exe"):
            python_folder = os.path.dirname(python_exe)
            python_exe = os.path.join(python_folder, "python.exe")
        
        try:
            subprocess.Popen([python_exe, MAIN_GUI_PATH])
            self._log("Main program launched")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch main program: {e}")


def main():
    root = tk.Tk()
    app = SessionManagerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
