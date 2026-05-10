import os
import sys
import time

from playwright.sync_api import sync_playwright


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BASE_DIR)


def verify_session(session_dir, account_num):
    print(f"\n{'='*60}", flush=True)
    print(f"正在验证账号 {account_num}...", flush=True)
    print(f"Session 目录: {session_dir}", flush=True)
    print(f"{'='*60}", flush=True)
    
    if not os.path.exists(session_dir):
        print(f"错误: Session 目录不存在", flush=True)
        return
    
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=session_dir,
            headless=False,
            viewport={"width": 1200, "height": 800},
        )
        page = context.pages[0] if context.pages else context.new_page()
        
        page.goto("https://www.xiaohongshu.com/user/profile/self", wait_until="domcontentloaded", timeout=30000)
        
        print("\n浏览器已打开，请查看当前登录的账号", flush=True)
        print("确认后请关闭浏览器窗口，继续验证下一个账号...", flush=True)
        
        input("\n按 Enter 关闭浏览器并继续...")
        
        context.close()


def main():
    sessions = []
    
    for name in os.listdir(REPO_DIR):
        if name.startswith("browser_session"):
            full_path = os.path.join(REPO_DIR, name)
            if os.path.isdir(full_path):
                sessions.append((name, full_path))
    
    sessions.sort(key=lambda x: x[0])
    
    print(f"发现 {len(sessions)} 个 session:", flush=True)
    for i, (name, path) in enumerate(sessions):
        print(f"  {i+1}. {name}", flush=True)
    
    print("\n将依次打开每个 session 的浏览器窗口，请确认账号", flush=True)
    input("按 Enter 开始验证...")
    
    for i, (name, path) in enumerate(sessions):
        verify_session(path, i + 1)
    
    print("\n" + "="*60, flush=True)
    print("验证完成！", flush=True)
    print("="*60, flush=True)


if __name__ == "__main__":
    main()
