import argparse
import os
import sys
import time

from playwright.sync_api import sync_playwright


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(BASE_DIR)
DEFAULT_SESSION_DIR = os.path.join(REPO_DIR, "browser_session")


def login_account(session_dir, wait_seconds=120):
    if not os.path.exists(session_dir):
        os.makedirs(session_dir, exist_ok=True)
    
    print(f"Session 目录: {session_dir}", flush=True)
    print(f"正在打开浏览器，请在浏览器中登录小红书账号...", flush=True)
    print(f"最多等待 {wait_seconds} 秒", flush=True)
    
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=session_dir,
            headless=False,
            slow_mo=80,
            viewport={"width": 1440, "height": 1000},
        )
        page = context.pages[0] if context.pages else context.new_page()
        
        page.goto("https://www.xiaohongshu.com/explore", wait_until="domcontentloaded", timeout=45000)
        
        deadline = time.time() + wait_seconds
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
                print(f"仍在等待登录... 剩余 {remaining} 秒", flush=True)
        
        if logged_in:
            print("\n登录成功！Session 已保存。", flush=True)
        else:
            print("\n等待超时，请重新运行脚本。", flush=True)
        
        context.close()
        return logged_in


def main():
    parser = argparse.ArgumentParser(description="初始化小红书账号 Session")
    parser.add_argument("--session-dir", default="", help="Session 目录路径，默认 browser_session")
    parser.add_argument("--wait", type=int, default=120, help="等待登录秒数，默认 120")
    parser.add_argument("--account", type=int, default=1, help="账号编号，用于自动生成 session 目录名")
    args = parser.parse_args()
    
    if args.session_dir:
        session_dir = args.session_dir
    else:
        session_dir = os.path.join(REPO_DIR, f"browser_session_{args.account}")
    
    print("=" * 60, flush=True)
    print(f"初始化账号 {args.account} 的 Session", flush=True)
    print("=" * 60, flush=True)
    
    success = login_account(session_dir, args.wait)
    
    if success:
        print(f"\nSession 目录: {session_dir}", flush=True)
        print("可以在 GUI 的'多账号Session'中填写该目录路径", flush=True)
        if args.account > 1:
            print(f"填写格式: browser_session_1,browser_session_2,...", flush=True)
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
