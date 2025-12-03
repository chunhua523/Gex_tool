import os
import sys
import subprocess
import hashlib
import time
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import unicodedata
import encodings
import traceback

# 設定日誌檔案
LOG_FILE = "launcher_log.txt"

def log(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")
    except:
        pass

# 嘗試匯入 requests，如果失敗則提示
try:
    import requests
except ImportError:
    log("ImportError: requests module not found.")
    messagebox.showerror("錯誤", "缺少 requests 模組，請確保已安裝。\npip install requests")
    sys.exit(1)

# 設定 GitHub 資訊
GITHUB_USER = "chunhua523"
GITHUB_REPO = "Gex_tool"
BRANCH = "master"
# 主程式檔案名稱
MAIN_SCRIPT = "GEX_chart_new.py"
# 依賴檔案 (如果有其他檔案需要同步，加在這裡)
FILES_TO_SYNC = [
    "GEX_chart_new.py",
    "auto_requirements.py",
    "service_account.json" # 注意：通常憑證不建議放公開 Repo，若為私有 Repo 需改用 Token 驗證
]

BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{BRANCH}/"

def get_local_hash(filename):
    """計算本地檔案的 SHA256"""
    if not os.path.exists(filename):
        return None
    sha256_hash = hashlib.sha256()
    with open(filename, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def get_remote_content(filename):
    """獲取遠端檔案內容"""
    url = BASE_URL + filename
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.content
        else:
            print(f"無法下載 {filename}: Status {response.status_code}")
            return None
    except Exception as e:
        print(f"連線錯誤: {e}")
        return None

def update_files(status_label, progress_bar, root):
    """檢查並更新檔案"""
    updated = False
    total_files = len(FILES_TO_SYNC)
    
    for i, filename in enumerate(FILES_TO_SYNC):
        # 更新 UI
        status_label.config(text=f"正在檢查: {filename}...")
        progress_bar['value'] = (i / total_files) * 100
        root.update_idletasks()

        # 略過 service_account.json 的自動更新，避免覆蓋使用者的憑證
        # 如果您希望強制更新憑證，請移除此判斷
        if filename == "service_account.json" and os.path.exists(filename):
            continue

        remote_content = get_remote_content(filename)
        if remote_content:
            # 計算遠端內容的 Hash
            remote_hash = hashlib.sha256(remote_content).hexdigest()
            local_hash = get_local_hash(filename)

            if local_hash != remote_hash:
                status_label.config(text=f"發現更新: {filename}，下載中...")
                try:
                    with open(filename, "wb") as f:
                        f.write(remote_content)
                    updated = True
                    print(f"已更新: {filename}")
                except Exception as e:
                    messagebox.showerror("錯誤", f"寫入檔案失敗: {e}")
        else:
            print(f"跳過 {filename} (無法獲取)")

    progress_bar['value'] = 100
    status_label.config(text="檢查完成，準備啟動...")
    root.after(1000, lambda: launch_app(root))

def launch_app(root):
    """執行主程式"""
    root.destroy()
    log("準備啟動主程式...")
    
    # 判斷執行環境
    if getattr(sys, 'frozen', False):
        python_exe = "python"
    else:
        python_exe = sys.executable
    
    log(f"使用 Python 直譯器: {python_exe}")

    # 準備環境變數，移除 PyInstaller 注入的 Tcl/Tk 變數以免干擾子程序
    # PyInstaller 打包後的執行檔會設定 TCL_LIBRARY 和 TK_LIBRARY 指向臨時目錄
    # 這會導致子程序 (使用系統 Python) 找不到正確的 Tcl/Tk 庫
    env = os.environ.copy()
    if getattr(sys, 'frozen', False):
        env.pop('TCL_LIBRARY', None)
        env.pop('TK_LIBRARY', None)
        log("已清除子程序的 TCL_LIBRARY/TK_LIBRARY 環境變數")

    # 使用 cmd /k 讓視窗在執行後保持開啟，以便查看錯誤訊息
    if os.name == 'nt':
        cmd = ["cmd", "/k", python_exe, MAIN_SCRIPT]
    else:
        cmd = [python_exe, MAIN_SCRIPT]

    log(f"執行指令: {cmd}")

    try:
        subprocess.Popen(cmd, env=env)
        log("主程式已啟動")
    except Exception as e:
        err_msg = f"無法啟動主程式:\n{e}\n請確認電腦已安裝 Python 並加入環境變數(PATH)。"
        log(err_msg)
        # 由於 root 已銷毀，這裡創建一個新的臨時 root 來顯示錯誤
        err_root = tk.Tk()
        err_root.withdraw()
        messagebox.showerror("啟動失敗", err_msg)
        err_root.destroy()

def main():
    try:
        log("Launcher 啟動")
        root = tk.Tk()
        root.title("GEX Tool Launcher")
        root.geometry("400x150")
        
        # 置中視窗
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 150) // 2
        root.geometry(f"400x150+{x}+{y}")

        tk.Label(root, text="GEX Tool 自動更新啟動器", font=("Arial", 14, "bold")).pack(pady=10)
        
        status_label = tk.Label(root, text="準備連線...", font=("Arial", 10))
        status_label.pack(pady=5)
        
        progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        progress_bar.pack(pady=10)

        # 使用 Thread 避免卡住 UI
        threading.Thread(target=update_files, args=(status_label, progress_bar, root), daemon=True).start()

        root.mainloop()
    except Exception as e:
        log(f"Launcher 發生未預期錯誤: {traceback.format_exc()}")
        # 嘗試顯示錯誤訊息
        try:
            messagebox.showerror("Launcher Error", f"發生嚴重錯誤:\n{e}")
        except:
            pass

if __name__ == "__main__":
    main()
