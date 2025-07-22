import sys, subprocess
from importlib import metadata               # Python 3.8+ 內建
from packaging import version                # pip 的相依套件，通常已隨 pip 一起安裝

# 1. 在這裡列出「至少需要的版本」
REQUIRED = {
    "pandas":      "2.2.4",
    "plotly":      "6.0.1",
    "yfinance":    "0.2.61",
    "ttkbootstrap":"1.13.9",
    "tkcalendar":  "1.6.1",
    "XlsxWriter":  "3.2.5",
    "gspread":     "6.2.1",
    "oauth2client":"4.1.3"
}

def ensure_one(pkg: str, min_ver: str):
    """
    若模組缺失或版本過舊，呼叫 pip 安裝／升級到 min_ver 以上。
    """
    try:
        cur_ver = metadata.version(pkg)
        if version.parse(cur_ver) >= version.parse(min_ver):
            print(f"{pkg} 目前版本 {cur_ver}，符合需求 {min_ver}")
            return                          # OK
        print(f"{pkg} 目前版本 {cur_ver}，低於需求 {min_ver}，升級中…")
    except metadata.PackageNotFoundError:
        print(f"{pkg} 未安裝，安裝中…")

    subprocess.check_call([sys.executable, "-m", "pip",
                           "install", f"{pkg}>={min_ver}", "--break-system-packages"])

def ensure_requirements():
    for pkg, ver in REQUIRED.items():
        ensure_one(pkg, ver)

# 2. 供主程式呼叫
if __name__ == "__main__":
    ensure_requirements()
