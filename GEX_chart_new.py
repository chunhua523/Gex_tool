import datetime
import typing
import auto_requirements
auto_requirements.ensure_requirements()

import os
import sqlite3
import tempfile
import tkinter as tk
from tkinter import messagebox, filedialog
import traceback
import pandas as pd
import plotly.graph_objects as go
import yfinance as yf
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry
from gspread.exceptions import APIError
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 全域控件
root = None
calendar_date = None
gex_entry = None
ticker_filter = None
start_date_filter = None
end_date_filter = None
tree = None
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 此 .py 檔所在資料夾
DB_PATH = os.path.join(BASE_DIR, "stocks.db")

# Google 試算表相關常數（請依需求自行修改）
SHEET_ID = '1u1opYwj_2bhOBhAM96CB7kYz9prWKQtXhmjU1cG15Dg'  # 試算表 ID
SHEET_ID_MINOR = '1H7MqEuVuu_xIN9B-rFMrevDLeaCn06z3dn0XBMt_-to'  # 試算表 ID
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'service_account.json')  # Service Account 憑證

# 全域狀態追蹤使用者選擇
user_conflict_choice = None  # "skip" 或 "overwrite" 或 "cancel"
apply_to_all = False
cancel_import = False
inserted_count = 0
all_tickers = []

# --- 資料庫相關 ---
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS stock_data (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        ticker TEXT NOT NULL,
                        date TEXT NOT NULL,
                        label TEXT NOT NULL,
                        value REAL NOT NULL)''')
    conn.commit()
    conn.close()

def get_latest_date_for_ticker(ticker: str):
    """回傳資料庫中指定 ticker 的最新日期 (datetime.date)；若無資料回傳 None"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT MAX(date) FROM stock_data WHERE ticker=?", (ticker,))
    row = cur.fetchone()
    conn.close()
    if row and row[0]:
        return pd.to_datetime(row[0]).date()
    return None


def insert_data(ticker: str, date_str: str, label: str, value: float):
    """將一筆資料寫入資料庫，如已存在就覆蓋"""
    global inserted_count
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""INSERT OR REPLACE INTO stock_data (ticker, date, label, value)
                   VALUES (?,?,?,?)""",
                (ticker, date_str, label, value))
    conn.commit()
    conn.close()
    inserted_count += 1

# 自訂彈出視窗讓使用者選擇如何處理重複資料
def ask_conflict_resolution(ticker, date, label):
    result = {"choice": None, "apply_all": False}

    def on_choice(choice):
        result["choice"] = choice
        result["apply_all"] = apply_var.get()
        dialog.destroy()

    # 使用 ttkbootstrap 的 Toplevel 來承接主視窗的主題
    dialog = tk.Toplevel(root)
    dialog.title("資料衝突")
    dialog.geometry("600x180")
    dialog.grab_set()

    # 標籤也改用 ttk.Label，確保跟隨 darkly 主題
    ttk.Label(dialog, text=f"{ticker} - {date} - {label} 已存在。\n是否要覆蓋？",
              padding=(20, 10)).pack()

    apply_var = tk.BooleanVar()
    # Checkbutton 改用 ttk.Checkbutton，並加上 bootstyle
    ttk.Checkbutton(dialog, text="套用至所有後續重複資料",
                    variable=apply_var,
                    bootstyle="secondary").pack()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)

    # 三個按鈕都改用 ttk.Button，並指定不同的 bootstyle
    ttk.Button(button_frame, text="覆蓋", width=10,
               bootstyle="success", command=lambda: on_choice("overwrite")
               ).pack(side="left", padx=5)
    ttk.Button(button_frame, text="跳過", width=10,
               bootstyle="warning", command=lambda: on_choice("skip")
               ).pack(side="left", padx=5)
    ttk.Button(button_frame, text="取消", width=10,
               bootstyle="danger", command=lambda: on_choice("cancel")
               ).pack(side="left", padx=5)

    dialog.wait_window()
    return result

# 插入數據
def insert_data(ticker, date, label, value):
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    if cancel_import:
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM stock_data WHERE ticker=? AND date=? AND label=?", (ticker, date, label))
    existing = cursor.fetchone()

    if existing:
        if not apply_to_all:
            res = ask_conflict_resolution(ticker, date, label)
            user_conflict_choice = res["choice"]
            apply_to_all = res["apply_all"]

        if user_conflict_choice == "cancel":
            cancel_import = True
            conn.close()
            return
        elif user_conflict_choice == "overwrite":
            cursor.execute("UPDATE stock_data SET value=? WHERE id=?", (value, existing[0]))
            inserted_count += 1
        elif user_conflict_choice == "skip":
            conn.close()
            return
    else:
        cursor.execute("INSERT INTO stock_data (ticker, date, label, value) VALUES (?, ?, ?, ?)", (ticker, date, label, value))
        inserted_count += 1

    conn.commit()
    conn.close()

# --- 功能函式 ---
def parse_gex_code(orig_date: str, gex_code: str) -> typing.Optional[str]:
    """
    解析 GEX TV Code 並寫入資料庫
    - 若 TV Code 自帶日期，優先使用
    - 解析前就先把 TV Code 原文存進資料庫（label = 'TV Code'）
    """
    global cancel_import
    if cancel_import:
        return None

    # 內嵌日期 > Date 欄
    embedded = _extract_date_from_tv_code(gex_code)
    date_str = embedded.isoformat() if embedded else orig_date.split(" ")[0]

    # 取 ticker（第一個 XXX:）
    m = re.search(r"([A-Za-z\.]+):", gex_code)
    if not m:
        messagebox.showwarning("格式錯誤", f"無法解析 GEX TV Code：{gex_code}")
        return None
    ticker = m.group(1).upper()

    # 👉 **先把原文存進去，之後任何匯入方式都不用再管**
    insert_data(ticker, date_str, "TV Code", gex_code.strip())

    code_body = gex_code[m.end():].strip()
    elements = re.split(r',\s*', code_body)
    i = 0
    while i < len(elements) - 1:
        labels = elements[i].strip()
        try:
            value = float(elements[i + 1].strip())
            for label in labels.split('&'):
                insert_data(ticker, date_str, label.strip(), value)
                if cancel_import:
                    return None
            i += 2
        except ValueError:
            i += 1
    return ticker

def single_entry():
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    user_conflict_choice = None
    apply_to_all = False
    cancel_import = False
    inserted_count = 0

    date = calendar_date.entry.get()
    gex_code = gex_entry.get()
    if not date or not gex_code:
        messagebox.showwarning("輸入錯誤", "請輸入日期和 GEX TV Code")
        return
    ticker = parse_gex_code(date, gex_code)
    if not cancel_import:
        populate_ticker_dropdown()
        if ticker:
            ticker_filter.set(ticker)
        refresh_table()
        messagebox.showinfo("完成", f"成功寫入 {inserted_count} 筆資料。")


def bulk_import():
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    user_conflict_choice = None
    apply_to_all = False
    cancel_import = False
    inserted_count = 0

    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv")])
    if not file_path:
        return

    # 嘗試從檔名解析日期 (例如: 20251212_TV Code.txt)
    filename = os.path.basename(file_path)
    default_date = None
    date_match = re.search(r"(\d{8})", filename)
    if date_match:
        try:
            default_date = pd.to_datetime(date_match.group(1), format='%Y%m%d').date().isoformat()
        except:
            pass
    
    current_date = default_date

    with open(file_path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]

    for line in lines:
        if cancel_import:
            messagebox.showinfo("已取消", f"成功寫入 {inserted_count} 筆資料。")
            return
        
        # 判斷是否為 GEX Code 行 (包含 ":")
        if ":" in line:
            # 若無 current_date，使用當日作為備案 (parse_gex_code 會優先嘗試內嵌日期)
            use_date = current_date if current_date else datetime.date.today().isoformat()
            parse_gex_code(use_date, line)
        else:
            # 嘗試解析為日期行 (舊格式相容)
            try:
                potential_date = line.split("_")[0]
                pd.to_datetime(potential_date)
                current_date = potential_date
            except:
                pass

    if not cancel_import:
        populate_ticker_dropdown()
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功寫入 {inserted_count} 筆資料。")

# 共用的小工具
def _parse_date(val) -> typing.Optional[datetime.date]:
    """將任何輸入轉成 datetime.date；轉換失敗回傳 None"""
    ts = pd.to_datetime(val, errors="coerce")
    return ts.date() if pd.notna(ts) else None

# --- 工具函式 ------------------------------------------------------------
def _extract_date_from_tv_code(tv_code: str) -> typing.Optional[datetime.date]:
    """
    若 TV Code 為「TICKER YYYYMMDD hhmmss TICKER: …」格式，
    取出中間的 YYYYMMDD 為日期；否則回傳 None
    """
    m = re.match(r'^[A-Za-z\.]+\s+(\d{8})\b', tv_code)
    if m:
        return pd.to_datetime(m.group(1), format='%Y%m%d').date()
    return None

# 先集中定義允許匯入的欄位
# ALLOWED_COLS = ['Open', 'High', 'Low', 'Close', 'TV Code']
ALLOWED_COLS = ['TV Code']

def _import_rows(ticker: str, df: pd.DataFrame, latest_date=None):
    """
    1. 每一列只看 TV Code 欄  
    2. 日期優先順序：TV Code 內嵌 > Date 欄  
    3. latest_date 仍用來過濾（以最終決定的日期比較）
    """
    if 'TV Code' not in df.columns:
        return

    for _, row in df.iterrows():
        tv_code = str(row['TV Code']).strip()
        if not tv_code or tv_code.lower() == 'nan':
            continue

        # 先判斷日期
        date_obj = _extract_date_from_tv_code(tv_code)
        if date_obj is None:                       # 沒嵌日期 → 用 Date 欄
            date_obj = _parse_date(row.get('Date'))
        if date_obj is None:
            continue
        if latest_date and date_obj < latest_date:
            continue

        parse_gex_code(date_obj.isoformat(), tv_code)

# --- 處理 Excel 匯入邏輯 ---
def process_excel(file_path):
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    user_conflict_choice = None
    apply_to_all = False
    cancel_import = False
    inserted_count = 0
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            _import_rows(sheet_name.strip(), df)          # ⬅️ 共用
        return True
    except Exception as e:
        messagebox.showerror("匯入錯誤", str(e))
        return False
    
def import_from_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return
    if process_excel(file_path):
        populate_ticker_dropdown()
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功寫入 {inserted_count} 筆資料。")

def auto_import_from_google():
    """
    1. 讀取試算表所有工作表
    2. 僅匯入 Open/High/Low/Close/TV Code
    3. 針對各 ticker 只匯入 >= 資料庫最新日期 之後的資料
       （latest_date 透過 _import_rows 的 latest_date 參數過濾）
    """
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print("⚠️  未找到 service_account.json，已跳過自動匯入")
        return

    try:
        # 建立 Google Sheets 連線（程式其餘部分保持原樣）:contentReference[oaicite:1]{index=1}
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            SERVICE_ACCOUNT_FILE, scope)
        client = gspread.authorize(creds)

        # 覆蓋模式、重置計數
        user_conflict_choice = None
        apply_to_all = False
        cancel_import = False
        inserted_count = 0

        sheet_ids = [SHEET_ID, SHEET_ID_MINOR]
        for s_id in sheet_ids:
            try:
                spreadsheet = client.open_by_key(s_id)
            except Exception as e:
                print(f"⚠️  無法開啟試算表 {s_id}: {e}")
                continue

            for ws in spreadsheet.worksheets():
                ticker = ws.title.strip()
                latest_date = get_latest_date_for_ticker(ticker)  # 可能為 None
                
                try:
                    # 使用 get_all_values() 獲取所有數據，然後手動處理標題
                    all_values = ws.get_all_values()
                    if not all_values:
                        print(f"⚠️  工作表 '{ticker}' 為空，跳過")
                        continue
                    
                    # 第一行作為標題
                    headers = all_values[0]
                    # 檢查標題是否有效
                    if not headers or all(not h.strip() for h in headers):
                        print(f"⚠️  工作表 '{ticker}' 標題行為空，跳過")
                        continue
                    
                    # 創建DataFrame，處理可能的標題重複
                    df = pd.DataFrame(all_values[1:], columns=headers)
                    
                except Exception as sheet_error:
                    print(f"⚠️  工作表 '{ticker}' 讀取失敗，跳過：{str(sheet_error)}")
                    continue
                
                # 檢查是否有必要的欄位
                if 'TV Code' not in df.columns:
                    print(f"⚠️  工作表 '{ticker}' 缺少 'TV Code' 欄位，跳過")
                    continue
                    
                _import_rows(ticker, df, latest_date)             # ⬅️ 共用

        if inserted_count:
            populate_ticker_dropdown()
            refresh_table()
            messagebox.showinfo("已從 Google Sheet 更新完成", f"成功寫入 {inserted_count} 筆資料。")
            print(f"✅ 自動匯入完成，共寫入 {inserted_count} 筆資料")
        else:
            messagebox.showinfo("完成", "無新資料")
            print("ℹ️  自動匯入：無新資料")
    except Exception as e:
        err = traceback.format_exc()
        messagebox.showerror("自動匯入錯誤", f"詳細錯誤:\n{err}")

# --- 從 Google 試算表匯入（使用 Service Account 認證） ---
def import_from_google():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print("⚠️  未找到 service_account.json，已取消匯入")
        return

    try:
        # 連線 Google Sheets
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            SERVICE_ACCOUNT_FILE, scope)
        client = gspread.authorize(creds)
        
        # Initialize counters manually
        global user_conflict_choice, apply_to_all, cancel_import, inserted_count
        user_conflict_choice = None
        apply_to_all = False
        cancel_import = False
        inserted_count = 0

        sheet_ids = [SHEET_ID, SHEET_ID_MINOR]
        
        for s_id in sheet_ids:
            try:
                spreadsheet = client.open_by_key(s_id)
                for ws in spreadsheet.worksheets():
                    try:
                        records = ws.get_all_records()
                        df = pd.DataFrame(records)
                        _import_rows(ws.title.strip(), df)
                    except Exception as inner_e:
                        print(f"⚠️  工作表 '{ws.title}' 讀取失敗 (ID: {s_id})：{inner_e}")
            except Exception as e:
                print(f"⚠️  無法開啟試算表 {s_id}: {e}")
                continue

        populate_ticker_dropdown()
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功寫入 {inserted_count} 筆資料。")

    except APIError as e:
        messagebox.showerror("API 錯誤", f"Google Sheets API 尚未啟用：\n{e.response.text}")
    except PermissionError as e:
        cause = e.__cause__ or e
        messagebox.showerror("授權錯誤", f"服務帳戶無權存取試算表：\n{cause}")
    except Exception as e:
        err = traceback.format_exc()
        messagebox.showerror("匯入錯誤", f"詳細錯誤:\n{err}")

def fetch_data(filter_ticker="", start_date=None, end_date=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    query = "SELECT * FROM stock_data WHERE 1=1"
    params = []
    if filter_ticker:
        query += " AND ticker = ?"
        params.append(filter_ticker)
    if start_date and end_date:
        query += " AND date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
    query += " ORDER BY date DESC"
    cursor.execute(query, params)
    data = cursor.fetchall()
    conn.close()
    return data


def delete_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("錯誤", "請選擇要刪除的記錄")
        return
    for item in selected:
        values = tree.item(item, "values")
        ticker, date, label, value = values
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM stock_data WHERE ticker=? AND date=? AND label=? AND value=?", (ticker, date, label, value))
        conn.commit()
        conn.close()
        tree.delete(item)


def refresh_table():
    for row in tree.get_children():
        tree.delete(row)
    selected_ticker = ticker_filter.get()
    start_date = start_date_filter.entry.get()
    end_date = end_date_filter.entry.get()
    if start_date and end_date:
        data = fetch_data(filter_ticker=selected_ticker, start_date=start_date, end_date=end_date)
    else:
        data = fetch_data(filter_ticker=selected_ticker)
    for row in data:
        tree.insert("", "end", values=row[1:])


def populate_ticker_dropdown():
    global all_tickers
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT ticker FROM stock_data")
    all_tickers = sorted([r[0] for r in cursor.fetchall()])
    conn.close()
    if all_tickers:
        ticker_filter['values'] = all_tickers
        # 如果當前有值且在列表中，保持不變；否則設為第一個
        current = ticker_filter.get()
        if not current and all_tickers:
            ticker_filter.set(all_tickers[0])

# 資料庫 helper：取得所有 ticker
def get_all_tickers():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT ticker FROM stock_data")
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows]

# 資料庫 helper：刪除既有當日 OHLC
def delete_ohlc(ticker, date_str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "DELETE FROM stock_data WHERE ticker = ? AND date = ? AND label IN ('Open','High','Low','Close')",
        (ticker, date_str)
    )
    conn.commit()
    conn.close()

# 更新 OHLC 按鈕的 callback
def update_ohlc(selected_date_input):
    """
    根據 calendar 選擇的日期，批次迴圈所有 ticker，
    從 yfinance 批次下載該日的 OHLC 資料，並寫入資料庫。
    支援傳入 DateEntry 物件、widget、字串或日期。
    """
    try:
        raw = selected_date_input
        # 處理不同 DateEntry 類型
        from ttkbootstrap.widgets import DateEntry as TtkDateEntry
        from tkcalendar import DateEntry as TkDateEntry
        if isinstance(raw, TtkDateEntry):
            raw = raw.entry.get()
        elif isinstance(raw, TkDateEntry):
            raw = raw.get_date()
        elif hasattr(raw, 'get_date'):
            raw = raw.get_date()
        elif hasattr(raw, 'entry') and hasattr(raw.entry, 'get'):
            raw = raw.entry.get()
        elif hasattr(raw, 'get'):
            raw = raw.get()

        # 轉為 datetime.date
        date = pd.to_datetime(raw).date()
        next_day = date + pd.Timedelta(days=1)

        tickers = get_all_tickers()
        # 根據 TV market 指數欄位前綴，並處理 BRK.B -> BRK-B
        ticker_map = {}
        ticker_names = []
        for t in tickers:
            if t in ["SPX", "NDX", "VIX"]:
                yf_name = f"^{t}"
            else:
                yf_name = t.replace(".", "-")
            ticker_names.append(yf_name)
            ticker_map[t] = yf_name

        # 批次下載所有 ticker 的單日資料
        df = yf.download(
            tickers=ticker_names,
            start=date,
            end=next_day,
            interval="1d",
            group_by="ticker",
            progress=False,
            auto_adjust=False
        )

        count = 0
        for t in tickers:
            name = ticker_map[t]
            # 若為多層索引，取該 ticker 區段，否則為單一 DataFrame
            # 注意：若只下載一支股票，yfinance 可能不會回傳 MultiIndex，
            # 但若 ticker_names 只有一個元素，df 就是那支股票的資料
            if len(ticker_names) > 1 and isinstance(df.columns, pd.MultiIndex):
                try:
                    data = df[name]
                except KeyError:
                    continue
            else:
                data = df
            
            if data.empty:
                continue
            try:
                row = data.iloc[0]
                ohlc = {
                    'Open': float(row['Open']),
                    'High': float(row['High']),
                    'Low': float(row['Low']),
                    'Close': float(row['Close'])
                }
            except (IndexError, ValueError, KeyError):
                continue

            # 刪除舊資料後插入
            delete_ohlc(t, str(date))
            for label, value in ohlc.items():
                insert_data(t, str(date), label, value)
            count += 1

        refresh_table()

        messagebox.showinfo("更新完成", f"已成功為 {count} 支 ticker 更新 {date} 的 OHLC 資料。")
    except Exception as e:
        messagebox.showerror("更新失敗", str(e))
                
# 資料庫 helper：從資料庫抓取指定 ticker 的所有 OHLC 歷史資料
def fetch_historical_ohlc_from_db(ticker):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(
        "SELECT date, label, value FROM stock_data WHERE ticker = ? AND label IN ('Open','High','Low','Close')",
        conn,
        params=(ticker,)
    )
    conn.close()
    if df.empty:
        return pd.DataFrame()
    df['date'] = pd.to_datetime(df['date'])
    pivot = df.pivot(index='date', columns='label', values='value')
    return pivot.sort_index()

# --- 新增函式：依篩選條件的 ticker & 日期區間，批次更新每天的 OHLC ---
def update_ohlc_range():
    try:
        t = ticker_filter.get().strip()
        start = start_date_filter.entry.get().strip()
        end = end_date_filter.entry.get().strip()
        if not t or not start or not end:
            messagebox.showwarning("參數不足", "請先選擇 Ticker 及起始/結束日期")
            return

        # 下載期間內的日線資料
        if t in ["SPX", "NDX", "VIX"]:
            yf_ticker = f"^{t}"
        else:
            yf_ticker = t.replace(".", "-")

        df = yf.download(
            tickers=yf_ticker,
            start=start,
            end=pd.to_datetime(end) + pd.Timedelta(days=1),
            interval="1d",
            group_by="ticker",
            progress=False,
            auto_adjust=False
        )

        if df.empty:
            messagebox.showinfo("無資料", f"{t} 在 {start} 到 {end} 期間無任何日線資料")
            return

        # 單 ticker 可能回傳單層 DataFrame
        data = df if not isinstance(df.columns, pd.MultiIndex) else df[yf_ticker]

        updated = 0
        for idx, row in data.iterrows():
            date_str = idx.date().isoformat()
            ohlc = {
                'Open': float(row['Open']),
                'High': float(row['High']),
                'Low':  float(row['Low']),
                'Close':float(row['Close'])
            }
            # 刪除舊資料、寫入新資料
            delete_ohlc(t, date_str)
            for lbl, val in ohlc.items():
                insert_data(t, date_str, lbl, val)
            updated += 1

        refresh_table()
        messagebox.showinfo("更新完成", f"{t} 共更新 {updated} 天的 OHLC 資料 ({start} ~ {end})")
    except Exception as e:
        messagebox.showerror("更新失敗", str(e))

def plot_graph():
    selected_ticker = ticker_filter.get()
    if not selected_ticker:
        messagebox.showwarning("錯誤", "請選擇 Ticker")
        return
    data = fetch_data(filter_ticker=selected_ticker)
    if not data:
        messagebox.showwarning("錯誤", "無數據")
        return
    df = pd.DataFrame(data, columns=["id", "ticker", "date", "label", "value"])
    df.drop(columns=["id"], inplace=True)
    df["date"] = pd.to_datetime(df["date"])
    df.sort_values("date", inplace=True)

    fig = go.Figure()

    color_map = {
        'Call Dominate':  '#FFD700',
        'Call Wall':      '#FFA500',
        'Call Wall CE':   '#FF7F50',
        'Gamma Field':    "#D75BF6",
        'Gamma Field CE': "#EAA1F8",
        'Key Delta':      '#ADFF2F',
        'Gamma Flip':     "#CBCBCB",
        'Gamma Flip CE':  "#FFFFFF",
        'Put Wall CE':    '#FF1493',
        'Put Wall':       '#DC143C',
        'Put Dominate':   '#8B0000',
    }

    labels = list(color_map.keys())   # ❷ 用 color_map 的 key 當繪圖順序

    # 繪製 GEX 折線圖（排除指定 labels）
    for label in labels:
        # if label in exclude_labels:
        #     continue
        if label in df["label"].unique():
            subset = df[df["label"] == label]
            fig.add_trace(go.Scatter(
                x=subset["date"],
                y=subset["value"],
                mode="lines+markers",
                name=label,
                line=dict(color=color_map[label]),   # ← 指定線色
            ))

    # 從資料庫取得資料
    # 讀取 OHLC；若缺資料僅警告，不中斷
    ohlc_df = fetch_historical_ohlc_from_db(selected_ticker)
    has_ohlc = (not ohlc_df.empty) and all(col in ohlc_df.columns
                                           for col in ['Open', 'High', 'Low', 'Close'])
    if not has_ohlc:
        messagebox.showwarning("缺少 OHLC",
                               f"{selected_ticker} 無 OHLC 資料，圖表將僅顯示其他指標")
    else:
        # 繪製 OHLC
        fig.add_trace(go.Candlestick(
            x=ohlc_df.index,
            open=ohlc_df["Open"],
            high=ohlc_df["High"],
            low=ohlc_df["Low"],
            close=ohlc_df["Close"],
            name=f"{selected_ticker} OHLC"
        ))

    # 設置標題和軸標籤
    fig.update_layout(
        title=f"{selected_ticker} OHLC & Gex Level Line Chart",
        xaxis_title="Date",
        yaxis_title="Price",
        xaxis_rangeslider_visible=False,
        template='plotly_dark'
    )

    fig.show()

# --- 自定義 Combobox 類別 ---
class SearchCombobox(ttk.Combobox):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.bind('<KeyRelease>', self.on_keyrelease)
        self.bind('<FocusOut>', self.on_focusout)
        self.bind('<Return>', self.on_return)
        self._listbox_window = None
        self._listbox = None

    def on_keyrelease(self, event):
        # 忽略導航鍵與功能鍵
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Return', 'Escape', 'Tab', 
                            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R']:
            return

        value = self.get()
        if not value:
            filtered = all_tickers
        else:
            filtered = [t for t in all_tickers if value.lower() in t.lower()]
        
        # 更新原生 values 但不自動 Post (避免游標跳動)
        self['values'] = filtered
        
        if filtered:
            self.show_listbox(filtered)
        else:
            self.hide_listbox()

    def show_listbox(self, values):
        if not self._listbox_window:
            self._listbox_window = tk.Toplevel(self)
            self._listbox_window.wm_overrideredirect(True)
            self._listbox_window.wm_attributes("-topmost", True)
            
            # 嘗試獲取主題顏色
            style = ttk.Style()
            try:
                bg = style.lookup('TEntry', 'fieldbackground')
                fg = style.lookup('TEntry', 'foreground')
                sel_bg = style.lookup('TEntry', 'selectbackground')
                sel_fg = style.lookup('TEntry', 'selectforeground')
            except:
                bg = 'white'
                fg = 'black'
                sel_bg = 'blue'
                sel_fg = 'white'
            
            self._listbox = tk.Listbox(self._listbox_window, height=5, bg=bg, fg=fg, 
                                       selectbackground=sel_bg, selectforeground=sel_fg, 
                                       relief="flat", borderwidth=1)
            self._listbox.pack(fill='both', expand=True)
            self._listbox.bind("<<ListboxSelect>>", self.on_select)
        
        self._listbox.delete(0, 'end')
        for item in values:
            self._listbox.insert('end', item)
            
        # 計算位置與大小
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        w = self.winfo_width()
        h = min(len(values), 10) * 20 # 估算高度
        self._listbox_window.geometry(f"{w}x{h}+{x}+{y}")
        self._listbox_window.deiconify()
        self._listbox_window.lift()

    def hide_listbox(self):
        if self._listbox_window:
            self._listbox_window.destroy()
            self._listbox_window = None
            self._listbox = None

    def on_select(self, event):
        if self._listbox and self._listbox.curselection():
            index = self._listbox.curselection()[0]
            val = self._listbox.get(index)
            self.set(val)
            self.hide_listbox()
            self.icursor(tk.END)
            self.selection_clear()

    def on_focusout(self, event):
        # 延遲關閉，以便讓點擊 Listbox 的事件先觸發
        self.after(150, self.hide_listbox)

    def on_return(self, event):
        if self._listbox_window and self._listbox and self._listbox.size() > 0:
             self.set(self._listbox.get(0))
             self.hide_listbox()
             self.icursor(tk.END)

# --- GUI 建構 ---
def build_gui():
    global root, calendar_date, gex_entry, ticker_filter, start_date_filter, end_date_filter, tree

    root = ttk.Window(themename="darkly")
    root.title("股票 GEX 管理系統")
    root.geometry("860x700")

    main = ttk.Frame(root, padding=10)
    main.pack(fill=BOTH, expand=True)

    import_menu = ttk.Menubutton(main, text="批次匯入", bootstyle=PRIMARY)
    import_menu.grid(row=0, column=0, sticky=W, pady=5)
    menu = ttk.Menu(import_menu)
    import_menu["menu"] = menu
    menu.add_command(label="從 TXT 匯入", command=bulk_import)
    menu.add_command(label="從 Excel 匯入", command=import_from_excel)
    menu.add_command(label="從 Google 試算表匯入", command=import_from_google)
    ttk.Button(main, text="從 Google sheet 更新最新 data", bootstyle=SUCCESS, command=auto_import_from_google).grid(row=0, column=1, sticky=W, pady=5)

    entry_frame = ttk.LabelFrame(main, text="單筆輸入", padding=10)
    entry_frame.grid(row=1, column=0, columnspan=2, sticky=W+E, pady=10)
    ttk.Label(entry_frame, text="日期:").grid(row=0, column=0, sticky=E)
    calendar_date = DateEntry(entry_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    calendar_date.grid(row=0, column=1, padx=5, sticky=W)

    ttk.Label(entry_frame, text="GEX Code:").grid(row=1, column=0, sticky=E)
    gex_entry = ttk.Entry(entry_frame, width=50)
    gex_entry.grid(row=1, column=1, padx=5, sticky=W)

    ttk.Button(entry_frame, text="新增記錄", bootstyle=SUCCESS, command=single_entry).grid(row=2, column=1, sticky=W, pady=5)
    btn_update_ohlc = ttk.Button(entry_frame, text="更新當日 OHLC", bootstyle=WARNING, command=lambda: update_ohlc(calendar_date))
    btn_update_ohlc.grid(row=2, column=1, padx=5, pady=5)

    filter_frame = ttk.LabelFrame(main, text="篩選條件", padding=10)
    filter_frame.grid(row=2, column=0, columnspan=2, sticky=W+E)

    ttk.Label(filter_frame, text="Ticker:").grid(row=0, column=0, sticky=E)
    ticker_filter = SearchCombobox(filter_frame)
    ticker_filter.grid(row=0, column=1, padx=5, sticky=W)
    # ticker_filter.bind('<KeyRelease>', on_ticker_type) # 已整合至 SearchCombobox 類別中

    ttk.Label(filter_frame, text="起始日期:").grid(row=1, column=0, sticky=E)
    start_date_filter = DateEntry(filter_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    start_date_filter.entry.delete(0, 'end')
    start_date_filter.grid(row=1, column=1, padx=5, sticky=W)

    ttk.Label(filter_frame, text="結束日期:").grid(row=2, column=0, sticky=E)
    end_date_filter = DateEntry(filter_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    end_date_filter.entry.delete(0, 'end')
    end_date_filter.grid(row=2, column=1, padx=5, sticky=W)

    ttk.Button(filter_frame, text="篩選", bootstyle=INFO, command=refresh_table).grid(row=3, column=1, sticky=W, pady=5)
    # 新增：依篩選條件批次更新 OHLC
    ttk.Button(filter_frame, text="更新 OHLC 區間", bootstyle=WARNING,
               command=update_ohlc_range).grid(row=3, column=1, padx=100, sticky=W, pady=5)
    ttk.Button(filter_frame, text="重置", bootstyle=DANGER, command=lambda: [ticker_filter.set(""), start_date_filter.entry.delete(0, 'end'), end_date_filter.entry.delete(0, 'end'), refresh_table()]).grid(row=3, column=1, padx=320, sticky=W, pady=5)

    table_frame = ttk.LabelFrame(main, text="資料表格", padding=10)
    table_frame.grid(row=3, column=0, columnspan=2, sticky=NSEW, pady=10)
    tree = ttk.Treeview(table_frame, columns=("Ticker", "date", "Label", "Value"), show="headings")
    tree.heading("Ticker", text="Ticker")
    tree.heading("date", text="Date")
    tree.heading("Label", text="Label")
    tree.heading("Value", text="Value")
    tree.pack(side=LEFT, fill=BOTH, expand=True)

    # 在 tree.pack() 之後，先確保 Treeview 有焦點
    tree.bind('<Button-1>', lambda e: tree.focus_set())

    def copy_selected_values(event=None):
        sel = tree.selection()
        if not sel:
            return
        # 取出每一列的第 4 欄（index=3），也就是 Value
        vals = [tree.item(item, 'values')[3] for item in sel]
        text = '\n'.join(str(v) for v in vals)
        # 複製到剪貼簿
        root.clipboard_clear()
        root.clipboard_append(text)
        # 若要顯示提示可取消下行
        # messagebox.showinfo("複製完成", f"已複製 {len(vals)} 筆 Value")
    
    root.bind_all('<Command-c>', copy_selected_values)
    root.bind_all('<Command-C>', copy_selected_values)
    root.bind_all('<Control-c>', copy_selected_values)
    root.bind_all('<Control-C>', copy_selected_values)

    scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=RIGHT, fill=Y)

    btn_frame = ttk.Frame(main)
    btn_frame.grid(row=4, column=0, columnspan=2, pady=10)
    ttk.Button(btn_frame, text="📈 繪製圖表", bootstyle=PRIMARY, command=plot_graph).grid(row=0, column=0, padx=5)
    ttk.Button(btn_frame, text="🗑️ 刪除選定", bootstyle=DANGER, command=delete_selected).grid(row=0, column=1, padx=5)

    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    main.columnconfigure(1, weight=1)
    table_frame.columnconfigure(0, weight=1)

    populate_ticker_dropdown()
    refresh_table()

    root.mainloop()

if __name__ == "__main__":
    init_db()
    build_gui()
