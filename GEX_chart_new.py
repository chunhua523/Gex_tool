import sqlite3
import tempfile
import tkinter as tk
from tkinter import messagebox, filedialog
import traceback
import pandas as pd
import plotly.graph_objects as go
import requests
import yfinance as yf
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry
from gspread.exceptions import APIError

# 如果需要從 Google Sheet 讀取私有試算表，請安裝 gspread 與 oauth2client
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
DB_PATH = "stocks.db"

# 全域狀態追蹤使用者選擇
user_conflict_choice = None  # "skip" 或 "overwrite" 或 "cancel"
apply_to_all = False
cancel_import = False
inserted_count = 0

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
def parse_gex_code(date, gex_code):
    global cancel_import
    if cancel_import:
        return None

    date = date.split(" ")[0]
    match = re.match(r"(\w+):\s*(.*)", gex_code)
    if not match:
        messagebox.showwarning("格式錯誤", "無法解析 GEX TV Code")
        return None
    ticker, code_body = match.groups()
    elements = re.split(r',\s*', code_body)
    i = 0
    while i < len(elements) - 1:
        labels = elements[i].strip()
        try:
            value = float(elements[i + 1].strip())
            for label in labels.split('&'):
                insert_data(ticker, date, label.strip(), value)
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

    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for i in range(0, len(lines), 2):
        if cancel_import:
            messagebox.showinfo("已取消", f"成功寫入 {inserted_count} 筆資料。")
            return
        date = lines[i].strip().split("_")[0]
        if i + 1 < len(lines):
            gex_code = lines[i + 1].strip()
            parse_gex_code(date, gex_code)

    if not cancel_import:
        populate_ticker_dropdown()
        refresh_table()
        messagebox.showinfo("匯入完成", f"成功寫入 {inserted_count} 筆資料。")

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
            if 'Date' not in df.columns:
                continue
            for _, row in df.iterrows():
                if cancel_import:
                    messagebox.showinfo("已取消", f"成功寫入 {inserted_count} 筆資料。")
                    return False
                date_val = row['Date']
                try:
                    date_str = str(pd.to_datetime(date_val).date())
                except:
                    date_str = str(date_val)
                for col in df.columns:
                    if col == 'Date':
                        continue
                    value = row[col]
                    if pd.isna(value):
                        continue
                    insert_data(sheet_name, date_str, col, value)
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

# --- 從 Google 試算表匯入（使用 Service Account 認證） ---
def import_from_google():
    creds_file = filedialog.askopenfilename(title="選擇 Service Account JSON", filetypes=[("JSON", "*.json")])
    if not creds_file: return
    SHEET_ID, GID = '1u1opYwj_2bhOBhAM96CB7kYz9prWKQtXhmjU1cG15Dg', '1369484985'
    try:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(SHEET_ID)
        sheet = next((ws for ws in spreadsheet.worksheets() if str(ws.id)==GID), None)
        if sheet is None: raise Exception("找不到指定工作表")
        df = pd.DataFrame(sheet.get_all_records())
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            df.to_excel(tmp.name, index=False, sheet_name='GoogleSheet')
            tmp_path = tmp.name
        if not process_excel(tmp_path): raise Exception("Google 試算表匯入失敗")
        populate_ticker_dropdown(); refresh_table()
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
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT ticker FROM stock_data")
    tickers = [r[0] for r in cursor.fetchall()]
    conn.close()
    if tickers:
        ticker_filter['values'] = tickers
        ticker_filter.set(tickers[0])

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
        # 根據 TV market 指數欄位前綴
        ticker_names = [f"^{t}" if t in ["SPX","NDX","VIX"] else t for t in tickers]

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
            name = f"^{t}" if t in ["SPX","NDX","VIX"] else t
            # 若為多層索引，取該 ticker 區段，否則為單一 DataFrame
            data = df[name] if isinstance(df.columns, pd.MultiIndex) else df
            if data.empty:
                continue
            row = data.iloc[0]
            ohlc = {
                'Open': float(row['Open']),
                'High': float(row['High']),
                'Low': float(row['Low']),
                'Close': float(row['Close'])
            }
            # 刪除舊資料後插入
            delete_ohlc(t, str(date))
            for label, value in ohlc.items():
                insert_data(t, str(date), label, value)
            count += 1

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

    # 要排除不畫出的資料
    exclude_labels = ['Open', 'High', 'Low', 'Close', 'Flip %', 'TV Code']

    # 繪製 GEX 折線圖（排除指定 labels）
    for label in df["label"].unique():
        if label in exclude_labels:
            continue
        subset = df[df["label"] == label]
        fig.add_trace(go.Scatter(
            x=subset["date"],
            y=subset["value"],
            mode="lines+markers",
            name=label
        ))

    # 從資料庫取得資料
    ohlc_df = fetch_historical_ohlc_from_db(selected_ticker)
    if ohlc_df.empty or any(col not in ohlc_df.columns for col in ['Open','High','Low','Close']):
        messagebox.showwarning("缺少資料", f"{selected_ticker} 沒有任何完整的 OHLC 歷史資料。")
        return
    
    # 繪製 OHLC 圖表
    fig.add_trace(go.Candlestick(
        x=ohlc_df.index,
        open=ohlc_df["Open"],
        high=ohlc_df["High"],
        low=ohlc_df["Low"],
        close=ohlc_df["Close"],
        name=selected_ticker + ' OHLC'
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

    entry_frame = ttk.LabelFrame(main, text="單筆輸入", padding=10)
    entry_frame.grid(row=1, column=0, columnspan=2, sticky=W+E, pady=10)
    ttk.Label(entry_frame, text="日期:").grid(row=0, column=0, sticky=E)
    calendar_date = DateEntry(entry_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    calendar_date.grid(row=0, column=1, padx=5, sticky=W)

    ttk.Label(entry_frame, text="GEX Code:").grid(row=1, column=0, sticky=E)
    gex_entry = ttk.Entry(entry_frame, width=50)
    gex_entry.grid(row=1, column=1, padx=5, sticky=W)

    ttk.Button(entry_frame, text="新增記錄", bootstyle=SUCCESS, command=single_entry).grid(row=2, column=1, sticky=W, pady=5)
    btn_update_ohlc = tk.Button(entry_frame, text="更新 OHLC", command=lambda: update_ohlc(calendar_date))
    btn_update_ohlc.grid(row=2, column=1, padx=5, pady=5)

    filter_frame = ttk.LabelFrame(main, text="篩選條件", padding=10)
    filter_frame.grid(row=2, column=0, columnspan=2, sticky=W+E)

    ttk.Label(filter_frame, text="Ticker:").grid(row=0, column=0, sticky=E)
    ticker_filter = ttk.Combobox(filter_frame, state="readonly")
    ticker_filter.grid(row=0, column=1, padx=5, sticky=W)

    ttk.Label(filter_frame, text="起始日期:").grid(row=1, column=0, sticky=E)
    start_date_filter = DateEntry(filter_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    start_date_filter.entry.delete(0, 'end')
    start_date_filter.grid(row=1, column=1, padx=5, sticky=W)

    ttk.Label(filter_frame, text="結束日期:").grid(row=2, column=0, sticky=E)
    end_date_filter = DateEntry(filter_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    end_date_filter.entry.delete(0, 'end')
    end_date_filter.grid(row=2, column=1, padx=5, sticky=W)

    ttk.Button(filter_frame, text="篩選", bootstyle=INFO, command=refresh_table).grid(row=3, column=1, sticky=W, pady=5)
    ttk.Button(filter_frame, text="重置", bootstyle=SECONDARY, command=lambda: [ticker_filter.set(""), start_date_filter.entry.delete(0, 'end'), end_date_filter.entry.delete(0, 'end'), refresh_table()]).grid(row=3, column=1, sticky=E, pady=5)

    table_frame = ttk.LabelFrame(main, text="資料表格", padding=10)
    table_frame.grid(row=3, column=0, columnspan=2, sticky=NSEW, pady=10)
    tree = ttk.Treeview(table_frame, columns=("Ticker", "date", "Label", "Value"), show="headings")
    tree.heading("Ticker", text="Ticker")
    tree.heading("date", text="Date")
    tree.heading("Label", text="Label")
    tree.heading("Value", text="Value")
    tree.pack(side=LEFT, fill=BOTH, expand=True)

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
    
    # 綁定 Ctrl+C（Windows/Linux）與 Command+C（macOS）
    tree.bind('<Control-c>', copy_selected_values)
    tree.bind('<Command-c>', copy_selected_values)

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
