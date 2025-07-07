import sqlite3
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import plotly.graph_objects as go
import yfinance as yf
import re
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry

# 全域控件
root = None
calendar_date = None
gex_entry = None
ticker_filter = None
start_date_filter = None
end_date_filter = None
tree = None

# 全域狀態追蹤使用者選擇
user_conflict_choice = None  # "skip" 或 "overwrite" 或 "cancel"
apply_to_all = False
cancel_import = False
inserted_count = 0

# --- 資料庫相關 ---
def init_db():
    conn = sqlite3.connect("stocks.db")
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

    conn = sqlite3.connect("stocks.db")
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


def import_from_excel():
    global user_conflict_choice, apply_to_all, cancel_import, inserted_count
    user_conflict_choice = None
    apply_to_all = False
    cancel_import = False
    inserted_count = 0

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return

    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if 'Date' in df.columns and 'TV Code' in df.columns:
                for date, tv_code in zip(df['Date'].dropna(), df['TV Code'].dropna()):
                    if cancel_import:
                        messagebox.showinfo("已取消", f"成功寫入 {inserted_count} 筆資料。")
                        return
                    parse_gex_code(str(date).split(" ")[0], str(tv_code))

        if not cancel_import:
            populate_ticker_dropdown()
            refresh_table()
            messagebox.showinfo("匯入完成", f"成功寫入 {inserted_count} 筆資料。")
    except Exception as e:
        messagebox.showerror("匯入錯誤", str(e))


def fetch_data(filter_ticker="", start_date=None, end_date=None):
    conn = sqlite3.connect("stocks.db")
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
        conn = sqlite3.connect("stocks.db")
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
    conn = sqlite3.connect("stocks.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT ticker FROM stock_data")
    tickers = [r[0] for r in cursor.fetchall()]
    conn.close()
    if tickers:
        ticker_filter['values'] = tickers
        ticker_filter.set(tickers[0])


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

    # 下載歷史數據
    ticker_name = f"^{selected_ticker}" if selected_ticker in ["SPX", "NDX", "VIX"] else selected_ticker
    value_3m = yf.download(ticker_name, period="3mo", interval="1d", multi_level_index=False)
    value_3m = value_3m.reset_index()
    value_3m.rename(columns={"Date": "date"}, inplace=True)
    value_3m["date"] = pd.to_datetime(value_3m["date"])

    fig = go.Figure()

    # 繪製 GEX 折線圖
    for label in df["label"].unique():
        subset = df[df["label"] == label]
        fig.add_trace(go.Scatter(x=subset["date"], y=subset["value"], mode="lines+markers", name=label))

    # 繪製 OHLC 圖表
    fig.add_trace(go.Candlestick(
        x=value_3m["date"],
        open=value_3m["Open"],
        high=value_3m["High"],
        low=value_3m["Low"],
        close=value_3m["Close"],
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

    entry_frame = ttk.LabelFrame(main, text="單筆輸入", padding=10)
    entry_frame.grid(row=1, column=0, columnspan=2, sticky=W+E, pady=10)
    ttk.Label(entry_frame, text="日期:").grid(row=0, column=0, sticky=E)
    calendar_date = DateEntry(entry_frame, bootstyle="dark", dateformat="%Y-%m-%d")
    calendar_date.grid(row=0, column=1, padx=5, sticky=W)

    ttk.Label(entry_frame, text="GEX Code:").grid(row=1, column=0, sticky=E)
    gex_entry = ttk.Entry(entry_frame, width=50)
    gex_entry.grid(row=1, column=1, padx=5, sticky=W)

    ttk.Button(entry_frame, text="新增記錄", bootstyle=SUCCESS, command=single_entry).grid(row=2, column=1, sticky=W, pady=5)

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
