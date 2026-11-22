import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


# ==========================================
# 1. 路徑、讀取與 Excel 匯出函式
# ==========================================
def get_output_dir():
    """取得執行檔案所在目錄下的 output 資料夾（不存在則自動建立）"""
    base_dir = Path(__file__).resolve().parent
    out_dir = base_dir / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def read_csv_auto_encoding(csv_path):
    """依序嘗試常見編碼讀取 CSV，避免中文亂碼"""
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5"]
    for enc in encodings:
        try:
            return pd.read_csv(csv_path, encoding=enc, dtype=str)
        except Exception:
            continue
    return pd.read_csv(csv_path, dtype=str)


def calc_display_width(text):
    """計算字串在 Excel 中的近似顯示寬度（中文字算 2，英數符號算 1）"""
    if text is None:
        return 0
    s = str(text)
    width = 0
    for ch in s:
        if ord(ch) > 127:
            width += 2
        else:
            width += 1
    return width


def get_excel_col_letter(idx):
    """將 0 開始的欄位索引轉成 Excel 欄名（0 -> A, 25 -> Z, 26 -> AA）"""
    letters = ""
    while True:
        idx, rem = divmod(idx, 26)
        letters = chr(65 + rem) + letters
        if idx == 0:
            break
        idx -= 1
    return letters


def auto_fit_column_width(worksheet, df, padding=3, max_width=60):
    """自動適配工作表各欄寬度"""
    for col_idx, col_name in enumerate(df.columns):
        max_len = calc_display_width(col_name)

        col_series = df[col_name].fillna("")
        for val in col_series:
            val_len = calc_display_width(val)
            if val_len > max_len:
                max_len = val_len

        final_width = min(max_len + padding, max_width)
        col_letter = get_excel_col_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = max(final_width, 8)


def sanitize_sheet_name(raw_name, used_names):
    """整理 Excel 分頁名稱（最長 31 字元、移除非法字元、避免重複）"""
    invalid_chars = ["\\", "/", "?", "*", ":", "[", "]"]
    clean_name = raw_name
    for ch in invalid_chars:
        clean_name = clean_name.replace(ch, "_")
    clean_name = clean_name.strip() or "Sheet"
    clean_name = clean_name[:31]

    base = clean_name
    counter = 2
    while clean_name in used_names:
        suffix = f"_{counter}"
        clean_name = base[: 31 - len(suffix)] + suffix
        counter += 1

    used_names.add(clean_name)
    return clean_name


def export_separate_excels(selected_paths, output_dir):
    """獨立模式：每個勾選的 CSV 各自匯出成一個 .xlsx 檔"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    created_files = []

    for csv_path in selected_paths:
        p = Path(csv_path)
        df = read_csv_auto_encoding(csv_path)
        excel_path = output_dir / f"{p.stem}_{timestamp}.xlsx"

        used_names = set()
        sheet_name = sanitize_sheet_name(p.stem, used_names)

        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            auto_fit_column_width(worksheet, df)

        created_files.append(excel_path.name)

    return created_files


def export_merged_excel(selected_paths, output_dir):
    """合併模式：將勾選的多個 CSV 寫入同一個 CSV合併報表_時間戳記.xlsx"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"CSV合併報表_{timestamp}.xlsx"
    used_names = set()

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for csv_path in selected_paths:
            p = Path(csv_path)
            df = read_csv_auto_encoding(csv_path)
            sheet_name = sanitize_sheet_name(p.stem, used_names)

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            auto_fit_column_width(worksheet, df)

    return [excel_path.name]


# ==========================================
# 2. 按鈕樣式與介面操作函式
# ==========================================
def make_button(parent, text, bg, hover_bg, fg="#FFFFFF", command=None, font_size=10, padx=12, pady=5):
    """建立高對比且支援滑鼠懸停變色的按鈕"""
    btn = tk.Button(
        parent,
        text=text,
        bg=bg,
        fg=fg,
        activebackground=hover_bg,
        activeforeground=fg,
        font=("Microsoft JhengHei", font_size, "bold"),
        relief="flat",
        bd=0,
        cursor="hand2",
        padx=padx,
        pady=pady,
        command=command,
    )
    btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg, fg=fg))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg, fg=fg))
    return btn


def browse_folder():
    """選擇 CSV 所在目錄"""
    folder = filedialog.askdirectory(title="請選擇包含 CSV 的資料夾")
    if folder:
        entry_dir.delete(0, tk.END)
        entry_dir.insert(0, folder)
        load_csv_list()


def clear_checkboxes():
    """清空目前的勾選清單畫面"""
    for widget in frame_list_inner.winfo_children():
        widget.destroy()
    csv_items.clear()


def load_csv_list():
    """讀取指定目錄下的 .csv 檔案並建立勾選清單"""
    folder = entry_dir.get().strip()
    clear_checkboxes()

    if not folder or not Path(folder).is_dir():
        messagebox.showwarning("提示", "請先選擇有效的 CSV 目錄！")
        return

    files = sorted(
        f for f in Path(folder).iterdir()
        if f.is_file() and f.suffix.lower() == ".csv"
    )

    if not files:
        tk.Label(
            frame_list_inner,
            text="此目錄下找不到 .csv 檔案",
            font=("Microsoft JhengHei", 10),
            bg="#FFFFFF",
            fg="#64748B",
        ).pack(anchor="w", padx=12, pady=12)
        update_status("目錄中無 CSV 檔案", color="#B45309")
        return

    for idx, file_path in enumerate(files):
        row_bg = "#F8FAFC" if idx % 2 == 1 else "#FFFFFF"
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(
            frame_list_inner,
            text=f"  {file_path.name}",
            variable=var,
            font=("Microsoft JhengHei", 10),
            bg=row_bg,
            fg="#1E293B",
            activebackground="#EFF6FF",
            activeforeground="#1E3A8A",
            selectcolor="#DBEAFE",
            anchor="w",
            padx=6,
            pady=3,
        )
        chk.pack(fill="x", padx=4, pady=1)
        csv_items.append((str(file_path), var))

    update_status(f"已載入 {len(csv_items)} 個 CSV 檔案", color="#0F766E")


def select_all(checked):
    """全選或取消全選"""
    for _, var in csv_items:
        var.set(checked)


def convert_selected():
    """依照目前模式（合併模式 / 獨立匯出）轉換已勾選的 CSV"""
    selected_paths = [path for path, var in csv_items if var.get()]
    if not selected_paths:
        messagebox.showwarning("尚未勾選", "請至少勾選一個要轉換的 CSV 檔案！")
        return

    output_dir = get_output_dir()
    is_merge = merge_var.get()

    try:
        if is_merge:
            created_files = export_merged_excel(selected_paths, output_dir)
        else:
            created_files = export_separate_excels(selected_paths, output_dir)

        mode_label = "合併模式（多分頁）" if is_merge else "獨立匯出模式"
        summary = (
            f"模式：{mode_label}\n"
            f"產出檔案數：{len(created_files)}\n"
            f"輸出目錄：{output_dir}\n"
            f"檔案清單：\n- " + "\n- ".join(created_files)
        )
        messagebox.showinfo("轉換完成", summary)
        update_status(f"轉換完成（{mode_label}）→ {output_dir}", color="#047857")

    except Exception as e:
        messagebox.showerror("轉換失敗", f"發生錯誤：\n{e}")
        update_status("轉換失敗", color="#B91C1C")


def update_status(msg, color="#1E293B"):
    """更新底部狀態列文字與顏色"""
    lbl_status.config(text=f"狀態：{msg}", fg=color)


def on_mousewheel(event):
    """讓滑鼠滾輪可以捲動檔案清單"""
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def on_canvas_configure(event):
    """同步調整捲動區域寬度"""
    canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.itemconfig(canvas_window_id, width=event.width)


# ==========================================
# 3. 建立彩色現代化主視窗與元件
# ==========================================
BG_MAIN = "#F0F4F8"
BG_CARD = "#FFFFFF"
TEXT_DARK = "#0F172A"
TEXT_MUTED = "#475569"
BORDER_COLOR = "#CBD5E1"

root = tk.Tk()
root.title("CSV 批次轉 Excel 工具（彩色版 / 合併模式 / 自動適配欄寬）")
root.geometry("720x540")
root.configure(bg=BG_MAIN)

csv_items = []
merge_var = tk.BooleanVar(value=True)  # 預設啟用「合併模式」

# 頂部標題橫幅
banner_top = tk.Frame(root, bg="#1D4ED8", padx=16, pady=10)
banner_top.pack(fill="x")

tk.Label(
    banner_top,
    text="📊 CSV 批次轉 Excel 工具",
    font=("Microsoft JhengHei", 14, "bold"),
    bg="#1D4ED8",
    fg="#FFFFFF",
).pack(side="left")

tk.Label(
    banner_top,
    text="自動編碼辨識 · 自動適配欄寬 · 輸出至 ./output",
    font=("Microsoft JhengHei", 9, "bold"),
    bg="#DBEAFE",
    fg="#1E3A8A",
    padx=8,
    pady=2,
).pack(side="right")

# 上方卡片：CSV 目錄選擇區
frame_top = tk.Frame(
    root,
    bg=BG_CARD,
    padx=12,
    pady=10,
    highlightbackground=BORDER_COLOR,
    highlightthickness=1,
)
frame_top.pack(fill="x", padx=12, pady=(12, 8))

tk.Label(
    frame_top,
    text="CSV 目錄：",
    font=("Microsoft JhengHei", 10, "bold"),
    bg=BG_CARD,
    fg=TEXT_DARK,
).pack(side="left")

entry_dir = tk.Entry(
    frame_top,
    font=("Microsoft JhengHei", 10),
    bg="#F8FAFC",
    fg=TEXT_DARK,
    relief="solid",
    bd=1,
)
entry_dir.pack(side="left", fill="x", expand=True, padx=8, ipady=3)

make_button(
    frame_top,
    text="📁 選擇目錄",
    bg="#3B82F6",
    hover_bg="#2563EB",
    fg="#FFFFFF",
    command=browse_folder,
).pack(side="left", padx=3)

make_button(
    frame_top,
    text="🔄 重新載入",
    bg="#0EA5E9",
    hover_bg="#0284C7",
    fg="#FFFFFF",
    command=load_csv_list,
).pack(side="left", padx=3)

# 中間控制列：合併模式開關 + 全選 / 取消全選
frame_actions = tk.Frame(
    root,
    bg="#EFF6FF",
    padx=12,
    pady=8,
    highlightbackground="#BFDBFE",
    highlightthickness=1,
)
frame_actions.pack(fill="x", padx=12, pady=(0, 8))

tk.Checkbutton(
    frame_actions,
    text="合併模式（預設：勾選的 CSV 寫入同一個 CSV合併報表_時間戳記.xlsx）",
    variable=merge_var,
    font=("Microsoft JhengHei", 10, "bold"),
    bg="#EFF6FF",
    fg="#065F46",
    activebackground="#EFF6FF",
    activeforeground="#065F46",
    selectcolor="#D1FAE5",
).pack(side="left")

make_button(
    frame_actions,
    text="取消全選",
    bg="#E2E8F0",
    hover_bg="#CBD5E1",
    fg="#1E293B",
    font_size=9,
    padx=10,
    pady=4,
    command=lambda: select_all(False),
).pack(side="right", padx=3)

make_button(
    frame_actions,
    text="全選",
    bg="#DBEAFE",
    hover_bg="#93C5FD",
    fg="#1E3A8A",
    font_size=9,
    padx=10,
    pady=4,
    command=lambda: select_all(True),
).pack(side="right", padx=3)

# 中間卡片：可捲動的勾選清單區
frame_list_outer = tk.LabelFrame(
    root,
    text=" 勾選要轉換的 CSV 檔案 ",
    font=("Microsoft JhengHei", 10, "bold"),
    bg=BG_CARD,
    fg="#1E3A8A",
    highlightbackground=BORDER_COLOR,
    highlightthickness=1,
    padx=6,
    pady=6,
)
frame_list_outer.pack(fill="both", expand=True, padx=12, pady=(0, 8))

canvas = tk.Canvas(frame_list_outer, bg=BG_CARD, highlightthickness=0)
scrollbar = tk.Scrollbar(frame_list_outer, orient="vertical", command=canvas.yview)
frame_list_inner = tk.Frame(canvas, bg=BG_CARD)

canvas_window_id = canvas.create_window((0, 0), window=frame_list_inner, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

canvas.bind("<Configure>", on_canvas_configure)
frame_list_inner.bind("<MouseWheel>", on_mousewheel)
canvas.bind("<MouseWheel>", on_mousewheel)

# 下方卡片：一鍵轉換按鈕與彩色狀態列
frame_bottom = tk.Frame(
    root,
    bg=BG_CARD,
    padx=12,
    pady=10,
    highlightbackground=BORDER_COLOR,
    highlightthickness=1,
)
frame_bottom.pack(fill="x", padx=12, pady=(0, 12))

make_button(
    frame_bottom,
    text="🚀 一鍵轉換成 Excel",
    bg="#059669",
    hover_bg="#047857",
    fg="#FFFFFF",
    font_size=11,
    padx=18,
    pady=7,
    command=convert_selected,
).pack(side="right")

lbl_status = tk.Label(
    frame_bottom,
    text=f"狀態：輸出目錄為 {get_output_dir()}",
    font=("Microsoft JhengHei", 9, "bold"),
    bg=BG_CARD,
    fg=TEXT_MUTED,
    anchor="w",
    justify="left",
)
lbl_status.pack(side="left", fill="x", expand=True)

root.mainloop()