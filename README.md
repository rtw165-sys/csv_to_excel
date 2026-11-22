# 📊 CSV 批次轉 Excel 工具（彩色版 / 合併模式）

一款以 **純函式架構（無 `class` 設計）** 打造的 Python Tkinter 桌面工具，用來指定目錄、勾選多個 `.csv` 檔案，並一鍵轉存為 Excel（`.xlsx`）報表。工具支援自動編碼偵測、自動適配中英文字元欄寬、合併多檔為單一活頁簿多分頁，並將成果固定輸出至執行檔所在目錄下的 `output/` 資料夾。

---

## ✨ 主要功能特色

* **指定目錄與勾選清單**：選擇包含 CSV 的資料夾後自動列出 `.csv` 清單，支援單檔勾選、**全選** 與 **取消全選**。
* **合併模式（預設啟用）**：將勾選的多個 CSV 寫入同一個 `CSV合併報表_YYYYMMDD_HHMMSS.xlsx`，每個 CSV 各自成為一個獨立工作表分頁（自動處理 Excel 分頁名稱 31 字元上限、非法字元與重名問題）；取消勾選則可改為各自獨立匯出 `.xlsx`。
* **輸出路徑固定在 `./output`**：一律輸出至這支程式所在目錄下的 `output/` 資料夾（透過 `Path(__file__).resolve().parent / "output"` 錨定絕對路徑），自動建立資料夾，不受目前工作目錄影響。
* **智慧編碼讀取**：依序嘗試 `utf-8-sig`、`utf-8`、`cp950`、`big5` 讀取 CSV（並以字串 `dtype=str` 讀入），避免繁體中文亂碼與門市店號、電話等前導零遺失。
* **自動適配欄寬**：依欄位標題與內容計算近似顯示寬度（中文字算 2、英數符號算 1），自動調整各欄寬度並設定上下限，開啟 Excel 不被截斷。
* **彩色高對比 UI 與純函式設計**：不使用 `class` 封裝，程式拆分為「路徑、讀取與 Excel 匯出函式」、「按鈕樣式與介面操作函式」、「主視窗元件」三段，結構清楚、易於閱讀與維護。

---

## 🛠️ 環境與套件需求

* **Python**：建議 Python 3.8 以上版本
* **必要套件**：`pandas`、`openpyxl`（`tkinter` 為 Python 標準函式庫，一般桌面版 Python 已內建）

安裝相依套件：

```bash
pip install pandas openpyxl
```

---

## 🚀 快速開始

1. 將主程式儲存為 `csv_to_excel_tool.py`。
2. 在終端機執行：

```bash
python csv_to_excel_tool.py
```

3. 在視窗中點擊 **「📁 選擇目錄」** 挑選 CSV 資料夾，確認要轉換的檔案勾選狀態與「合併模式」開關。
4. 點擊 **「🚀 一鍵轉換成 Excel」**，完成後即可在 `csv_to_excel_tool.py` 同層的 `output/` 資料夾中找到產出的 `.xlsx` 檔案。

---

## 📁 輸出檔案說明

* **合併模式（預設勾選）**：
  `./output/CSV合併報表_YYYYMMDD_HHMMSS.xlsx`（每個勾選的 `.csv` 對應一個分頁）
* **獨立匯出模式（取消勾選合併模式）**：
  `./output/<原檔名>_YYYYMMDD_HHMMSS.xlsx`

---

## 🔧 常見自訂調整位置

* **調整最大/最小欄寬或留白**：修改 `auto_fit_column_width(worksheet, df, padding=3, max_width=60)` 的參數。
* **調整 UI 配色**：修改主視窗區塊的色碼變數（如 `BG_MAIN`、`BG_CARD`）或各按鈕的 `bg` / `hover_bg` / `fg` 參數。