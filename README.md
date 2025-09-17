# 📦 調貨建議生成系統 v1.7

本系統是一個基於Streamlit的Web應用，旨在幫助零售企業根據庫存、銷量和安全庫存等數據，智能生成店鋪間的調貨建議，以優化庫存分配，減少缺貨和庫存積壓。

## ✨ 核心功能

- **智能調貨建議**：根據預設的業務規則，自動識別可轉出和應接收的商品及店鋪。
- **A/B 模式選擇**：支持“保守轉貨”與“加強轉貨”兩種模式，以應對不同的業務需求。
- **數據驅動**：基於上傳的Excel數據文件進行分析。
- **可視化分析**：提供圖表展示各營運單位的調出與接收情況。
- **報告匯出**：可將生成的調貨建議和統計摘要匯出為符合規範的雙工作表Excel文件。

## 🚀 如何運行

1.  **安裝依賴**：
    ```bash
    pip install -r requirements.txt
    ```

2.  **運行應用**：
    - Windows:
      ```bash
      run.bat
      ```
    - macOS / Linux:
      ```bash
      chmod +x run.sh
      ./run.sh
      ```

3.  **訪問應用**：
    在瀏覽器中打開應用啟動後顯示的URL（通常是 `http://localhost:8501`）。

## 🧪 如何測試

本專案包含一個自動化測試腳本，用於驗證核心功能的正確性。

1.  **運行測試**：
    ```bash
    python test_chart_v1.7.py
    ```
    此腳本會：
    - 如果 `test_data_v1.7.xlsx` 不存在，則會自動生成。
    - 執行數據預處理。
    - 在“保守”和“加強”兩種模式下生成調貨建議。
    - 創建並保存在兩種模式下的分析圖表 (`test_chart_A.png`, `test_chart_B.png`)。

## 🛠️ 技術棧

- **前端**：Streamlit (>=1.28.0)
- **數據處理**：pandas (>=2.0.0), numpy (>=1.24.0)
- **Excel處理**：openpyxl (>=3.1.0), xlrd, xlsxwriter
- **視覺化**：matplotlib (>=3.7.0), seaborn (>=0.13.0)

## 📁 專案結構

```
.
├── .venv/                  # 虛擬環境
├── app.py                  # Streamlit主應用文件
├── utils.py                # 核心功能模組
├── requirements.txt        # 依賴包列表
├── README.md               # 項目說明文檔
├── VERSION.md              # 版本更新記錄
├── run.bat                 # Windows運行腳本
├── run.sh                  # macOS/Linux運行腳本
├── test_chart_v1.7.py      # 自動化測試腳本
└── test_data_v1.7.xlsx     # 測試數據文件
```