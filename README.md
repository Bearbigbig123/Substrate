# OSAT SPC System

## Repository Notes

This repo now includes:

- `input/column_aliases.json`: friendly rawdata column-name mapping config
- `input/unified_rawdata_sample.csv`: sample Unified_Vertical rawdata
- `generate_unified_test_data.py`: script for generating Unified_Vertical sample rawdata
- `input/group_weekly_summary_history_template.csv`: group weekly summary CSV template

統計製程管制 (SPC) 分析系統，提供網頁介面與 API 服務，支援 OOB 分析、Tool Matching 分析、和 SPC CPK Dashboard。

## 系統架構

- **前端**：Streamlit 網頁應用程式 (`streamlit_app.py`)
- **後端**：FastAPI REST API 服務 (`main.py`)
- **分析引擎**：原有的 SPC 分析邏輯 (`oob_eng.py`, `tool_matching_widget_osat.py`)

## 快速啟動

### 1. 環境準備

建議使用虛擬環境：

```cmd
python -m venv .venv
.venv\Scripts\activate
```

安裝相依套件：

```cmd
pip install -r requirements.txt
```

### 2. 啟動後端 API 服務

在命令提示字元中執行：

```cmd
uvicorn main:app --host localhost --port 8000 --reload
```

成功啟動後可訪問：
- API 文件：http://localhost:8000/docs
- 健康檢查：http://localhost:8000/health

### 3. 啟動前端網頁介面

開啟另一個命令提示字元，執行：

```cmd
streamlit run streamlit_app.py --server.port 8501
```

或者使用預設埠口：

```cmd
streamlit run streamlit_app.py
```

成功啟動後會自動開啟瀏覽器，或手動訪問：http://localhost:8501

### 4. 登入系統

系統預設帳號密碼：
- 帳號：`admin`
- 密碼：`password`

可透過環境變數自訂：
```cmd
set OOB_USER=your_username
set OOB_PASS=your_password
```

## 功能模組

### Split Chart
將大型 CSV 檔案依據 Chart 資訊分割成個別檔案，支援：
- Type2 垂直分割
- Type3 水平分割

### OOB/SPC 分析
執行統計製程管制分析，產生：
- SPC 控制圖
- 週報控制圖
- 違規規則檢測
- Excel 報告輸出

### Tool Matching
工具匹配分析，包含：
- Mean/Sigma 指標分析
- 統計檢定
- 分組比較

### SPC CPK Dashboard
製程能力分析儀表板：
- 多時間窗口 CPK 計算
- 趨勢分析 (R1/R2 衰退率)
- K 值 (偏移度) 計算
- 互動式圖表與統計摘要

## 檔案結構

```
├── main.py                          # FastAPI 後端服務
├── streamlit_app.py                 # Streamlit 前端應用
├── oob_eng.py                       # SPC 分析核心邏輯
├── tool_matching_widget_osat.py     # Tool Matching 分析
├── spc_cpk_dashboard_osat.py        # CPK Dashboard (PyQt 版本, Streamlit沒用到)
├── requirements.txt                 # Python 相依套件
├── input/                           # 輸入資料夾
│   ├── All_Chart_Information.xlsx   # Chart 設定檔
│   └── raw_charts/                  # 原始資料 CSV 檔案
├── output/                          # 輸出圖表資料夾
└── temp_uploads/                    # 暫存上傳檔案
```

## 分析基準日（base_date）

UI 左側提供「**分析基準日**」日期選擇器，所有分析功能均以此日期為時間基準。

| 功能 | 欄位名稱 | 基準日用途 |
|------|---------|-----------|
| OOB / SPC 分析 | `base_date` | 轉為當天 `23:59:59` 作為**當週結束時間**（`weekly_end_date`），往前推 6 天為當週範圍；基線為往前一年（可自動擴展至兩年） |
| Tool Matching | `base_date` | 作為資料視窗的**結束點**，搭配 `filter_mode: "specified_date"` 進行 1M / 6M 回溯篩選 |
| CPK Dashboard | `end_date` | 作為 CPK 計算視窗的**結束點**，往前計算 L1（1 個月）、L2（2 個月） |

> **注意**：舊版透過 `All_Chart_Information.xlsx` 的 `Time` Sheet 讀取 `execTime` 的邏輯已移除，現在統一由 UI 傳入 `base_date` 決定時間基準。

## 分析資料門檻

### OOB / SPC

| 條件 | 行為 |
|------|------|
| 當週資料 = 0 筆 | 整張圖跳過，不輸出任何結果 |
| 基線（近 1 年）< 10 筆 | 自動擴展至 2 年重新計算 |
| 擴展後仍 < 10 筆 | OOB 規則全部輸出 `NO_HIGHLIGHT`，圖表仍產生 |
| K-shift 當週 < 1 筆 | 回傳 `NO_HIGHLIGHT` |
| K-shift 當週 = 1 筆 | 從基線借點，湊不到 5 筆則回傳 `NO_HIGHLIGHT` |
| Sticking Rate 當週 < 10 筆 | 向基線借 10–20 筆合併後計算 |
| category_LT_shift 當週 < 20 筆 | 從基線補足至 20 筆後計算 |

## API 端點

- `GET /health` - 服務健康檢查
- `GET /` - API 資訊與預設路徑
- `POST /process` - OOB/SPC 分析
- `POST /split` - CSV 檔案分割
- `POST /tool-matching` - Tool Matching 分析
- `POST /spc-cpk` - SPC CPK Dashboard 分析

