import streamlit as st
import requests
import time
import pandas as pd
import base64
import os
import uuid
import io
import xlsxwriter
import hashlib
import plotly.graph_objects as go
from PIL import Image
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# --- 配置 ---
API_BASE_URL = "http://localhost:8000"

# --- 帳號設定（sha256 雜湊；可自行新增帳號）---
# 產生方式：import hashlib; hashlib.sha256(b"your_password").hexdigest()
_ACCOUNTS = {
    "admin": hashlib.sha256(b"admin123").hexdigest(),
    "osat":  hashlib.sha256(b"osat2026").hexdigest(),
}
POLLING_INTERVAL = 2.0  # 輪詢頻率(秒)

# ==========================================
# 本地 CSV 拆分函數 (UI-local Split Logic)
# ==========================================
def _local_sanitize_fn(name: str) -> str:
    for ch in '<>:"/\\|?*\'':
        name = name.replace(ch, "")
    return name.strip()

def _local_read_csv(filepath: str, header_val=None) -> pd.DataFrame:
    for enc in ["utf-8-sig", "utf-8", "big5", "cp950", "latin1", "cp1252"]:
        try:
            return pd.read_csv(filepath, header=header_val, encoding=enc)
        except Exception:
            pass
    raise ValueError(f"無法讀取檔案: {os.path.basename(filepath)}")

def _local_split_type3_horizontal(input_path: str, out_dir: str) -> bool:
    try:
        print(f"[LocalSplit][Type3_Horizontal] 讀取檔案: {os.path.basename(input_path)}")
        df = _local_read_csv(input_path, header_val=None)
        new_columns = []
        for col1, col2 in zip(df.iloc[0], df.iloc[1]):
            if pd.isna(col2):
                new_columns.append(str(col1))
            elif pd.isna(col1):
                new_columns.append(str(col2))
            else:
                new_columns.append(f"{col1}_{col2}")
        df = df.iloc[2:].copy()
        df.columns = new_columns
        chartname_col_name = None
        for col in df.columns:
            if "GroupName" in col and "ChartName" in col:
                chartname_col_name = col
                break
        if chartname_col_name is None:
            raise ValueError("Cannot find combined 'GroupName' and 'ChartName' header column")
        chartname_idx = df.columns.get_loc(chartname_col_name)
        universal_info_columns = df.columns[:chartname_idx + 1].tolist()
        chart_columns = df.columns[chartname_idx + 1:]
        for chart_col in chart_columns:
            temp_df = df[universal_info_columns].copy()
            temp_df["point_val"] = df[chart_col]
            if "_" in chart_col:
                groupname, chartname = chart_col.split("_", 1)
            else:
                groupname, chartname = "", chart_col
            temp_df["GroupName"] = groupname
            temp_df["ChartName"] = chartname
            if "point_time" in temp_df.columns:
                try:
                    temp_df["point_time"] = pd.to_datetime(temp_df["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
                except Exception:
                    pass
            final_cols = ["GroupName", "ChartName", "point_time", "point_val"] + [
                c for c in universal_info_columns
                if c not in ["GroupName", "ChartName", "point_time", "point_val", chartname_col_name]
            ]
            existing = [c for c in final_cols if c in temp_df.columns]
            temp_df = temp_df[existing]
            fn = os.path.join(out_dir, f"{_local_sanitize_fn(str(groupname))}_{_local_sanitize_fn(str(chartname))}.csv")
            if not temp_df.empty:
                temp_df.to_csv(fn, index=False, encoding="utf-8-sig")
        written = len([f for f in os.listdir(out_dir) if f.endswith(".csv")])
        print(f"[LocalSplit][Type3_Horizontal] 完成，寫出 {written} 個 CSV")
        return True
    except Exception as e:
        print(f"[LocalSplit][Type3_Horizontal] 失敗: {e}")
        return False

def _local_split_type2_vertical(input_path: str, out_dir: str) -> bool:
    try:
        print(f"[LocalSplit][Type2_Vertical] 讀取檔案: {os.path.basename(input_path)}")
        df = _local_read_csv(input_path, header_val="infer")
        if not all(c in df.columns for c in ["GroupName", "ChartName", "point_time", "point_val"]):
            raise ValueError("Missing required columns for Type2_Vertical")
        if "point_time" in df.columns:
            try:
                df["point_time"] = pd.to_datetime(df["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
            except Exception:
                pass
        for _, row in df[["GroupName", "ChartName"]].drop_duplicates().iterrows():
            g, c = row["GroupName"], row["ChartName"]
            temp_df = df[(df["GroupName"] == g) & (df["ChartName"] == c)].copy()
            other_cols = [col for col in temp_df.columns if col not in ["GroupName", "ChartName", "point_time", "point_val"]]
            existing = [col for col in ["GroupName", "ChartName", "point_time", "point_val"] + other_cols if col in temp_df.columns]
            temp_df[existing].to_csv(os.path.join(out_dir, f"{_local_sanitize_fn(str(g))}_{_local_sanitize_fn(str(c))}.csv"), index=False, encoding="utf-8-sig")
        written = len([f for f in os.listdir(out_dir) if f.endswith(".csv")])
        print(f"[LocalSplit][Type2_Vertical] 完成，寫出 {written} 個 CSV")
        return True
    except Exception as e:
        print(f"[LocalSplit][Type2_Vertical] 失敗: {e}")
        return False

def _local_split_vendor_vertical(input_path: str, out_dir: str) -> bool:
    try:
        print(f"[LocalSplit][Vendor_Vertical] 讀取檔案: {os.path.basename(input_path)}")
        df = _local_read_csv(input_path, header_val="infer")
        lot_mean_col = "Lot Mean Valid" if "Lot Mean Valid" in df.columns else "Lot Mean"
        col_map = {"Part ID": "GroupName", "Item Name": "ChartName", "Report Time": "point_time", lot_mean_col: "point_val", "Vendor Site": "Matching"}
        required = ["Part ID", "Item Name", "Report Time", lot_mean_col, "Vendor Site"]
        missing = [k for k in required if k not in df.columns]
        if missing:
            raise ValueError(f"Missing vendor columns: {missing}")
        df = df.rename(columns=col_map)
        if "point_time" in df.columns:
            try:
                df["point_time"] = pd.to_datetime(df["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
            except Exception:
                pass
        for _, row in df[["GroupName", "ChartName"]].drop_duplicates().iterrows():
            g, c = row["GroupName"], row["ChartName"]
            temp_df = df[(df["GroupName"] == g) & (df["ChartName"] == c)].copy()
            other_cols = [col for col in temp_df.columns if col not in ["GroupName", "ChartName", "point_time", "point_val"]]
            existing = [col for col in ["GroupName", "ChartName", "point_time", "point_val"] + other_cols if col in temp_df.columns]
            temp_df[existing].to_csv(os.path.join(out_dir, f"{_local_sanitize_fn(str(g))}_{_local_sanitize_fn(str(c))}.csv"), index=False, encoding="utf-8-sig")
        written = len([f for f in os.listdir(out_dir) if f.endswith(".csv")])
        print(f"[LocalSplit][Vendor_Vertical] 完成，寫出 {written} 個 CSV")
        return True
    except Exception as e:
        print(f"[LocalSplit][Vendor_Vertical] 失敗: {e}")
        return False

def _local_split_test_horizontal(input_path: str, out_dir: str) -> bool:
    try:
        print(f"[LocalSplit][Test_Horizontal] 讀取檔案: {os.path.basename(input_path)}")
        df = _local_read_csv(input_path, header_val="infer")
        col_map = {"Part ID": "GroupName", "FT Test End Time": "point_time", "Test Site": "Matching"}
        missing = [k for k in col_map if k not in df.columns]
        if missing:
            raise ValueError(f"Missing test columns: {missing}")
        df = df.rename(columns=col_map)
        if "point_time" in df.columns:
            try:
                df["point_time"] = pd.to_datetime(df["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
            except Exception:
                pass
        matching_idx = df.columns.get_loc("Matching")
        id_cols = df.columns[:matching_idx + 1].tolist()
        value_cols = df.columns[matching_idx + 1:].tolist()
        if not value_cols:
            raise ValueError("No test item columns found after 'Matching' column")
        df_melted = df.melt(id_vars=id_cols, value_vars=value_cols, var_name="ChartName", value_name="point_val").dropna(subset=["point_val"])
        standard_cols = ["GroupName", "ChartName", "point_time", "point_val", "Matching"]
        for _, row in df_melted[["GroupName", "ChartName"]].drop_duplicates().iterrows():
            g, c = row["GroupName"], row["ChartName"]
            temp_df = df_melted[(df_melted["GroupName"] == g) & (df_melted["ChartName"] == c)].copy()
            existing = [col for col in standard_cols if col in temp_df.columns]
            temp_df[existing].to_csv(os.path.join(out_dir, f"{_local_sanitize_fn(str(g))}_{_local_sanitize_fn(str(c))}.csv"), index=False, encoding="utf-8-sig")
        written = len([f for f in os.listdir(out_dir) if f.endswith(".csv")])
        print(f"[LocalSplit][Test_Horizontal] 完成，寫出 {written} 個 CSV")
        return True
    except Exception as e:
        print(f"[LocalSplit][Test_Horizontal] 失敗: {e}")
        return False

def _local_split_file(input_path: str, mode: str) -> str:
    """在本地執行 CSV 拆分，回傳 split_data 絕對路徑；失敗時拋出 Exception。"""
    split_dir = os.path.abspath(os.path.join("temp_uploads", str(uuid.uuid4()), "split_data"))
    os.makedirs(split_dir, exist_ok=True)
    print(f"[LocalSplit] 開始拆分 | mode={mode} | input={input_path} | output_dir={split_dir}")
    dispatch = {
        "Type3_Horizontal": _local_split_type3_horizontal,
        "Type2_Vertical":   _local_split_type2_vertical,
        "Vendor_Vertical":  _local_split_vendor_vertical,
        "Test_Horizontal":  _local_split_test_horizontal,
    }
    fn = dispatch.get(mode)
    if fn is None:
        raise ValueError(f"Unknown split mode: {mode}")
    ok = fn(input_path, split_dir)
    if not ok:
        raise RuntimeError(f"Split failed for mode: {mode}")
    output_files = os.listdir(split_dir)
    print(f"[LocalSplit] 拆分完成 | mode={mode} | 產出 {len(output_files)} 個檔案 | dir={split_dir}")
    return split_dir

st.set_page_config(page_title="OSAT SPC System", layout="wide")

# --- CSS：優化進度條顏色、隱藏預設 Header、調整版面 ---
st.markdown("""
    <style>
    /* 1. 這裡保留你的進度條漸層色 */
    .stProgress > div > div > div > div { 
        background-image: linear-gradient(to right, #344CB7, #577BC1); 
    }
    
    /* 2. 💡 這裡移除了隱藏 Header 的代碼，小人圖示會回來 */
    
    /* 3. 調整頁面間距，既然 Header 回來了，頂部 padding 縮小一點才不會留白太多 */
    .block-container { padding-top: 3rem; }
    
    /* 4. 讓所有 st.subheader (h3) 小一號 */
    h3 {
        font-size: 1.3rem !important;
        font-weight: 600 !important;
        color: #262730;
        margin-bottom: 0.5rem !important;
        padding-top: 0.5rem !important;
    }
    
    /* 5. 增加 st.popover (Settings 彈出視窗) 的高度 */
    div[data-testid="stPopoverBody"] {
        max-height: 95vh !important;
        height: auto !important;
        overflow-y: auto !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 狀態管理 ---
if 'task_id' not in st.session_state: st.session_state.task_id = None
if 'last_task_id' not in st.session_state: st.session_state.last_task_id = None
if 'status' not in st.session_state: st.session_state.status = "idle"
if 'progress' not in st.session_state: st.session_state.progress = 0
if 'results' not in st.session_state: st.session_state.results = None
if 'current_mode' not in st.session_state: st.session_state.current_mode = None
if 'full_excel_data' not in st.session_state: st.session_state.full_excel_data = None
if 'trigger_analysis' not in st.session_state: st.session_state.trigger_analysis = False
if 'pending_endpoint' not in st.session_state: st.session_state.pending_endpoint = None
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'login_user' not in st.session_state: st.session_state.login_user = ""
if 'poll_count' not in st.session_state: st.session_state.poll_count = 0
# 跨功能切換時保留上傳檔案路徑（Streamlit rerun 會清空 file_uploader）
if 'saved_excel_path' not in st.session_state: st.session_state.saved_excel_path = None
if 'saved_raw_dir' not in st.session_state: st.session_state.saved_raw_dir = None
if 'saved_split_raw_dir' not in st.session_state: st.session_state.saved_split_raw_dir = None
if 'saved_split_id' not in st.session_state: st.session_state.saved_split_id = None
if 'saved_split_info' not in st.session_state: st.session_state.saved_split_info = None

# ==========================================
# LOGIN GATE
# ==========================================
if not st.session_state.logged_in:
    st.markdown("""
            <style>
            /* 1. 隱藏原本突兀的白色卡片，改用透明毛玻璃質感 */
            .stApp {
                background-color: #0E1117;
            }

            /* 2. 登入容器：徹底拿掉白色背景 */
            .login-box {
                max-width: 420px;
                margin: 12vh auto 0 auto;
                padding: 2rem;
                text-align: center;
                background: rgba(255, 255, 255, 0.03); /* 極微弱的白色透明 */
                border: 1px solid rgba(255, 255, 255, 0.1); /* 纖細邊框 */
                border-radius: 15px 15px 0 0;
                backdrop-filter: blur(10px); /* 毛玻璃效果 */
            }

            .login-title {
                color: #FFFFFF;
                font-size: 2rem;
                font-weight: 800;
                letter-spacing: 1px;
                text-shadow: 0 0 10px rgba(52, 76, 183, 0.5); /* 標題微光 */
                margin-bottom: 0.5rem;
            }

            .login-sub {
                color: #888;
                font-size: 0.9rem;
                margin-bottom: 0;
            }

            /* 3. 處理下方 Form：無縫接軌上方的透明框 */
            div[data-testid="stForm"] {
                max-width: 420px;
                margin: -1px auto 0 auto; /* 讓上下框完美接合 */
                background: rgba(255, 255, 255, 0.05) !important; 
                border: 1px solid rgba(255, 255, 255, 0.1) !important;
                border-top: none !important; /* 拿掉銜接處的線 */
                border-radius: 0 0 15px 15px !important;
                padding: 2rem !important;
                box-shadow: 0 20px 40px rgba(0,0,0,0.4);
            }

            /* 4. 讓 Input 框更帥：深黑底+淡藍邊 */
            input {
                background-color: #0E1117 !important;
                color: white !important;
                border: 1px solid #344CB7 !important;
            }

            /* 5. 登入按鈕：改用藍色漸層光，更符合 QC 戰情室的科技感 */
            button[kind="primaryFormSubmit"] {
                width: 100% !important;
                background: linear-gradient(90deg, #344CB7, #577BC1) !important;
                border: none !important;
                color: white !important;
                font-weight: 700 !important;
                height: 3rem !important;
                transition: 0.3s !important;
                box-shadow: 0 4px 15px rgba(52, 76, 183, 0.3) !important;
            }
            
            button[kind="primaryFormSubmit"]:hover {
                box-shadow: 0 0 20px rgba(52, 76, 183, 0.6) !important;
                transform: translateY(-2px);
            }
            </style>

            <div class='login-box'>
                <div class='login-title'>OSAT SPC SYSTEM</div>
                <div class='login-sub'>PRECISION QUALITY CONTROL</div>
            </div>
        """, unsafe_allow_html=True)

    _, login_col, _ = st.columns([1, 1.2, 1])
    with login_col:
        username = st.text_input("帳號", placeholder="Username", label_visibility="collapsed", key="_lu")
        password = st.text_input("密碼", placeholder="Password", type="password", label_visibility="collapsed", key="_lp")
        if st.button("登入", type="primary", use_container_width=True):
            hashed = hashlib.sha256(password.encode()).hexdigest()
            if username in _ACCOUNTS and _ACCOUNTS[username] == hashed:
                st.session_state.logged_in = True
                st.session_state.login_user = username
                st.rerun()
            else:
                st.error("帳號或密碼錯誤")
    st.stop()
if 'pending_payload' not in st.session_state: st.session_state.pending_payload = None
if 'pending_mode' not in st.session_state: st.session_state.pending_mode = None
if 'auto_split_info' not in st.session_state: st.session_state.auto_split_info = None

# ==========================================
# 0. 輔助函數：產生包含圖片的完整 Excel
# ==========================================
def generate_full_excel_with_images(data_list, mode):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Analysis_Results')

    df = pd.DataFrame(data_list)

    # 處理 Metrics 展開 (用於 CPK 模式)
    if 'metrics' in df.columns:
        metrics_df = pd.json_normalize(df['metrics'])
        df = pd.concat([df.drop(columns=['metrics']), metrics_df], axis=1)

    # 定義 Excel 樣式
    header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11, 'bold': True, 'bg_color': '#344CB7', 'font_color': 'white'})
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10})

    # 決定要排除的欄位 (圖片路徑不需要顯示為文字)
    exclude_cols = ['chart_path', 'weekly_chart_path', 'by_tool_color_path', 'by_tool_group_path', 'qq_plot_path', 'chart_image', 'spc_chart_path', 'boxplot_chart_path', 'timeline_chart_path', 'chart_data']
    data_cols = [c for c in df.columns if c not in exclude_cols]

    is_oob_mode = mode == "OOB/SPC"
    is_cpk_mode = mode == "CPK Dashboard"
    is_tm_mode  = mode == "Tool Matching"

    # CPK 違規計算與格式
    violation_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10, 'bg_color': '#FFD0D0'})
    if is_cpk_mode and ('r1' in df.columns or 'r2' in df.columns):
        def _excel_cpk_viol(row):
            try: r1f = float(row.get('r1')) if row.get('r1') is not None and pd.notna(row.get('r1')) else None
            except: r1f = None
            try: r2f = float(row.get('r2')) if row.get('r2') is not None and pd.notna(row.get('r2')) else None
            except: r2f = None
            h1 = r1f is not None and r1f >= 25
            h2 = r2f is not None and r2f >= 20
            if h1 and h2: return "H(R1+R2)"
            if h1: return "H(R1)"
            if h2: return "H(R2)"
            return ""
        df['cpk_violation'] = df.apply(_excel_cpk_viol, axis=1)

    # OOB: 5 張圖表欄位
    img_fields = [
        ('chart_path', 'Total SPC'),
        ('weekly_chart_path', 'Weekly SPC'),
        ('by_tool_color_path', 'By Tool (Color)'),
        ('by_tool_group_path', 'By Tool (Group)'),
        ('qq_plot_path', 'Q-Q Plot')
    ]
    # CPK: 1 張圖表欄位 (base64)
    cpk_img_fields = [('chart_image', 'SPC Chart')]
    # TM: 3 張圖表欄位 (file path)
    tm_img_fields = [
        ('spc_chart_path', 'SPC Chart'),
        ('boxplot_chart_path', 'Boxplot'),
        ('timeline_chart_path', 'Timeline'),
    ]

    if is_oob_mode:
        num_img_cols = len(img_fields)
    elif is_cpk_mode:
        num_img_cols = len(cpk_img_fields)
    elif is_tm_mode:
        num_img_cols = len(tm_img_fields)
    else:
        num_img_cols = 0
    start_col = num_img_cols

    # 寫入圖片的表頭
    if is_oob_mode:
        for i, (key, title) in enumerate(img_fields):
            worksheet.write(0, i, title, header_format)
    elif is_cpk_mode:
        for i, (key, title) in enumerate(cpk_img_fields):
            worksheet.write(0, i, title, header_format)
    elif is_tm_mode:
        for i, (key, title) in enumerate(tm_img_fields):
            worksheet.write(0, i, title, header_format)

    # 寫入文字數據的表頭
    col_widths = {}
    for i, col_name in enumerate(data_cols):
        worksheet.write(0, start_col + i, col_name.replace('_', ' ').title(), header_format)
        col_widths[start_col + i] = max(len(str(col_name)), 12)

    scale_factor = 0.28 if is_oob_mode else (0.40 if is_tm_mode else 0.55)
    max_image_height = 0
    max_image_width = 0

    # CPK 模式：建立暫存目錄用於 base64 圖片解碼
    import tempfile
    _cpk_tmp_dir = tempfile.mkdtemp() if is_cpk_mode else None
    _tm_scale_factor = 0.40

    # 迭代資料列
    for row_idx, row in df.iterrows():
        excel_row = row_idx + 1

        # 插入圖片 (OOB: 5 張路徑圖片；CPK: 1 張 base64 圖片)
        if is_oob_mode:
            for col_offset, (key, _) in enumerate(img_fields):
                img_path = row.get(key)
                if pd.notna(img_path) and isinstance(img_path, str) and os.path.exists(img_path):
                    try:
                        options = {'x_scale': scale_factor, 'y_scale': scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, img_path, options)
                        with Image.open(img_path) as img:
                            scaled_h = img.height * scale_factor
                            scaled_w = img.width * scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)
        elif is_cpk_mode:
            for col_offset, (key, _) in enumerate(cpk_img_fields):
                img_b64 = row.get(key)
                if pd.notna(img_b64) and isinstance(img_b64, str) and img_b64:
                    try:
                        img_bytes = base64.b64decode(img_b64)
                        tmp_path = os.path.join(_cpk_tmp_dir, f"cpk_{row_idx}_{col_offset}.png")
                        with open(tmp_path, 'wb') as f:
                            f.write(img_bytes)
                        options = {'x_scale': scale_factor, 'y_scale': scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, tmp_path, options)
                        with Image.open(tmp_path) as img:
                            scaled_h = img.height * scale_factor
                            scaled_w = img.width * scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)
        elif is_tm_mode:
            for col_offset, (key, _) in enumerate(tm_img_fields):
                img_path = row.get(key)
                if pd.notna(img_path) and isinstance(img_path, str) and os.path.exists(img_path):
                    try:
                        options = {'x_scale': _tm_scale_factor, 'y_scale': _tm_scale_factor, 'x_offset': 5, 'y_offset': 5, 'object_position': 1}
                        worksheet.insert_image(excel_row, col_offset, img_path, options)
                        with Image.open(img_path) as img:
                            scaled_h = img.height * _tm_scale_factor
                            scaled_w = img.width * _tm_scale_factor
                            max_image_height = max(max_image_height, scaled_h)
                            max_image_width = max(max_image_width, scaled_w)
                    except Exception:
                        worksheet.write(excel_row, col_offset, "Image Error", cell_format)

        # 寫入文字數據
        _cpk_v = row.get('cpk_violation', '') if 'cpk_violation' in df.columns else ''
        _row_fmt = violation_format if (is_cpk_mode and str(_cpk_v) not in ['', 'nan', 'NaN', 'None']) else cell_format
        for i, col_name in enumerate(data_cols):
            val = row.get(col_name)
            if pd.isna(val):
                val = ""
            elif isinstance(val, (list, dict)):
                val = str(val)
            worksheet.write(excel_row, start_col + i, val, _row_fmt)
            col_widths[start_col + i] = max(col_widths.get(start_col + i, 12), len(str(val)) + 2)

    # 設定寬度與高度
    if (is_oob_mode or is_cpk_mode or is_tm_mode) and max_image_height > 0:
        if is_oob_mode:
            col_w_cap, row_h_cap = 55, 150
        elif is_tm_mode:
            col_w_cap, row_h_cap = 75, 200
        else:
            col_w_cap, row_h_cap = 80, 200
        for i in range(num_img_cols):
            worksheet.set_column(i, i, min((max_image_width / 7) + 2, col_w_cap))
        for row_idx in range(1, len(df) + 1):
            worksheet.set_row(row_idx, min((max_image_height * 0.75) + 10, row_h_cap))
    else:
        for row_idx in range(1, len(df) + 1):
            worksheet.set_row(row_idx, 20)

    for col_idx, width in col_widths.items():
        worksheet.set_column(col_idx, col_idx, min(width, 40))

    # 凍結首列（不凍結欄位）
    worksheet.freeze_panes(1, 0)

    workbook.close()
    if _cpk_tmp_dir and os.path.exists(_cpk_tmp_dir):
        import shutil
        shutil.rmtree(_cpk_tmp_dir, ignore_errors=True)
    return output.getvalue()


# ==========================================
# 1. 頂部導覽橫條 (Header)
# ==========================================
col1, col2, col3, col4 = st.columns([1.5, 3.8, 0.8, 0.8], gap="medium")

with col1:
    with st.popover("⚙️ Settings & Run", use_container_width=True):
        st.markdown("##### 分析設定")
        mode = st.radio("選擇功能", ["OOB/SPC", "Tool Matching", "CPK Dashboard"])
        base_date = st.date_input("分析基準日", value=datetime.now())
        
       
        # --- 檔案上傳區塊 (水平排列) ---
        st.markdown("###### 📁 上傳自訂檔案 (若不傳則使用預設)")
        up_col_left, up_col_right = st.columns(2)
        
        with up_col_left:
            excel_file = st.file_uploader("1️⃣ All Charts (Excel)", type=["xlsx"], help="上傳定義控制線的 Excel 檔")
            
        with up_col_right:
            csv_files = st.file_uploader("2️⃣ Raw Data (CSV, 多選)", type=["csv"], accept_multiple_files=True, help="上傳產線原始資料 CSV 檔")
            
        st.divider()
        
        if st.button("🚀 Start Analysis", type="primary", use_container_width=True):
            st.session_state.current_mode = mode
            st.session_state.results = None
            st.session_state.status = "idle"
            st.session_state.progress = 0
            
            current_excel_path = None
            current_raw_dir = None

            # --- 1. 處理「新上傳」的檔案 ---
            if excel_file or csv_files:
                upload_session_id = str(uuid.uuid4())
                base_upload_dir = os.path.abspath(os.path.join("temp_uploads", "ui_uploads", upload_session_id))
                os.makedirs(base_upload_dir, exist_ok=True)
                
                if excel_file:
                    excel_file.seek(0)
                    current_excel_path = os.path.join(base_upload_dir, excel_file.name)
                    with open(current_excel_path, "wb") as f:
                        f.write(excel_file.read())
                    st.session_state.saved_excel_path = current_excel_path
                
                if csv_files:
                    current_raw_dir = os.path.join(base_upload_dir, "raw_charts")
                    os.makedirs(current_raw_dir, exist_ok=True)
                    for csv in csv_files:
                        csv.seek(0)
                        csv_path = os.path.join(current_raw_dir, csv.name)
                        with open(csv_path, "wb") as f:
                            f.write(csv.read())
                    st.session_state.saved_raw_dir = current_raw_dir
                    
                    # 上傳新 CSV 時，清除舊的 Split 記憶
                    st.session_state.saved_split_raw_dir = None
                    st.session_state.saved_split_id = None
                    st.session_state.saved_split_info = None

            # --- 2. 自動偵測與拆分 (僅在新上傳 CSV 時觸發) ---
            if csv_files and len(csv_files) == 1 and current_raw_dir:
                first_saved = os.path.join(current_raw_dir, csv_files[0].name)
                try:
                    peek = pd.read_csv(first_saved, nrows=0)
                    detected_cols = set(peek.columns)
                    detected_split_mode = None
                    
                    if {"Part ID", "Item Name", "Report Time", "Vendor Site"}.issubset(detected_cols) and ("Lot Mean" in detected_cols or "Lot Mean Valid" in detected_cols):
                        detected_split_mode = "Vendor_Vertical"
                    elif {"Part ID", "FT Test End Time", "Test Site"}.issubset(detected_cols):
                        detected_split_mode = "Test_Horizontal"
                    elif {"GroupName", "ChartName", "point_time", "point_val"}.issubset(detected_cols):
                        detected_split_mode = "Type2_Vertical"
                    else:
                        peek_no_header = pd.read_csv(first_saved, nrows=3, header=None)
                        flat_vals = peek_no_header.iloc[0:2].fillna("").astype(str).values.flatten().tolist()
                        if any("GroupName" in val for val in flat_vals) and any("ChartName" in val for val in flat_vals):
                            detected_split_mode = "Type3_Horizontal"

                    if detected_split_mode:
                        try:
                            with st.spinner("正在自動準備資料夾..."):
                                split_data_dir = _local_split_file(first_saved, detected_split_mode)
                            st.session_state.saved_split_raw_dir = split_data_dir
                            st.session_state.saved_split_id = None
                            st.session_state.saved_split_info = f"🔀 自動偵測到 **{detected_split_mode}** 格式，已完成拆分"
                        except Exception as split_err:
                            st.error(f"⚠️ 自動拆分失敗：{str(split_err)}")
                except Exception as e:
                    st.error(f"⚠️ 讀取或拆分檔案時發生錯誤：{str(e)}")

            # --- 3. 如果本次沒有上傳，強制沿用 Session 中的舊檔案 ---
            if not current_excel_path: current_excel_path = st.session_state.get("saved_excel_path")
            if not current_raw_dir: current_raw_dir = st.session_state.get("saved_raw_dir")
            
            auto_split_raw_dir = st.session_state.get("saved_split_raw_dir")
            st.session_state.auto_split_info = st.session_state.get("saved_split_info")

            # --- 4. 防呆機制：如果真的沒檔案可送，擋住並警告 ---
            if not current_excel_path and not current_raw_dir and not auto_split_raw_dir:
                st.error("⚠️ 系統找不到分析資料，請重新上傳檔案！")
                st.stop()

            # --- 5. 組裝 API Payload ---
            payload = {}
            if mode == "OOB/SPC":
                endpoint = "/process"
                payload["base_date"] = base_date.strftime("%Y-%m-%d")
                if current_excel_path: payload["filepath"] = current_excel_path
                if auto_split_raw_dir: payload["raw_data_directory"] = auto_split_raw_dir
                elif current_raw_dir: payload["raw_data_directory"] = current_raw_dir
                
            elif mode == "Tool Matching":
                endpoint = "/tool-matching"
                payload = {"base_date": base_date.strftime("%Y-%m-%d"), "filter_mode": "specified_date"}
                if current_excel_path: payload["chart_excel_path"] = current_excel_path
                if auto_split_raw_dir: payload["raw_data_directory"] = auto_split_raw_dir
                elif current_raw_dir: payload["raw_data_directory"] = current_raw_dir
                
            else: # CPK Dashboard
                endpoint = "/spc-cpk"
                payload = {"end_date": base_date.strftime("%Y-%m-%d")}
                if current_excel_path: payload["chart_excel_path"] = current_excel_path
                if auto_split_raw_dir: payload["raw_data_directory"] = auto_split_raw_dir
                elif current_raw_dir: payload["raw_data_directory"] = current_raw_dir

            # 先儲存參數並關閉彈窗，重繪後再送出 API
            st.session_state.pending_mode = mode
            st.session_state.pending_endpoint = endpoint
            st.session_state.pending_payload = payload
            st.session_state.trigger_analysis = True
            st.rerun()

with col2:
    if st.session_state.status == "processing":
        st.progress(st.session_state.progress / 100.0)
        st.caption(f"⏳ {st.session_state.current_mode} 執行中... ({st.session_state.progress}%)")
    elif st.session_state.status == "completed":
        st.progress(1.0)
        st.caption("✅ 任務已完成，請點擊左下方表格任意一列檢視圖表")
    elif st.session_state.status == "failed":
        st.progress(0)
        err_msg = ""
        if st.session_state.task_id and st.session_state.task_id in st.session_state.get('_last_errors', {}):
            err_msg = st.session_state['_last_errors'][st.session_state.task_id]
        st.caption(f"❌ 任務失敗{f': {err_msg}' if err_msg else ''}")
    if st.session_state.auto_split_info:
        st.caption(st.session_state.auto_split_info)

with col3:
    if st.button("🔄 Reset", use_container_width=True):
        st.session_state.task_id = None
        st.session_state.last_task_id = None
        st.session_state.status = "idle"
        st.session_state.results = None
        st.session_state.progress = 0
        st.session_state.full_excel_data = None
        st.session_state.auto_split_info = None
        # 保留 saved_* 檔案記憶，不清除，讓切換功能或重跑時不需重新上傳
        # st.session_state.saved_excel_path = None
        # st.session_state.saved_raw_dir = None
        # st.session_state.saved_split_raw_dir = None
        # st.session_state.saved_split_id = None
        st.rerun()

with col4:
    st.caption(f"👤 {st.session_state.login_user}")

# ==========================================
# 1.5 延遲執行分析 (彈窗關閉後才送出 API)
# ==========================================
if st.session_state.trigger_analysis:
    st.session_state.trigger_analysis = False
    st.session_state.current_mode = st.session_state.pending_mode
    st.session_state.results = None
    st.session_state.poll_count = 0
    try:
        resp = requests.post(f"{API_BASE_URL}{st.session_state.pending_endpoint}", json=st.session_state.pending_payload)
        if resp.status_code == 200:
            st.session_state.task_id = resp.json().get("task_id")
            st.session_state.status = "processing"
            st.session_state.progress = 0
            st.rerun()
        else:
            st.error(f"啟動失敗: {resp.text}")
    except Exception as e:
        st.error(f"API 連線失敗: {e}")

# ==========================================
# 2. 輪詢進度 (Polling Logic)
# ==========================================
_MAX_POLL_COUNT = 1800  # 1800 × 2秒 ≈ 60分鐘上限

if st.session_state.status == "processing" and st.session_state.task_id:
    # 前端超時保護：超過 _MAX_POLL_COUNT 次輪詢後自動中止
    if st.session_state.get("poll_count", 0) > _MAX_POLL_COUNT:
        st.session_state.status = "failed"
        st.session_state.poll_count = 0
        if '_last_errors' not in st.session_state:
            st.session_state['_last_errors'] = {}
        st.session_state['_last_errors'][st.session_state.task_id] = "前端輪詢超時（>60分鐘），請確認後台是否正常"
        st.error("分析超時，請確認後台狀態後重新送出")
        st.rerun()
    try:
        resp = requests.get(f"{API_BASE_URL}/process/status/{st.session_state.task_id}").json()
        st.session_state.progress = resp.get("progress", 0)
        status = resp.get("status")
        if status == "completed":
            # 狀態字典只有輕量資訊；完整結果從獨立端點讀取（不走 Manager.dict）
            result_resp = requests.get(f"{API_BASE_URL}/process/result/{st.session_state.task_id}")
            st.session_state.results = result_resp.json()
            st.session_state.poll_count = 0
            st.session_state.status = "completed"
            st.rerun()
        elif status == "failed":
            st.session_state.status = "failed"
            st.session_state.poll_count = 0
            err = resp.get('error', '')
            if err:
                if '_last_errors' not in st.session_state:
                    st.session_state['_last_errors'] = {}
                st.session_state['_last_errors'][st.session_state.task_id] = err
                st.error(f"後端出錯: {err}")
            st.rerun()
        else:
            st.session_state.poll_count = st.session_state.get("poll_count", 0) + 1
            time.sleep(POLLING_INTERVAL)
            st.rerun()
    except Exception as e:
        st.error(f"輪詢失敗: {e}")

st.divider()

# ==========================================
# 3. 報表與圖表顯示區 (Master-Detail with AgGrid)
# ==========================================
if st.session_state.results:
    res = st.session_state.results
    data_list = []
    if st.session_state.current_mode in ["OOB/SPC", "Tool Matching"]:
        data_list = res.get("results", [])
    elif st.session_state.current_mode == "CPK Dashboard":
        data_list = res.get("charts", [])
        
    if data_list:
        df_summary = pd.DataFrame(data_list)
        
        if 'metrics' in df_summary.columns:
            metrics_df = pd.json_normalize(df_summary['metrics'])
            df_processed = pd.concat([df_summary.drop(columns=['metrics']), metrics_df], axis=1)
        else:
            df_processed = df_summary.copy()

        # CPK 模式：計算 R1/R2 違規欄位
        if st.session_state.current_mode == "CPK Dashboard":
            def _cpk_viol(row):
                try:
                    r1v = row.get('r1')
                    r1f = float(r1v) if r1v is not None and pd.notna(r1v) and str(r1v) not in ['nan', 'NaN', '', '-'] else None
                except Exception:
                    r1f = None
                try:
                    r2v = row.get('r2')
                    r2f = float(r2v) if r2v is not None and pd.notna(r2v) and str(r2v) not in ['nan', 'NaN', '', '-'] else None
                except Exception:
                    r2f = None
                h1 = r1f is not None and r1f >= 25
                h2 = r2f is not None and r2f >= 20
                if h1 and h2: return "H(R1+R2)"
                if h1: return "H(R1)"
                if h2: return "H(R2)"
                return ""
            if 'r1' in df_processed.columns or 'r2' in df_processed.columns:
                df_processed['cpk_violation'] = df_processed.apply(_cpk_viol, axis=1)

        c_top_left, c_top_right = st.columns([1.6, 2.4], gap="small")
        
        with c_top_left:
            # === 💡 在標題旁邊加入「下載完整 Excel 報告」按鈕 ===
            title_col, btn_col = st.columns([1, 1])
            with title_col:
                st.markdown("##### Summary Table")
            
            with btn_col:
                # 如果是新的任務，則在背景產生一份 Excel 暫存在 Session 裡
                if st.session_state.last_task_id != st.session_state.task_id:
                    st.session_state.full_excel_data = generate_full_excel_with_images(data_list, st.session_state.current_mode)
                    st.session_state.last_task_id = st.session_state.task_id
                
                st.download_button(
                    label="📥 下載完整 Excel 報告 (含圖片)",
                    data=st.session_state.full_excel_data,
                    file_name=f"SPC_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # --- 任務摘要統計列 ---
            _mode_now = st.session_state.get("current_mode", "")
            if _mode_now == "OOB/SPC":
                _total = len(data_list)
                _oos_n = sum(1 for r in data_list if (r.get("oos_cnt") or 0) > 0)
                _ooc_n = sum(1 for r in data_list if (r.get("ooc_cnt") or 0) > 0)
                _oob_n = sum(1 for r in data_list if r.get("OOB_Rule") not in [None, "", "N/A", "-", "nan", "NaN"])
                _we_n  = sum(1 for r in data_list if r.get("WE_Rule")  not in [None, "", "N/A", "-", "nan", "NaN"])
                st.dataframe(pd.DataFrame({
                    "Total": [_total], "OOS>0": [_oos_n], "OOC>0": [_ooc_n],
                    "OOB Violations": [_oob_n], "WE Violations": [_we_n],
                }), hide_index=True, use_container_width=True)
            elif _mode_now == "Tool Matching":
                _total = len(data_list)
                _abn_n = sum(1 for r in data_list if r.get("need_matching") or r.get("abnormal_type") not in [None, "", "nan", "NaN"])
                st.dataframe(pd.DataFrame({"Total": [_total], "Abnormal": [_abn_n]}), hide_index=True, use_container_width=True)
            elif _mode_now == "CPK Dashboard":
                def _sf(v):
                    try: return float(v) if v not in [None, '', 'nan', 'NaN'] else 0.0
                    except: return 0.0
                _total = len(data_list)
                _r1_n = sum(1 for r in data_list if _sf((r.get('metrics') or {}).get('r1') or r.get('r1')) >= 25)
                _r2_n = sum(1 for r in data_list if _sf((r.get('metrics') or {}).get('r2') or r.get('r2')) >= 20)
                _both = sum(1 for r in data_list if _sf((r.get('metrics') or {}).get('r1') or r.get('r1')) >= 25 and _sf((r.get('metrics') or {}).get('r2') or r.get('r2')) >= 20)
                st.dataframe(pd.DataFrame({"Total": [_total], "R1 Viol.(≥25%)": [_r1_n], "R2 Viol.(≥20%)": [_r2_n], "Both": [_both]}), hide_index=True, use_container_width=True)

            # --- 💡 資料清洗與自動置頂邏輯 ---
            keep_list = ['gname', 'cname', 'group', 'group_name', 'chart_name', 'Characteristics', 'characteristics', 'WE_Rule', 'OOB_Rule', 'ooc_cnt', 'oos_cnt', 'abnormal_type', 'cpk', 'cpk_l1', 'cpk_l2', 'r1', 'r2', 'cpk_violation', 'k_value', 'mean_index', 'sigma_index', 'data_cnt']
            existing_cols = [c for c in keep_list if c in df_processed.columns]
            display_df = df_processed[existing_cols].copy()
            
            normal_keywords = ["", "NA", "N/A", "NONE", "NAN", "NORMAL", "PASS", "NO_HIGHLIGHT", "FALSE", "<NA>", "0", "OK", "-"]
            
            for rule_col in ['WE_Rule', 'OOB_Rule']:
                if rule_col in display_df.columns:
                    clean_str = display_df[rule_col].astype(str).str.strip().str.upper()
                    display_df.loc[clean_str.isin(normal_keywords), rule_col] = ""

            # --- 無資料 / 點數不足 標示（直接寫入 OOB_Rule，不新增欄位）---
            if 'OOB_Rule' in display_df.columns:
                _no_data_mask = pd.Series(False, index=df_processed.index)
                _insuf_mask   = pd.Series(False, index=df_processed.index)
                if 'no_data' in df_processed.columns:
                    _no_data_mask = df_processed['no_data'].fillna(False).astype(bool)
                if 'baseline_insufficient' in df_processed.columns:
                    _insuf_mask = df_processed['baseline_insufficient'].fillna(False).astype(bool)
                display_df.loc[_no_data_mask, 'OOB_Rule'] = "⚠ 無資料"
                display_df.loc[_insuf_mask & ~_no_data_mask, 'OOB_Rule'] = "⚠ 點數不足"

            has_we = (display_df['WE_Rule'] != "") if 'WE_Rule' in display_df.columns else pd.Series(False, index=display_df.index)
            has_oob = (display_df['OOB_Rule'] != "") if 'OOB_Rule' in display_df.columns else pd.Series(False, index=display_df.index)
            has_tm = (display_df['abnormal_type'] != "") if 'abnormal_type' in display_df.columns else pd.Series(False, index=display_df.index)
            has_cpk_viol = (display_df['cpk_violation'] != "") if 'cpk_violation' in display_df.columns else pd.Series(False, index=display_df.index)
            display_df['has_issue'] = has_we | has_oob | has_tm | has_cpk_viol
            
            _sort_candidates = ['has_issue', 'group_name', 'gname', 'chart_name', 'cname']
            _sort_cols = [c for c in _sort_candidates if c in display_df.columns]
            _ascending = [False if c == 'has_issue' else True for c in _sort_cols]
            display_df = display_df.sort_values(by=_sort_cols, ascending=_ascending)          
            display_df = display_df.drop(columns=['has_issue'])
            
            if 'WE_Rule' in display_df.columns:
                display_df.loc[display_df['WE_Rule'] == "", 'WE_Rule'] = "-"
            if 'OOB_Rule' in display_df.columns:
                display_df.loc[display_df['OOB_Rule'] == "", 'OOB_Rule'] = "-"
            
            _na_vals = ["N/A", "None", "nan", "NaN", "none", "<NA>"]
            for _col in display_df.columns:
                if _col not in ['WE_Rule', 'OOB_Rule']: 
                    if display_df[_col].dtype == object:
                        display_df[_col] = display_df[_col].fillna("-").replace(_na_vals, "-")
                    else:
                        display_df[_col] = display_df[_col].where(display_df[_col].notna(), other="-")

            # --- 建立 AgGrid ---
            gb = GridOptionsBuilder.from_dataframe(display_df)
            gb.configure_selection(selection_mode="single", use_checkbox=False)
            gb.configure_default_column(resizable=True)
            
            col_settings = {
                "gname": {"header_name": "Group", "width": 90},
                "cname": {"header_name": "Chart", "width": 140},
                "group": {"header_name": "Group", "width": 90},
                "group_name": {"header_name": "Group", "width": 90},
                "chart_name": {"header_name": "Chart", "width": 140},
                "Characteristics": {"header_name": "Char.", "width": 75},
                "characteristics": {"header_name": "Char.", "width": 75},
                "WE_Rule": {"header_name": "WE Rule", "width": 200},
                "OOB_Rule": {"header_name": "OOB Rule", "width": 200},
                "abnormal_type": {"header_name": "Abnormal", "width": 90},
                "cpk": {"header_name": "Cpk", "width": 70},
                "cpk_l1": {"header_name": "Cpk L1", "width": 70},
                "cpk_l2": {"header_name": "Cpk L2", "width": 70},
                "r1": {"header_name": "R1", "width": 65},
                "r2": {"header_name": "R2", "width": 65},
                "cpk_violation": {"header_name": "Violation", "width": 90},
                "k_value": {"header_name": "K Value", "width": 80},
                "ooc_cnt": {"header_name": "OOC", "width": 60},
                "oos_cnt": {"header_name": "OOS", "width": 60},
                "mean_index": {"header_name": "Mean Idx", "width": 90},
                "sigma_index": {"header_name": "Sigma Idx", "width": 90},
                "data_cnt": {"header_name": "N", "width": 70}
            }
            
            for col in display_df.columns:
                if col in col_settings:
                    gb.configure_column(col, **col_settings[col])
            
            gridOptions = gb.build()
            
            grid_response = AgGrid(
                display_df,
                gridOptions=gridOptions,
                update_mode=GridUpdateMode.SELECTION_CHANGED,
                fit_columns_on_grid_load=False,
                height=350,
                theme='streamlit' 
            )
            
            selected_rows = grid_response.get('selected_rows')

        with c_top_right:
            
            item = None
            if selected_rows is not None and len(selected_rows) > 0:
                if isinstance(selected_rows, pd.DataFrame):
                    sel_dict = selected_rows.iloc[0].to_dict()
                else:
                    sel_dict = selected_rows[0]
                    
                # 同時支援 OOB (group_name/chart_name) 與 Tool Matching (gname/cname/group) 欄位
                g_name = sel_dict.get('group_name') or sel_dict.get('gname')
                c_name = sel_dict.get('chart_name') or sel_dict.get('cname')
                site_id = sel_dict.get('group')  # Tool Matching 模式的 site (matching_group)
                item = next(
                    (x for x in data_list if
                     (str(x.get('group_name', '')) == str(g_name) or str(x.get('gname', '')) == str(g_name)) and
                     (str(x.get('chart_name', '')) == str(c_name) or str(x.get('cname', '')) == str(c_name)) and
                     (not site_id or str(x.get('group', '')) == str(site_id))),
                    None
                )
            
            if item:
                if item.get('no_data'):
                    _reason_map = {
                        'csv_not_found':    '找不到對應的 CSV 檔案',
                        'csv_read_error':   '無法讀取 CSV',
                        'empty':            'CSV 存在但無任何資料點',
                        'preprocess_failed':'欄位缺失或資料全為離群值',
                    }
                    _reason = _reason_map.get(item.get('no_data_reason', ''), '未知原因')
                    st.warning(f"⚠ **此 Chart 無資料可分析**（{_reason}），無圖表可顯示。")
                elif item.get('baseline_insufficient'):
                    st.warning("⚠ **基準期點數不足（< 10 筆）**，無法計算 OOB/統計分析。以下圖表僅供資料瀏覽，分析結果均為 N/A。")

                if not item.get('no_data') and st.session_state.current_mode == "OOB/SPC":
                    if item.get('chart_path'):
                        st.markdown(f"**All Data SPC - {item.get('chart_name')}**")
                        chart_data = item.get('chart_data')
                        if chart_data:
                            try:
                                _SITE_COLORS = [
                                    '#5863F8','#E83F6F','#087E8B','#FF6B35','#4CC9F0',
                                    '#7209B7','#F72585','#3A0CA3','#FBBF24','#10B981',
                                ]

                                df_pts = pd.DataFrame(chart_data)
                                df_pts['point_val'] = pd.to_numeric(df_pts['point_val'], errors='coerce')
                                df_pts = df_pts.dropna(subset=['point_val'])
                                df_pts = df_pts.sort_values('point_time').reset_index(drop=True)
                                df_pts['_idx'] = df_pts.index  # 等距 x 軸用

                                fig = go.Figure()

                                # --- 背景底色：Baseline(藍) vs Weekly(紅) ---
                                _wk_start = item.get('weekly_start')
                                _wk_end   = item.get('weekly_end')
                                if _wk_start and _wk_end and not df_pts.empty:
                                    _before_wk = df_pts[df_pts['point_time'] < _wk_start]
                                    _after_wk  = df_pts[df_pts['point_time'] > _wk_end]
                                    _wk_start_idx = (float(_before_wk['_idx'].max()) + 0.5) if not _before_wk.empty else -0.5
                                    _wk_end_idx   = (float(_after_wk['_idx'].min()) - 0.5) if not _after_wk.empty else (len(df_pts) - 0.5)
                                    fig.add_vrect(
                                        x0=-0.5, x1=_wk_start_idx,
                                        fillcolor='rgba(55,114,255,0.08)',
                                        layer='below', line_width=0,
                                    )
                                    fig.add_vrect(
                                        x0=_wk_start_idx, x1=_wk_end_idx,
                                        fillcolor='rgba(232,63,111,0.10)',
                                        layer='below', line_width=0,
                                    )

                                # --- 資料線（依 Site 分組上色）---
                                if 'Matching' in df_pts.columns:
                                    _sites = sorted(df_pts['Matching'].astype(str).unique())
                                    for _si, _site in enumerate(_sites):
                                        _grp = df_pts[df_pts['Matching'].astype(str) == _site]
                                        _color = _SITE_COLORS[_si % len(_SITE_COLORS)]
                                        fig.add_trace(go.Scatter(
                                            x=_grp['_idx'],
                                            y=_grp['point_val'],
                                            customdata=_grp['point_time'],
                                            mode='markers+lines',
                                            name=str(_site),
                                            line=dict(width=1.2, color=_color),
                                            marker=dict(size=5, color=_color,
                                                        line=dict(width=0.5, color='white')),
                                            hovertemplate=(
                                                '<b>%{fullData.name}</b><br>'
                                                '時間: %{customdata}<br>'
                                                '數值: %{y:.4g}'
                                                '<extra></extra>'
                                            ),
                                        ))
                                else:
                                    fig.add_trace(go.Scatter(
                                        x=df_pts['_idx'],
                                        y=df_pts['point_val'],
                                        customdata=df_pts['point_time'],
                                        mode='markers+lines',
                                        name='Data',
                                        line=dict(width=1.5, color='#5863F8'),
                                        marker=dict(size=5, color='#5863F8',
                                                    line=dict(width=0.5, color='white')),
                                        hovertemplate='時間: %{customdata}<br>數值: %{y:.4g}<extra></extra>',
                                    ))

                                # --- 控制線 ---
                                def _safe_float(v):
                                    try:
                                        f = float(v)
                                        return None if (f != f) else f
                                    except (TypeError, ValueError):
                                        return None

                                _ucl_v = _safe_float(item.get('UCL'))
                                _lcl_v = _safe_float(item.get('LCL'))

                                # --- 當週資料子集（用於違規標記）---
                                if _wk_start and _wk_end and not df_pts.empty:
                                    _df_wk = df_pts[
                                        (df_pts['point_time'] >= _wk_start) &
                                        (df_pts['point_time'] <= _wk_end)
                                    ].copy()
                                else:
                                    _df_wk = df_pts.copy()

                                # WE1：當週點 > UCL
                                if _ucl_v is not None and not _df_wk.empty:
                                    _we1 = _df_wk[_df_wk['point_val'] > _ucl_v]
                                    if not _we1.empty:
                                        fig.add_trace(go.Scatter(
                                            x=_we1['_idx'], y=_we1['point_val'],
                                            customdata=_we1['point_time'],
                                            mode='markers', name='WE1 (>UCL)',
                                            marker=dict(symbol='circle-open', size=16,
                                                        color='#E83F6F',
                                                        line=dict(width=2.5, color='#E83F6F')),
                                            hovertemplate=(
                                                '<b style="color:#E83F6F">⚠ WE1 違規</b><br>'
                                                '時間: %{customdata}<br>數值: %{y:.4g}'
                                                '<extra></extra>'
                                            ),
                                        ))

                                # WE5：當週點 < LCL
                                if _lcl_v is not None and not _df_wk.empty:
                                    _we5 = _df_wk[_df_wk['point_val'] < _lcl_v]
                                    if not _we5.empty:
                                        fig.add_trace(go.Scatter(
                                            x=_we5['_idx'], y=_we5['point_val'],
                                            customdata=_we5['point_time'],
                                            mode='markers', name='WE5 (<LCL)',
                                            marker=dict(symbol='circle-open', size=16,
                                                        color='#E83F6F',
                                                        line=dict(width=2.5, color='#E83F6F')),
                                            hovertemplate=(
                                                '<b style="color:#E83F6F">⚠ WE5 違規</b><br>'
                                                '時間: %{customdata}<br>數值: %{y:.4g}'
                                                '<extra></extra>'
                                            ),
                                        ))

                                # Record High：當週最高點（若本週創歷史新高）
                                if item.get('record_high') and not _df_wk.empty:
                                    _rh_val = _df_wk['point_val'].max()
                                    _rh_pts = _df_wk[_df_wk['point_val'] == _rh_val]
                                    fig.add_trace(go.Scatter(
                                        x=_rh_pts['_idx'], y=_rh_pts['point_val'],
                                        customdata=_rh_pts['point_time'],
                                        mode='markers', name='Record High ★',
                                        marker=dict(symbol='star', size=14,
                                                    color='#FBBF24',
                                                    line=dict(width=1, color='#D97706')),
                                        hovertemplate=(
                                            '<b style="color:#D97706">★ Record High</b><br>'
                                            '時間: %{customdata}<br>數值: %{y:.4g}'
                                            '<extra></extra>'
                                        ),
                                    ))

                                # Record Low：當週最低點（若本週創歷史新低）
                                if item.get('record_low') and not _df_wk.empty:
                                    _rl_val = _df_wk['point_val'].min()
                                    _rl_pts = _df_wk[_df_wk['point_val'] == _rl_val]
                                    fig.add_trace(go.Scatter(
                                        x=_rl_pts['_idx'], y=_rl_pts['point_val'],
                                        customdata=_rl_pts['point_time'],
                                        mode='markers', name='Record Low ★',
                                        marker=dict(symbol='star', size=14,
                                                    color='#4CC9F0',
                                                    line=dict(width=1, color='#0284C7')),
                                        hovertemplate=(
                                            '<b style="color:#0284C7">★ Record Low</b><br>'
                                            '時間: %{customdata}<br>數值: %{y:.4g}'
                                            '<extra></extra>'
                                        ),
                                    ))

                                for _val, _color, _label, _dash in [
                                    (_safe_float(item.get('UCL')),    '#E83F6F', 'UCL',    'dash'),
                                    (_safe_float(item.get('LCL')),    '#E83F6F', 'LCL',    'dash'),
                                    (_safe_float(item.get('Target')), '#087E8B', 'Target', 'dot'),
                                    (_safe_float(item.get('USL')),    '#FF6B35', 'USL',    'dashdot'),
                                    (_safe_float(item.get('LSL')),    '#FF6B35', 'LSL',    'dashdot'),
                                ]:
                                    if _val is not None:
                                        fig.add_hline(
                                            y=_val,
                                            line_dash=_dash,
                                            line_color=_color,
                                            line_width=1.2,
                                            annotation_text=f'<b>{_label}</b>: {_val:.4g}',
                                            annotation_position='right',
                                            annotation_font_size=10,
                                            annotation_font_color=_color,
                                            annotation_bgcolor='rgba(255,255,255,0.75)',
                                        )

                                # --- 標題（含控制線數值）---
                                _ucl = _safe_float(item.get('UCL'))
                                _lcl = _safe_float(item.get('LCL'))
                                _tgt = _safe_float(item.get('Target'))
                                _sub = '  |  '.join([
                                    p for p in [
                                        f"UCL: {_ucl:.4g}" if _ucl is not None else None,
                                        f"Target: {_tgt:.4g}" if _tgt is not None else None,
                                        f"LCL: {_lcl:.4g}" if _lcl is not None else None,
                                    ] if p
                                ])
                                _title_text = f"<b>{item.get('chart_name', '')}</b>"
                                if _sub:
                                    _title_text += f"<br><span style='font-size:11px;color:#444'>{_sub}</span>"

                                fig.update_layout(
                                    height=390,
                                    margin=dict(l=10, r=90, t=50, b=55),
                                    title=dict(
                                        text=_title_text,
                                        font=dict(size=13, color='#1a1a1a'),
                                        x=0, xanchor='left',
                                        pad=dict(l=5),
                                    ),
                                    legend=dict(
                                        orientation='h',
                                        yanchor='bottom', y=1.02,
                                        xanchor='right', x=1,
                                        font=dict(size=10, color='#222'),
                                        bgcolor='rgba(255,255,255,0.9)',
                                        bordercolor='rgba(150,150,150,0.6)',
                                        borderwidth=1,
                                    ),
                                    hovermode='closest',
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    font=dict(color='#222'),
                                )
                                _n_pts  = len(df_pts)
                                _t_step = max(1, _n_pts // 20)
                                _t_vals = list(range(0, _n_pts, _t_step))
                                _t_text = [str(df_pts.at[i, 'point_time'])[:10] for i in _t_vals]
                                fig.update_xaxes(
                                    showgrid=False, gridcolor='rgba(0,0,0,0.10)',
                                    tickangle=-90,
                                    tickmode='array',
                                    tickvals=_t_vals,
                                    ticktext=_t_text,
                                    tickfont=dict(size=9, color='#333'),
                                    showline=True, linecolor='rgba(0,0,0,0.30)',
                                    zeroline=False,
                                )
                                fig.update_yaxes(
                                    showgrid=False, gridcolor='rgba(0,0,0,0.10)',
                                    tickfont=dict(size=10, color='#333'),
                                    showline=True, linecolor='rgba(0,0,0,0.20)',
                                    zeroline=False,
                                )
                                st.plotly_chart(fig, use_container_width=True)
                            except Exception:
                                with open(item['chart_path'], 'rb') as _f: st.image(_f.read(), use_container_width=True)
                        else:
                            with open(item['chart_path'], 'rb') as _f: st.image(_f.read(), use_container_width=True)
                    else:
                        st.info("無主圖資料")
                elif st.session_state.current_mode == "CPK Dashboard" and item.get('chart_image'):
                    img_bytes = base64.b64decode(item['chart_image'])
                    st.image(img_bytes, caption=f"{item.get('group_name')} - {item.get('chart_name')}", use_container_width=True)
                elif st.session_state.current_mode == "Tool Matching":
                    spc_path = item.get('spc_chart_path')
                    if spc_path and os.path.exists(spc_path):
                        with open(spc_path, 'rb') as _f: st.image(_f.read(), use_container_width=True)
                    else:
                        st.info("此項目無 SPC 圖表。")
            else:
                st.markdown(
                    """
                    <div style='background-color: #f9f9f9; border: 2px dashed #ccc; border-radius: 10px; padding: 50px; text-align: center; margin-top: 20px;'>
                        <h2 style='color: #555;'>👈 請點擊左側表格</h2>
                        <p style='color: #777; font-size: 1.1em;'>滑鼠點擊表格中的<strong style='color: #344CB7;'>任意一列</strong>，右側將自動載入對應圖表。</p>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )

        # ==========================================
        # 下方區塊：2x2 網格放置剩餘圖表
        # ==========================================
        if item and st.session_state.current_mode == "OOB/SPC":
            
            rest_charts = [
                ("**Weekly SPC**", item.get('weekly_chart_path')),
                ("**By Tool SPC**", item.get('by_tool_color_path')),
                ("**By Tool Group SPC**", item.get('by_tool_group_path')),
                ("**Q-Q Plot**", item.get('qq_plot_path'))
            ]
            valid_rest_charts = [(title, path) for title, path in rest_charts if path and os.path.exists(path)]
            
            if valid_rest_charts:
                st.divider()
                
                for i in range(0, len(valid_rest_charts), 2):
                    bottom_cols = st.columns(2)
                    with bottom_cols[0]:
                        st.markdown(valid_rest_charts[i][0])
                        with open(valid_rest_charts[i][1], 'rb') as _f: st.image(_f.read(), use_container_width=True)
                    if i + 1 < len(valid_rest_charts):
                        with bottom_cols[1]:
                            st.markdown(valid_rest_charts[i+1][0])
                            with open(valid_rest_charts[i+1][1], 'rb') as _f: st.image(_f.read(), use_container_width=True)

        elif item and st.session_state.current_mode == "Tool Matching":
            tl_path = item.get('timeline_chart_path')
            box_path = item.get('boxplot_chart_path')
            charts_to_show = []
            if tl_path and os.path.exists(tl_path):
                charts_to_show.append(("**Timeline (All Tools)**", tl_path))
            if box_path and os.path.exists(box_path):
                charts_to_show.append(("**Boxplot**", box_path))
            if charts_to_show:
                st.divider()
                bottom_cols = st.columns(len(charts_to_show))
                for col, (title, path) in zip(bottom_cols, charts_to_show):
                    with col:
                        st.markdown(title)
                        with open(path, 'rb') as _f: st.image(_f.read(), use_container_width=True)
else:
    if st.session_state.status == "idle":
        st.markdown("<h3 style='text-align: center; color: #888; padding-top: 100px;'>點擊左上角 Settings 開始分析</h3>", unsafe_allow_html=True)