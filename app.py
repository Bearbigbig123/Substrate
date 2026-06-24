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
import json
import re
import plotly.graph_objects as go
from PIL import Image
from datetime import datetime
from st_aggrid import AgGrid, GridOptionsBuilder
from oob_eng import find_matching_file

# --- 配置 ---
API_BASE_URL = "http://localhost:8000"
EXCEL_EXPORT_VERSION = "raw_input_group_sheets_v3"
COLUMN_ALIAS_CONFIG_PATH = os.path.abspath(os.path.join("input", "column_aliases.json"))

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


def _normalize_header_name(value: str) -> str:
    text = str(value or "").strip().casefold()
    return re.sub(r"[^a-z0-9]+", "", text)


_DEFAULT_COLUMN_ALIASES = {
    "group_name": ["GroupName", "group_name", "group", "gname", "Part ID", "part_id", "part id"],
    "chart_name": ["ChartName", "chart_name", "chart", "cname", "Item Name", "item_name", "item name"],
    "point_time": ["point_time", "Point Time", "time", "datetime", "Report Time", "FT Test End Time", "report time", "ft test end time"],
    "point_val": ["point_val", "Point Value", "value", "Lot Mean", "Lot Mean Valid", "Mean", "lot mean", "lot mean valid"],
    "matching": ["Matching", "Vendor Site", "Test Site", "vendor site", "test site", "site", "tool", "tool_name"],
    "batch_id": ["Batch_ID", "batch_id", "Batch ID", "batch id", "lot", "lot_id", "lot id"],
    "mean": ["Mean", "mean", "Lot Mean", "Lot Mean Valid", "lot mean", "lot mean valid"],
    "usl": ["USL", "usl"],
    "lsl": ["LSL", "lsl"],
    "ucl": ["UCL", "ucl"],
    "lcl": ["LCL", "lcl"],
    "target": ["Target", "target"],
    "std": ["Std", "std"],
    "cpk": ["cpk", "Cpk", "CPK"],
    "material_no": ["Material_no", "material_no", "material no", "Material No"],
    "chart_id": ["ChartID", "chart_id", "chart id", "Chart Id"],
}

_DEFAULT_SITE_COLUMN_PATTERNS = [
    r"^[Ss][Ii][Tt][Ee]_?\d+$",
]


def _load_column_alias_config():
    aliases = {key: list(values) for key, values in _DEFAULT_COLUMN_ALIASES.items()}
    site_patterns = list(_DEFAULT_SITE_COLUMN_PATTERNS)

    if not os.path.exists(COLUMN_ALIAS_CONFIG_PATH):
        return aliases, site_patterns

    try:
        with open(COLUMN_ALIAS_CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            config = json.load(f)
    except Exception:
        return aliases, site_patterns

    config_aliases = config.get("aliases", {})
    if isinstance(config_aliases, dict):
        for key, values in config_aliases.items():
            if not isinstance(values, list):
                continue
            merged_values = []
            for value in aliases.get(key, []) + values:
                text = str(value).strip()
                if text and text not in merged_values:
                    merged_values.append(text)
            aliases[key] = merged_values

    config_site_patterns = config.get("site_column_patterns", [])
    if isinstance(config_site_patterns, list):
        merged_patterns = []
        for pattern in site_patterns + [str(v).strip() for v in config_site_patterns]:
            if pattern and pattern not in merged_patterns:
                merged_patterns.append(pattern)
        site_patterns = merged_patterns

    return aliases, site_patterns


_COLUMN_ALIASES, _SITE_COLUMN_PATTERNS = _load_column_alias_config()


def _build_column_lookup(columns) -> dict:
    return {_normalize_header_name(col): col for col in columns}


def _find_alias_column(columns, alias_key: str):
    lookup = _build_column_lookup(columns)
    for alias in _COLUMN_ALIASES.get(alias_key, []):
        matched = lookup.get(_normalize_header_name(alias))
        if matched is not None:
            return matched
    return None


def _has_alias_columns(columns, alias_keys) -> bool:
    return all(_find_alias_column(columns, key) is not None for key in alias_keys)


def _rename_alias_columns(df: pd.DataFrame, alias_map: dict) -> pd.DataFrame:
    renamed_df = df.copy()
    rename_pairs = {}
    for alias_key, canonical_name in alias_map.items():
        matched = _find_alias_column(renamed_df.columns, alias_key)
        if matched and matched != canonical_name:
            rename_pairs[matched] = canonical_name
    if rename_pairs:
        renamed_df = renamed_df.rename(columns=rename_pairs)
    return renamed_df

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
            norm_col = _normalize_header_name(col)
            if "groupname" in norm_col and "chartname" in norm_col:
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
        df = _rename_alias_columns(df, {
            "group_name": "GroupName",
            "chart_name": "ChartName",
            "point_time": "point_time",
            "point_val": "point_val",
            "cpk": "cpk",
            "std": "Std",
        })
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
        lot_mean_col = _find_alias_column(df.columns, "point_val")
        group_col = _find_alias_column(df.columns, "group_name")
        chart_col = _find_alias_column(df.columns, "chart_name")
        time_col = _find_alias_column(df.columns, "point_time")
        matching_col = _find_alias_column(df.columns, "matching")
        if not all([lot_mean_col, group_col, chart_col, time_col, matching_col]):
            raise ValueError("Missing vendor columns")
        df = df.rename(columns={
            group_col: "GroupName",
            chart_col: "ChartName",
            time_col: "point_time",
            lot_mean_col: "point_val",
            matching_col: "Matching",
        })
        df = _rename_alias_columns(df, {
            "cpk": "cpk",
            "std": "Std",
            "batch_id": "Batch_ID",
        })
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
        group_col = _find_alias_column(df.columns, "group_name")
        time_col = _find_alias_column(df.columns, "point_time")
        matching_col = _find_alias_column(df.columns, "matching")
        if not all([group_col, time_col, matching_col]):
            raise ValueError("Missing test columns")
        df = df.rename(columns={
            group_col: "GroupName",
            time_col: "point_time",
            matching_col: "Matching",
        })
        df = _rename_alias_columns(df, {
            "cpk": "cpk",
            "std": "Std",
            "batch_id": "Batch_ID",
        })
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


def _local_split_unified_file(input_path: str) -> dict:
    """Unified Vertical 格式本地拆分：
    回傳 {"allchartinfo_path": str, "oob_dir": str, "cpk_dir": str}
    OOB/SPC 使用 Mean 欄位當 point_val；CPK 使用 Site1…SiteN melt 後各自當 point_val。
    """
    import re as _re
    import numpy as np

    base_dir = os.path.abspath(os.path.join("temp_uploads", str(uuid.uuid4()), "unified_split"))
    os.makedirs(base_dir, exist_ok=True)

    for enc in ["utf-8-sig", "utf-8", "big5", "cp950", "latin1", "cp1252"]:
        try:
            df = pd.read_csv(input_path, encoding=enc)
            break
        except Exception:
            continue
    else:
        raise ValueError(f"無法讀取檔案: {os.path.basename(input_path)}")

    df = _rename_alias_columns(df, {
        "group_name": "GroupName",
        "chart_name": "ChartName",
        "mean": "Mean",
        "usl": "USL",
        "lsl": "LSL",
        "ucl": "UCL",
        "lcl": "LCL",
        "target": "Target",
        "material_no": "Material_no",
        "chart_id": "ChartID",
        "point_time": "point_time",
        "batch_id": "Batch_ID",
        "cpk": "cpk",
        "std": "Std",
    })

    site_cols = [
        c for c in df.columns
        if any(_re.match(pattern, str(c)) for pattern in _SITE_COLUMN_PATTERNS)
    ]
    if not site_cols:
        raise ValueError("找不到 Site 欄位（如 Site1, Site2 …）")

    required_meta = ["GroupName", "ChartName", "Mean"]
    missing = [c for c in required_meta if c not in df.columns]
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    # ── 1. 建立 allchartinfo.xlsx ──────────────────────────────────
    meta_cols = ["GroupName", "ChartName", "USL", "LSL", "UCL", "LCL", "Target"]
    existing_meta = [c for c in meta_cols if c in df.columns]
    charts_df = df[existing_meta].drop_duplicates(subset=["GroupName", "ChartName"]).copy()

    for col in ["USL", "LSL", "UCL", "LCL", "Target"]:
        if col not in charts_df.columns:
            charts_df[col] = np.nan

    if "Material_no" in df.columns:
        mat = df[["GroupName", "ChartName", "Material_no"]].drop_duplicates(subset=["GroupName", "ChartName"])
        charts_df = charts_df.merge(mat, on=["GroupName", "ChartName"], how="left")
    else:
        charts_df["Material_no"] = ""

    if "ChartID" in df.columns:
        cid = df[["GroupName", "ChartName", "ChartID"]].drop_duplicates(subset=["GroupName", "ChartName"])
        charts_df = charts_df.merge(cid, on=["GroupName", "ChartName"], how="left")
    else:
        charts_df["ChartID"] = ""

    def _infer_char(row):
        has_usl = pd.notna(row.get("USL")) and str(row.get("USL", "")).strip() != ""
        has_lsl = pd.notna(row.get("LSL")) and str(row.get("LSL", "")).strip() != ""
        if has_usl and has_lsl:
            return "Nominal"
        elif has_usl:
            return "Smaller"
        elif has_lsl:
            return "Bigger"
        return "Nominal"

    charts_df["Characteristics"] = charts_df.apply(_infer_char, axis=1)

    # ── Resolution 估算（GCD 法，從 Site 欄位量測値估算） ──────────────
    import math as _math
    def _est_res(vals):
        clean = np.array([v for v in pd.to_numeric(pd.Series(vals), errors="coerce") if pd.notna(v) and np.isfinite(v)])
        if len(clean) < 2:
            return None
        sample = clean[:500]
        max_dec = 0
        for v in sample:
            s = f"{v:.8f}".rstrip("0")
            if "." in s:
                max_dec = max(max_dec, len(s.split(".")[1]))
        max_dec = min(max_dec, 6)
        if max_dec == 0:
            return 1.0
        mult = 10 ** max_dec
        int_vals = np.round(sample * mult).astype(np.int64)
        g = 0
        for v in int_vals:
            g = _math.gcd(g, int(abs(v)))
        return round(g / mult, max_dec) if g else None

    uniq = df[["GroupName", "ChartName"]].drop_duplicates()
    res_map = {}
    for _, _r in uniq.iterrows():
        _g, _c = _r["GroupName"], _r["ChartName"]
        _m = (df["GroupName"] == _g) & (df["ChartName"] == _c)
        res_map[(_g, _c)] = _est_res(df[_m][site_cols].values.flatten())
    charts_df["Resolution"] = charts_df.apply(
        lambda r: res_map.get((r["GroupName"], r["ChartName"])), axis=1
    )

    allchartinfo_path = os.path.join(base_dir, "allchartinfo_generated.xlsx")
    charts_df.to_excel(allchartinfo_path, sheet_name="Chart", index=False)

    # ── 2. OOB/SPC per-chart CSVs (point_val = Mean) ──────────────
    oob_dir = os.path.join(base_dir, "oob_charts")
    os.makedirs(oob_dir, exist_ok=True)

    id_base = ["GroupName", "ChartName"]
    optional_id = ["point_time", "Batch_ID", "cpk"]
    oob_id_cols = id_base + [c for c in optional_id if c in df.columns]

    def _sanitize(name: str) -> str:
        for ch in '<>:"/\\|?*\'':
            name = name.replace(ch, "")
        return name.strip()

    uniq = df[["GroupName", "ChartName"]].drop_duplicates()
    for _, row in uniq.iterrows():
        gname, cname = row["GroupName"], row["ChartName"]
        mask = (df["GroupName"] == gname) & (df["ChartName"] == cname)
        sub = df[mask].copy()
        oob_df = sub[[c for c in oob_id_cols if c in sub.columns]].copy()
        oob_df["point_val"] = pd.to_numeric(sub["Mean"], errors="coerce")
        oob_df = oob_df.dropna(subset=["point_val"])
        if "point_time" in oob_df.columns:
            try:
                oob_df["point_time"] = pd.to_datetime(oob_df["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
            except Exception:
                pass
        if not oob_df.empty:
            oob_df.to_csv(os.path.join(oob_dir, f"{_sanitize(str(gname))}_{_sanitize(str(cname))}.csv"), index=False, encoding="utf-8-sig")

    # ── 3. CPK per-chart CSVs (melt Site cols → point_val) ─────────
    cpk_dir = os.path.join(base_dir, "cpk_charts")
    os.makedirs(cpk_dir, exist_ok=True)

    cpk_id_cols = id_base + [c for c in optional_id if c in df.columns]

    for _, row in uniq.iterrows():
        gname, cname = row["GroupName"], row["ChartName"]
        mask = (df["GroupName"] == gname) & (df["ChartName"] == cname)
        sub = df[mask].copy()
        avail_id = [c for c in cpk_id_cols if c in sub.columns]
        avail_sites = [c for c in site_cols if c in sub.columns]
        melted = sub.melt(id_vars=avail_id, value_vars=avail_sites, var_name="site_id", value_name="point_val")
        melted["point_val"] = pd.to_numeric(melted["point_val"], errors="coerce")
        melted = melted.dropna(subset=["point_val"])
        if "point_time" in melted.columns:
            try:
                melted["point_time"] = pd.to_datetime(melted["point_time"], errors="coerce").dt.strftime("%Y/%m/%d %H:%M")
            except Exception:
                pass
        if not melted.empty:
            melted.to_csv(os.path.join(cpk_dir, f"{_sanitize(str(gname))}_{_sanitize(str(cname))}.csv"), index=False, encoding="utf-8-sig")

    print(f"[UnifiedSplit] charts={len(uniq)}, allchartinfo={allchartinfo_path}, oob={oob_dir}, cpk={cpk_dir}")
    return {"allchartinfo_path": allchartinfo_path, "oob_dir": oob_dir, "cpk_dir": cpk_dir}


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
if 'saved_unified_allchartinfo_path' not in st.session_state: st.session_state.saved_unified_allchartinfo_path = None
if 'saved_unified_oob_dir' not in st.session_state: st.session_state.saved_unified_oob_dir = None
if 'saved_unified_cpk_dir' not in st.session_state: st.session_state.saved_unified_cpk_dir = None
if 'saved_weekly_summary_path' not in st.session_state: st.session_state.saved_weekly_summary_path = None
if 'analysis_base_date' not in st.session_state: st.session_state.analysis_base_date = None

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
def _oob_count(rule_str):
    if rule_str is None or str(rule_str).strip().upper() in {"N/A", "NAN", "NONE", "", "-"}:
        return 0
    return len([t for t in str(rule_str).split(",") if t.strip()])


def _to_int(value):
    try:
        if value is None or str(value).strip().upper() in {"", "N/A", "NAN", "NONE", "-", "NAT"}:
            return 0
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _get_result_group_name(record) -> str:
    return str(record.get("group_name") or record.get("gname") or record.get("group") or "Unknown")


def _get_result_chart_name(record) -> str:
    return str(record.get("chart_name") or record.get("cname") or "")


def _week_label(*candidates):
    now = pd.Timestamp.now()
    iso_week = int(now.isocalendar().week)
    return f"W{now.year % 10}{iso_week:02d}"


def _week_sort_key(label):
    text = str(label).strip().upper()
    if text.startswith("W") and text[1:].isdigit():
        return int(text[1:])
    return 99999


def _build_oob_summary_tables(data_list, analysis_base_date=None):
    kpi_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
    group_week_data = {}
    week_data = {}

    for record in data_list:
        group_name = _get_result_group_name(record)
        week_label = _week_label(
            record.get("weekly_end"),
            record.get("weekly_start"),
            analysis_base_date,
        )
        if not week_label:
            continue

        if group_name not in group_week_data:
            group_week_data[group_name] = {}
        if week_label not in group_week_data[group_name]:
            group_week_data[group_name][week_label] = {kpi: 0 for kpi in kpi_names}

        group_week_data[group_name][week_label]["OOB"] += _oob_count(record.get("OOB_Rule"))
        group_week_data[group_name][week_label]["OOC"] += _to_int(record.get("ooc_cnt"))
        group_week_data[group_name][week_label]["OOS"] += _to_int(record.get("oos_cnt"))
        group_week_data[group_name][week_label]["CPK<1.33"] += _to_int(record.get("cpk_below_133_cnt"))

        if week_label not in week_data:
            week_data[week_label] = {kpi: 0 for kpi in kpi_names}
        week_data[week_label]["OOB"] += _oob_count(record.get("OOB_Rule"))
        week_data[week_label]["OOC"] += _to_int(record.get("ooc_cnt"))
        week_data[week_label]["OOS"] += _to_int(record.get("oos_cnt"))
        week_data[week_label]["CPK<1.33"] += _to_int(record.get("cpk_below_133_cnt"))

    group_week_df = pd.DataFrame()
    if group_week_data:
        sorted_group_weeks = sorted(
            {week for group_data in group_week_data.values() for week in group_data.keys()},
            key=_week_sort_key,
        )
        group_week_rows = []
        group_week_index = []
        for group_name in sorted(group_week_data.keys()):
            for kpi in kpi_names:
                group_week_index.append((group_name, kpi))
                group_week_rows.append([
                    group_week_data[group_name].get(week, {}).get(kpi, 0)
                    for week in sorted_group_weeks
                ])
        group_week_df = pd.DataFrame(
            group_week_rows,
            index=pd.MultiIndex.from_tuples(group_week_index, names=["Group", "KPI"]),
            columns=sorted_group_weeks,
        )

    trend_df = pd.DataFrame()
    if week_data:
        sorted_weeks = sorted(week_data.keys(), key=_week_sort_key)
        trend_df = pd.DataFrame(
            {
                week: [
                    week_data[week]["OOB"],
                    week_data[week]["OOC"],
                    week_data[week]["OOS"],
                    week_data[week]["CPK<1.33"],
                ]
                for week in sorted_weeks
            },
            index=kpi_names,
        )

    return trend_df, group_week_df


def _format_group_week_display_df(group_week_df: pd.DataFrame) -> pd.DataFrame:
    if group_week_df is None or group_week_df.empty:
        return pd.DataFrame()
    display_df = group_week_df.reset_index()
    display_df.loc[display_df["Group"].duplicated(), "Group"] = ""
    return display_df


def _get_export_raw_data_directory(mode: str):
    original_raw_dir = st.session_state.get("saved_raw_dir")
    if original_raw_dir:
        return original_raw_dir
    if mode == "OOB/SPC":
        return (
            st.session_state.get("saved_unified_oob_dir")
            or st.session_state.get("saved_split_raw_dir")
            or st.session_state.get("saved_raw_dir")
        )
    if mode == "CPK Dashboard":
        return (
            st.session_state.get("saved_unified_cpk_dir")
            or st.session_state.get("saved_split_raw_dir")
            or st.session_state.get("saved_raw_dir")
        )
    return st.session_state.get("saved_split_raw_dir") or st.session_state.get("saved_raw_dir")


def _sanitize_excel_sheet_name(name: str, used_names):
    cleaned = re.sub(r"[\[\]:*?/\\\\]", "_", str(name)).strip()
    cleaned = cleaned.strip("'") or "Sheet"
    cleaned = cleaned[:31]
    base = cleaned or "Sheet"
    candidate = base
    suffix = 1
    while candidate in used_names:
        suffix_text = f"_{suffix}"
        candidate = f"{base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def _normalize_group_key(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.upper() in {"", "N/A", "NONE", "NAN", "<NA>", "UNKNOWN"}:
        return ""
    return text.casefold()


def _collect_group_rawdata_frames(data_list, raw_data_directory: str):
    if not raw_data_directory or not os.path.isdir(raw_data_directory):
        return []

    csv_paths = sorted(
        os.path.join(raw_data_directory, name)
        for name in os.listdir(raw_data_directory)
        if name.lower().endswith(".csv")
    )
    if not csv_paths:
        return []

    result_group_map = {}
    result_chart_map = {}
    for record in data_list:
        group_name = _get_result_group_name(record)
        group_key = _normalize_group_key(group_name)
        if group_key and group_key not in result_group_map:
            result_group_map[group_key] = group_name
        chart_name = _get_result_chart_name(record)
        if group_key and chart_name:
            result_chart_map.setdefault(group_key, set()).add(str(chart_name).strip())

    grouped_frames_map = {}
    grouped_columns = {}
    for csv_path in csv_paths:
        try:
            raw_df = _local_read_csv(csv_path, header_val="infer")
        except Exception:
            continue
        if raw_df is None or raw_df.empty:
            continue

        raw_df.columns = [str(c).strip() for c in raw_df.columns]
        matched_group_col = _find_alias_column(raw_df.columns, "group_name")
        matched_chart_col = _find_alias_column(raw_df.columns, "chart_name")

        if matched_group_col:
            for raw_group_name, sub_df in raw_df.groupby(matched_group_col, dropna=False, sort=False):
                group_key = _normalize_group_key(raw_group_name)
                if not group_key:
                    continue
                display_group_name = result_group_map.get(group_key, str(raw_group_name).strip())

                grouped_frames_map.setdefault(display_group_name, [])
                grouped_columns.setdefault(display_group_name, [])

                for col_name in sub_df.columns:
                    if col_name not in grouped_columns[display_group_name]:
                        grouped_columns[display_group_name].append(col_name)
                grouped_frames_map[display_group_name].append(sub_df.copy())
            continue

        for group_key, display_group_name in result_group_map.items():
            candidate_charts = sorted(result_chart_map.get(group_key, set()))
            matched_csv = None
            for chart_name in candidate_charts:
                found_path = find_matching_file(raw_data_directory, display_group_name, chart_name)
                if found_path and os.path.abspath(found_path) == os.path.abspath(csv_path):
                    matched_csv = found_path
                    break
            if not matched_csv:
                continue

            grouped_frames_map.setdefault(display_group_name, [])
            grouped_columns.setdefault(display_group_name, [])

            for col_name in raw_df.columns:
                if col_name not in grouped_columns[display_group_name]:
                    grouped_columns[display_group_name].append(col_name)
            grouped_frames_map[display_group_name].append(raw_df.copy())

    grouped_frames = []
    for group_name in sorted(grouped_frames_map.keys()):
        frames = grouped_frames_map.get(group_name) or []
        columns = grouped_columns.get(group_name) or []
        if not frames or not columns:
            continue
        merged_raw_df = pd.concat(frames, ignore_index=True, sort=False)
        remaining_columns = [c for c in merged_raw_df.columns if c not in columns]
        merged_raw_df = merged_raw_df[columns + remaining_columns]
        merged_raw_df = merged_raw_df.drop_duplicates(ignore_index=True)
        grouped_frames.append((group_name, merged_raw_df))
    return grouped_frames


def _write_dataframe_to_sheet(worksheet, df: pd.DataFrame, start_row: int, title: str, header_format, cell_format) -> int:
    worksheet.write(start_row, 0, title, header_format)
    if df is None or df.empty:
        worksheet.write(start_row + 1, 0, "No data", cell_format)
        return start_row + 3

    display_df = df.copy()
    for col_idx, col_name in enumerate(display_df.columns):
        worksheet.write(start_row + 1, col_idx, str(col_name), header_format)

    for row_offset, (_, row) in enumerate(display_df.iterrows(), start=2):
        for col_idx, value in enumerate(row):
            if pd.isna(value):
                value = ""
            worksheet.write(start_row + row_offset, col_idx, value, cell_format)

    for col_idx, col_name in enumerate(display_df.columns):
        max_len = len(str(col_name))
        for value in display_df.iloc[:, col_idx].tolist():
            value_str = "" if pd.isna(value) else str(value)
            max_len = max(max_len, len(value_str))
        worksheet.set_column(col_idx, col_idx, min(max(max_len + 2, 12), 40))

    if display_df.columns.size > 0 and str(display_df.columns[0]).strip().upper() in {"KPI", "GROUP"}:
        worksheet.set_column(0, 0, 28)

    return start_row + len(display_df) + 4


def _is_valid_file_path(value) -> bool:
    if not isinstance(value, str):
        return False
    if value.strip().upper() in {"", "N/A", "NONE", "NAN", "<NA>"}:
        return False
    return os.path.exists(value)


def generate_full_excel_with_images(
    data_list,
    mode,
    summary_weekly_df=None,
    summary_group_weekly_df=None,
    raw_data_directory=None,
):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    include_summary_sheet = mode == "OOB/SPC"
    summary_sheet = workbook.add_worksheet('Summary') if include_summary_sheet else None
    worksheet = workbook.add_worksheet('Analysis_Results')

    df = pd.DataFrame(data_list)

    # 處理 Metrics 展開 (用於 CPK 模式)
    if 'metrics' in df.columns:
        metrics_df = pd.json_normalize(df['metrics'])
        df = pd.concat([df.drop(columns=['metrics']), metrics_df], axis=1)

    # 定義 Excel 樣式
    header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11, 'bold': True, 'bg_color': '#344CB7', 'font_color': 'white'})
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10})

    if include_summary_sheet:
        summary_row = 0
        if summary_weekly_df is not None:
            summary_weekly_display_df = summary_weekly_df.reset_index().rename(columns={"index": "KPI"})
            summary_row = _write_dataframe_to_sheet(
                summary_sheet,
                summary_weekly_display_df,
                summary_row,
                "Weekly Summary",
                header_format,
                cell_format,
            )
        if summary_group_weekly_df is not None:
            summary_group_display_df = _format_group_week_display_df(summary_group_weekly_df)
            summary_row = _write_dataframe_to_sheet(
                summary_sheet,
                summary_group_display_df,
                summary_row,
                "Group Weekly Summary",
                header_format,
                cell_format,
            )
        summary_sheet.freeze_panes(1, 0)

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

    active_oob_img_fields = img_fields
    if is_oob_mode:
        active_oob_img_fields = [
            (key, title)
            for key, title in img_fields
            if key in df.columns and df[key].apply(_is_valid_file_path).any()
        ]

    if is_oob_mode:
        num_img_cols = len(active_oob_img_fields)
    elif is_cpk_mode:
        num_img_cols = len(cpk_img_fields)
    elif is_tm_mode:
        num_img_cols = len(tm_img_fields)
    else:
        num_img_cols = 0
    start_col = num_img_cols

    # 寫入圖片的表頭
    if is_oob_mode:
        for i, (key, title) in enumerate(active_oob_img_fields):
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
            for col_offset, (key, _) in enumerate(active_oob_img_fields):
                img_path = row.get(key)
                if _is_valid_file_path(img_path):
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
                if _is_valid_file_path(img_path):
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

    used_sheet_names = {"Analysis_Results"}
    if include_summary_sheet:
        used_sheet_names.add("Summary")
    for group_name, raw_df in _collect_group_rawdata_frames(data_list, raw_data_directory):
        sheet_name = _sanitize_excel_sheet_name(f"RAW_{group_name}", used_sheet_names)
        raw_sheet = workbook.add_worksheet(sheet_name)
        _write_dataframe_to_sheet(
            raw_sheet,
            raw_df,
            0,
            group_name,
            header_format,
            cell_format,
        )
        raw_sheet.freeze_panes(2, 0)

    workbook.close()
    if _cpk_tmp_dir and os.path.exists(_cpk_tmp_dir):
        import shutil
        shutil.rmtree(_cpk_tmp_dir, ignore_errors=True)
    return output.getvalue()


def _read_weekly_summary_csv(source):
    """Load historical weekly summary CSV into overall/group weekly summary tables."""
    if source is None:
        return pd.DataFrame(), pd.DataFrame()

    try:
        df = pd.read_csv(source, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(source, encoding="big5")

    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    metric_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
    normalized_cols = {str(c).strip().lower(): c for c in df.columns}
    group_col = next((normalized_cols[c] for c in ["group", "group_name", "gname"] if c in normalized_cols), None)

    def _coerce_week_label(raw_value):
        week_label = str(raw_value or "").strip()
        if not week_label:
            return ""
        if not week_label.upper().startswith("W"):
            try:
                dt = pd.to_datetime(week_label)
                week_label = f"W{dt.month}{dt.day:02d}"
            except Exception:
                pass
        return week_label

    def _build_group_week_df(long_df: pd.DataFrame, week_col_name: str, group_col_name: str) -> pd.DataFrame:
        grouped_data = {}
        for _, row in long_df.iterrows():
            week_label = _coerce_week_label(row.get(week_col_name))
            group_name = str(row.get(group_col_name, "")).strip()
            if not week_label or not group_name:
                continue
            group_upper = group_name.upper()
            if group_upper in {"ALL", "OVERALL", "TOTAL"}:
                continue
            grouped_data.setdefault(group_name, {})
            grouped_data[group_name][week_label] = {
                metric: _to_int(row.get(metric)) if metric in long_df.columns else 0
                for metric in metric_names
            }

        if not grouped_data:
            return pd.DataFrame()

        sorted_weeks = sorted(
            {week for group_weeks in grouped_data.values() for week in group_weeks.keys()},
            key=_week_sort_key,
        )
        rows = []
        index = []
        for group_name in sorted(grouped_data.keys()):
            for metric in metric_names:
                index.append((group_name, metric))
                rows.append([
                    grouped_data[group_name].get(week, {}).get(metric, 0)
                    for week in sorted_weeks
                ])
        out_df = pd.DataFrame(
            rows,
            index=pd.MultiIndex.from_tuples(index, names=["Group", "KPI"]),
            columns=sorted_weeks,
        )
        for col in out_df.columns:
            out_df[col] = pd.to_numeric(out_df[col], errors="coerce").fillna(0).astype(int)
        return out_df

    # Long format: Week,OOB,OOC,OOS,CPK<1.33
    week_col = next((normalized_cols[c] for c in ["week", "weekly", "weekly_end", "date"] if c in normalized_cols), None)
    if week_col and any(m in df.columns for m in metric_names):
        if group_col:
            aggregate_cols = [week_col, group_col]
        else:
            aggregate_cols = [week_col]
        metric_cols = [metric for metric in metric_names if metric in df.columns]
        if metric_cols:
            df = (
                df.groupby(aggregate_cols, dropna=False, as_index=False)[metric_cols]
                .sum(min_count=1)
            )

        overall_source_df = df
        if group_col:
            group_series = df[group_col].astype(str).str.strip().str.upper()
            overall_source_df = df[group_series.isin(["", "ALL", "OVERALL", "TOTAL"])].copy()

        overall_df = pd.DataFrame(index=metric_names)
        for _, row in overall_source_df.iterrows():
            week_label = _coerce_week_label(row.get(week_col))
            if not week_label:
                continue
            overall_df[week_label] = [
                _to_int(row.get(metric)) if metric in overall_source_df.columns else 0
                for metric in metric_names
            ]

        group_week_df = _build_group_week_df(df, week_col, group_col) if group_col else pd.DataFrame()
        if not overall_df.empty:
            overall_df = overall_df.fillna(0).astype(int)
        return overall_df, group_week_df

    # Wide format: first column is metric, remaining columns are weeks.
    first_col = df.columns[0]
    metric_series = df[first_col].astype(str).str.strip()
    if set(metric_names).intersection(set(metric_series)):
        wide = df.copy()
        wide[first_col] = metric_series
        wide = wide.set_index(first_col)
        wide = wide.reindex(metric_names).fillna(0)
        for col in wide.columns:
            wide[col] = pd.to_numeric(wide[col], errors="coerce").fillna(0).astype(int)
        return wide, pd.DataFrame()

    return pd.DataFrame(), pd.DataFrame()


def _merge_weekly_summary_tables(history_df: pd.DataFrame, current_df: pd.DataFrame) -> pd.DataFrame:
    metric_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
    frames = [df for df in [history_df, current_df] if df is not None and not df.empty]
    if not frames:
        return pd.DataFrame()

    merged = pd.concat(frames, axis=1)
    merged = merged.loc[:, ~merged.columns.duplicated(keep="last")]
    merged = merged.reindex(metric_names).fillna(0)
    for col in merged.columns:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0).astype(int)
    return merged


def _merge_group_weekly_summary_tables(history_df: pd.DataFrame, current_df: pd.DataFrame) -> pd.DataFrame:
    metric_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
    frames = [df for df in [history_df, current_df] if df is not None and not df.empty]
    if not frames:
        return pd.DataFrame()

    merged = pd.concat(frames, axis=1)
    merged = merged.loc[:, ~merged.columns.duplicated(keep="last")]
    if isinstance(merged.index, pd.MultiIndex):
        ordered_index = []
        group_names = sorted(dict.fromkeys(merged.index.get_level_values("Group")))
        for group_name in group_names:
            for metric_name in metric_names:
                key = (group_name, metric_name)
                if key in merged.index:
                    ordered_index.append(key)
        merged = merged.reindex(pd.MultiIndex.from_tuples(ordered_index, names=["Group", "KPI"]))
    merged = merged.fillna(0)
    for col in merged.columns:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0).astype(int)
    return merged


def _auto_weekly_summary_path() -> str:
    return os.path.abspath(os.path.join("input", "weekly_summary_history.csv"))


def _save_weekly_summary_csv(summary_df: pd.DataFrame, group_summary_df: pd.DataFrame, output_path: str) -> None:
    if (summary_df is None or summary_df.empty) and (group_summary_df is None or group_summary_df.empty):
        return
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    metric_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
    rows = []

    if summary_df is not None and not summary_df.empty:
        for week in sorted(summary_df.columns, key=_week_sort_key):
            rows.append({
                "Week": week,
                "Group": "ALL",
                "OOB": _to_int(summary_df.at["OOB", week]) if "OOB" in summary_df.index else 0,
                "OOC": _to_int(summary_df.at["OOC", week]) if "OOC" in summary_df.index else 0,
                "OOS": _to_int(summary_df.at["OOS", week]) if "OOS" in summary_df.index else 0,
                "CPK<1.33": _to_int(summary_df.at["CPK<1.33", week]) if "CPK<1.33" in summary_df.index else 0,
            })

    if group_summary_df is not None and not group_summary_df.empty:
        group_df = group_summary_df.copy()
        if isinstance(group_df.index, pd.MultiIndex):
            for (group_name, metric_name), row in group_df.iterrows():
                if metric_name not in metric_names:
                    continue
                for week in sorted(group_df.columns, key=_week_sort_key):
                    existing_row = next(
                        (
                            item for item in rows
                            if item["Week"] == week and item["Group"] == str(group_name)
                        ),
                        None,
                    )
                    if existing_row is None:
                        existing_row = {
                            "Week": week,
                            "Group": str(group_name),
                            "OOB": 0,
                            "OOC": 0,
                            "OOS": 0,
                            "CPK<1.33": 0,
                        }
                        rows.append(existing_row)
                    existing_row[metric_name] = _to_int(row.get(week))

    if not rows:
        return

    out_df = pd.DataFrame(rows, columns=["Week", "Group", "OOB", "OOC", "OOS", "CPK<1.33"])
    out_df["__group_sort"] = out_df["Group"].astype(str).str.upper().map(lambda x: 0 if x == "ALL" else 1)
    out_df["__week_sort"] = out_df["Week"].map(_week_sort_key)
    out_df = out_df.sort_values(by=["__week_sort", "__group_sort", "Group"]).drop(columns=["__group_sort", "__week_sort"])
    out_df.to_csv(output_path, index=False, encoding="utf-8-sig")


# ==========================================
# 1. 頂部導覽橫條 (Header)
# ==========================================
col1, col_dl, col2, col3, col4 = st.columns([1.5, 1.8, 2.0, 0.8, 0.8], gap="medium")

with col1:
    with st.popover("⚙️ Settings & Run", use_container_width=True):
        st.markdown("##### 分析設定")
        mode = st.radio("選擇功能", ["OOB/SPC", "CPK Dashboard"])
        base_date = st.date_input("分析基準日", value=datetime.now())
        
       
        # --- 檔案上傳區塊 (水平排列) ---
        st.markdown("###### 📁 上傳自訂檔案 (若不傳則使用預設)")
        csv_files = st.file_uploader("Raw Data (CSV, 多選)", type=["csv"], accept_multiple_files=True, help="上傳產線原始資料 CSV 檔")

        weekly_summary_file = None
        if mode == "OOB/SPC":
            weekly_summary_file = st.file_uploader(
                "Optional: Historical Weekly Summary (CSV)",
                type=["csv"],
                help="支援 Week,OOB,OOC,OOS,CPK<1.33、Week,Group,OOB,OOC,OOS,CPK<1.33，或目前 Weekly Summary 寬表格式。Group=ALL 代表整體 weekly summary。",
            )
            
        st.divider()
        
        if st.button("🚀 Start Analysis", type="primary", use_container_width=True):
            st.session_state.current_mode = mode
            st.session_state.results = None
            st.session_state.status = "idle"
            st.session_state.progress = 0
            
            current_raw_dir = None

            # --- 1. 處理「新上傳」的檔案 ---
            if csv_files or weekly_summary_file:
                upload_session_id = str(uuid.uuid4())
                base_upload_dir = os.path.abspath(os.path.join("temp_uploads", "ui_uploads", upload_session_id))
                os.makedirs(base_upload_dir, exist_ok=True)

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
                    st.session_state.saved_unified_allchartinfo_path = None
                    st.session_state.saved_unified_oob_dir = None
                    st.session_state.saved_unified_cpk_dir = None

                if weekly_summary_file:
                    weekly_summary_file.seek(0)
                    weekly_summary_path = os.path.join(base_upload_dir, weekly_summary_file.name)
                    with open(weekly_summary_path, "wb") as f:
                        f.write(weekly_summary_file.read())
                    st.session_state.saved_weekly_summary_path = weekly_summary_path

            # --- 2. 自動偵測與拆分 (僅在新上傳 CSV 時觸發) ---
            if csv_files and len(csv_files) == 1 and current_raw_dir:
                first_saved = os.path.join(current_raw_dir, csv_files[0].name)
                try:
                    peek = pd.read_csv(first_saved, nrows=0)
                    detected_cols = set(peek.columns)
                    detected_split_mode = None
                    normalized_detected_cols = {_normalize_header_name(c) for c in detected_cols}
                    
                    import re as _split_re
                    _has_site = any(
                        any(_split_re.match(pattern, str(c)) for pattern in _SITE_COLUMN_PATTERNS)
                        for c in detected_cols
                    )
                    if _has_alias_columns(detected_cols, ["group_name", "chart_name", "point_time", "matching"]) and _find_alias_column(detected_cols, "point_val"):
                        detected_split_mode = "Vendor_Vertical"
                    elif _has_alias_columns(detected_cols, ["group_name", "point_time", "matching"]):
                        detected_split_mode = "Test_Horizontal"
                    elif _has_alias_columns(detected_cols, ["group_name", "chart_name", "mean"]) and _has_site:
                        detected_split_mode = "Unified_Vertical"
                    elif _has_alias_columns(detected_cols, ["group_name", "chart_name", "point_time", "point_val"]):
                        detected_split_mode = "Type2_Vertical"
                    else:
                        peek_no_header = pd.read_csv(first_saved, nrows=3, header=None)
                        flat_vals = peek_no_header.iloc[0:2].fillna("").astype(str).values.flatten().tolist()
                        normalized_flat_vals = [_normalize_header_name(val) for val in flat_vals]
                        if any("groupname" in val for val in normalized_flat_vals) and any("chartname" in val for val in normalized_flat_vals):
                            detected_split_mode = "Type3_Horizontal"

                    if detected_split_mode == "Unified_Vertical":
                        try:
                            with st.spinner("正在自動準備資料夾..."):
                                unified_res = _local_split_unified_file(first_saved)
                            st.session_state.saved_unified_allchartinfo_path = unified_res["allchartinfo_path"]
                            st.session_state.saved_unified_oob_dir = unified_res["oob_dir"]
                            st.session_state.saved_unified_cpk_dir = unified_res["cpk_dir"]
                            st.session_state.saved_split_raw_dir = None
                            st.session_state.saved_split_id = None
                            st.session_state.saved_split_info = "🔀 自動偵測到 **Unified_Vertical** 格式，已完成拆分"
                        except Exception as split_err:
                            st.error(f"⚠️ 自動拆分失敗：{str(split_err)}")
                    elif detected_split_mode:
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
            if not current_raw_dir: current_raw_dir = st.session_state.get("saved_raw_dir")
            
            auto_split_raw_dir = st.session_state.get("saved_split_raw_dir")
            st.session_state.auto_split_info = st.session_state.get("saved_split_info")

            _unified_allchartinfo = st.session_state.get("saved_unified_allchartinfo_path")
            _unified_oob_dir = st.session_state.get("saved_unified_oob_dir")
            _unified_cpk_dir = st.session_state.get("saved_unified_cpk_dir")
            _is_unified = bool(_unified_allchartinfo and _unified_oob_dir and _unified_cpk_dir)

            # --- 4. 防呆機制：如果真的沒檔案可送，擋住並警告 ---
            if not current_raw_dir and not auto_split_raw_dir and not _is_unified:
                st.error("⚠️ 系統找不到分析資料，請重新上傳檔案！")
                st.stop()

            # --- 5. 組裝 API Payload ---
            payload = {}
            if mode == "OOB/SPC":
                endpoint = "/process"
                payload["base_date"] = base_date.strftime("%Y-%m-%d")
                st.session_state.analysis_base_date = base_date.strftime("%Y-%m-%d")
                if _is_unified:
                    payload["filepath"] = _unified_allchartinfo
                    payload["raw_data_directory"] = _unified_oob_dir
                else:
                    if auto_split_raw_dir: payload["raw_data_directory"] = auto_split_raw_dir
                    elif current_raw_dir: payload["raw_data_directory"] = current_raw_dir

            else: # CPK Dashboard
                endpoint = "/spc-cpk"
                payload = {"end_date": base_date.strftime("%Y-%m-%d")}
                if _is_unified:
                    payload["chart_excel_path"] = _unified_allchartinfo
                    payload["raw_data_directory"] = _unified_cpk_dir
                else:
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

with col_dl:
    if (
        st.session_state.get("status") == "completed"
        and st.session_state.get("results")
        and (
            st.session_state.get("full_excel_data") is None
            or st.session_state.get("full_excel_export_version") != EXCEL_EXPORT_VERSION
        )
    ):
        _export_res = st.session_state.results
        _export_data_list = []
        if st.session_state.current_mode in ["OOB/SPC", "Tool Matching"]:
            _export_data_list = _export_res.get("results", [])
        elif st.session_state.current_mode == "CPK Dashboard":
            _export_data_list = _export_res.get("charts", [])

        _export_trend_df = pd.DataFrame()
        _export_group_week_df = pd.DataFrame()
        if st.session_state.current_mode == "OOB/SPC" and _export_data_list:
            _current_trend_df, _export_group_week_df = _build_oob_summary_tables(
                _export_data_list,
                st.session_state.get("analysis_base_date"),
            )
            _export_history_df = pd.DataFrame()
            _export_group_history_df = pd.DataFrame()
            _export_auto_summary_path = _auto_weekly_summary_path()
            _export_weekly_summary_path = st.session_state.get("saved_weekly_summary_path")
            if not _export_weekly_summary_path and os.path.exists(_export_auto_summary_path):
                _export_weekly_summary_path = _export_auto_summary_path
            if _export_weekly_summary_path and os.path.exists(_export_weekly_summary_path):
                try:
                    _export_history_df, _export_group_history_df = _read_weekly_summary_csv(_export_weekly_summary_path)
                except Exception:
                    _export_history_df = pd.DataFrame()
                    _export_group_history_df = pd.DataFrame()
            _export_trend_df = _merge_weekly_summary_tables(_export_history_df, _current_trend_df)
            _export_group_week_df = _merge_group_weekly_summary_tables(_export_group_history_df, _export_group_week_df)
            if not _export_trend_df.empty:
                _export_trend_df = _export_trend_df[sorted(_export_trend_df.columns, key=_week_sort_key)]
            if not _export_group_week_df.empty:
                _export_group_week_df = _export_group_week_df[sorted(_export_group_week_df.columns, key=_week_sort_key)]

        st.session_state.full_excel_data = generate_full_excel_with_images(
            _export_data_list,
            st.session_state.current_mode,
            summary_weekly_df=_export_trend_df,
            summary_group_weekly_df=_export_group_week_df,
            raw_data_directory=_get_export_raw_data_directory(st.session_state.current_mode),
        )
        st.session_state.full_excel_export_version = EXCEL_EXPORT_VERSION

    if st.session_state.get('full_excel_data'):
        st.download_button(
            label="📥 下載 Excel 報告",
            data=st.session_state.full_excel_data,
            file_name=f"SPC_Full_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

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

        # 如果是新的任務，則在背景產生一份 Excel 暫存在 Session 裡
        _group_week_df = pd.DataFrame()
        _trend_df = pd.DataFrame()
        if st.session_state.get("current_mode") == "OOB/SPC":
            _current_trend_df, _group_week_df = _build_oob_summary_tables(
                data_list,
                st.session_state.get("analysis_base_date"),
            )
            _history_df = pd.DataFrame()
            _group_history_df = pd.DataFrame()
            _auto_summary_path = _auto_weekly_summary_path()
            _weekly_summary_path = st.session_state.get("saved_weekly_summary_path")
            if not _weekly_summary_path and os.path.exists(_auto_summary_path):
                _weekly_summary_path = _auto_summary_path
            if _weekly_summary_path and os.path.exists(_weekly_summary_path):
                try:
                    _history_df, _group_history_df = _read_weekly_summary_csv(_weekly_summary_path)
                except Exception as _weekly_import_err:
                    st.warning(f"Historical Weekly Summary CSV could not be loaded: {_weekly_import_err}")
            _trend_df = _merge_weekly_summary_tables(_history_df, _current_trend_df)
            _group_week_df = _merge_group_weekly_summary_tables(_group_history_df, _group_week_df)
            if not _trend_df.empty:
                _trend_df = _trend_df[sorted(_trend_df.columns, key=_week_sort_key)]
            if not _group_week_df.empty:
                _group_week_df = _group_week_df[sorted(_group_week_df.columns, key=_week_sort_key)]
                try:
                    _save_weekly_summary_csv(_trend_df, _group_week_df, _auto_summary_path)
                    st.session_state.saved_weekly_summary_path = _auto_summary_path
                except Exception as _weekly_save_err:
                    st.warning(f"Weekly Summary CSV could not be saved: {_weekly_save_err}")

        if (
            st.session_state.last_task_id != st.session_state.task_id
            or st.session_state.full_excel_data is None
            or st.session_state.get("full_excel_export_version") != EXCEL_EXPORT_VERSION
        ):
            st.session_state.full_excel_data = generate_full_excel_with_images(
                data_list,
                st.session_state.current_mode,
                summary_weekly_df=_trend_df,
                summary_group_weekly_df=_group_week_df,
                raw_data_directory=_get_export_raw_data_directory(st.session_state.current_mode),
            )
            st.session_state.last_task_id = st.session_state.task_id
            st.session_state.full_excel_export_version = EXCEL_EXPORT_VERSION

        if st.session_state.get("current_mode") == "OOB/SPC":
            summary_tab, oob_chart_tab = st.tabs(["Summary", "OOB + 圖"])
        else:
            summary_tab = st.container()
            oob_chart_tab = st.container()

        with summary_tab:
            # ==========================================
            # Weekly Trend Table（橫向：週別；縱向：OOB/OOC/OOS/CPK<1.33）
            # ==========================================
            if st.session_state.get("current_mode") == "OOB/SPC":
                def _oob_count(rule_str):
                    """計算 OOB_Rule 中違規規則數量（逗號分隔）。"""
                    if rule_str is None or str(rule_str).strip().upper() in {"N/A", "NAN", "NONE", "", "-"}:
                        return 0
                    return len([t for t in str(rule_str).split(",") if t.strip()])

                def _to_int(value):
                    try:
                        if value is None or str(value).strip().upper() in {"", "N/A", "NAN", "NONE", "-", "NAT"}:
                            return 0
                        return int(float(value))
                    except (TypeError, ValueError):
                        return 0

                def _week_label(*candidates):
                    """Return current system week label like W625."""
                    now = pd.Timestamp.now()
                    iso_week = int(now.isocalendar().week)
                    return f"W{now.year % 10}{iso_week:02d}"

                _kpi_names = ["OOB", "OOC", "OOS", "CPK<1.33"]
                _group_week_data = {}
                _week_data = {}
                for _r in data_list:
                    _group_name = (
                        _r.get("group_name")
                        or _r.get("gname")
                        or _r.get("group")
                        or "Unknown"
                    )
                    _group_name = str(_group_name)
                    _wk = _week_label(
                        _r.get("weekly_end"),
                        _r.get("weekly_start"),
                        st.session_state.get("analysis_base_date"),
                    )
                    if not _wk:
                        continue
                    if _group_name not in _group_week_data:
                        _group_week_data[_group_name] = {}
                    if _wk not in _group_week_data[_group_name]:
                        _group_week_data[_group_name][_wk] = {kpi: 0 for kpi in _kpi_names}
                    _group_week_data[_group_name][_wk]["OOB"] += _oob_count(_r.get("OOB_Rule"))
                    _group_week_data[_group_name][_wk]["OOC"] += _to_int(_r.get("ooc_cnt"))
                    _group_week_data[_group_name][_wk]["OOS"] += _to_int(_r.get("oos_cnt"))
                    _group_week_data[_group_name][_wk]["CPK<1.33"] += _to_int(_r.get("cpk_below_133_cnt"))

                    if _wk not in _week_data:
                        _week_data[_wk] = {kpi: 0 for kpi in _kpi_names}
                    _week_data[_wk]["OOB"]      += _oob_count(_r.get("OOB_Rule"))
                    _week_data[_wk]["OOC"]      += _to_int(_r.get("ooc_cnt"))
                    _week_data[_wk]["OOS"]      += _to_int(_r.get("oos_cnt"))
                    _week_data[_wk]["CPK<1.33"] += _to_int(_r.get("cpk_below_133_cnt"))

                _group_week_df = pd.DataFrame()
                if _group_week_data:
                    _sorted_group_weeks = sorted(
                        {week for group_data in _group_week_data.values() for week in group_data.keys()}
                    )
                    _group_week_rows = []
                    _group_week_index = []
                    for _group_name in _group_week_data:
                        for _kpi in _kpi_names:
                            _group_week_index.append((_group_name, _kpi))
                            _group_week_rows.append([
                                _group_week_data[_group_name].get(_wk, {}).get(_kpi, 0)
                                for _wk in _sorted_group_weeks
                            ])
                    _group_week_df = pd.DataFrame(
                        _group_week_rows,
                        index=pd.MultiIndex.from_tuples(_group_week_index, names=["Group", "KPI"]),
                        columns=_sorted_group_weeks,
                    )

                _trend_df = pd.DataFrame()
                if _week_data:
                    _sorted_weeks = sorted(_week_data.keys())
                    _trend_df = pd.DataFrame(
                        {_wk: [_week_data[_wk]["OOB"],
                                _week_data[_wk]["OOC"],
                                _week_data[_wk]["OOS"],
                                _week_data[_wk]["CPK<1.33"]]
                         for _wk in _sorted_weeks},
                        index=["OOB", "OOC", "OOS", "CPK<1.33"]
                    )

                _history_df = pd.DataFrame()
                _group_history_df = pd.DataFrame()
                _auto_summary_path = _auto_weekly_summary_path()
                _weekly_summary_path = st.session_state.get("saved_weekly_summary_path")
                if not _weekly_summary_path and os.path.exists(_auto_summary_path):
                    _weekly_summary_path = _auto_summary_path
                if _weekly_summary_path and os.path.exists(_weekly_summary_path):
                    try:
                        _history_df, _group_history_df = _read_weekly_summary_csv(_weekly_summary_path)
                    except Exception as _weekly_import_err:
                        st.warning(f"Historical Weekly Summary CSV could not be loaded: {_weekly_import_err}")

                _trend_df = _merge_weekly_summary_tables(_history_df, _trend_df)
                _group_week_df = _merge_group_weekly_summary_tables(_group_history_df, _group_week_df)
                if not _trend_df.empty:
                    def _week_sort_key(label):
                        text = str(label).strip().upper()
                        if text.startswith("W") and text[1:].isdigit():
                            return int(text[1:])
                        return 99999

                    _trend_df = _trend_df[sorted(_trend_df.columns, key=_week_sort_key)]
                if not _group_week_df.empty:
                    _group_week_df = _group_week_df[sorted(_group_week_df.columns, key=_week_sort_key)]
                    try:
                        _save_weekly_summary_csv(_trend_df, _group_week_df, _auto_summary_path)
                        st.session_state.saved_weekly_summary_path = _auto_summary_path
                    except Exception as _weekly_save_err:
                        st.warning(f"Weekly Summary CSV could not be saved: {_weekly_save_err}")
                    st.markdown("##### KPI Trend")
                    _trend_plot_df = _trend_df.T.reset_index().rename(columns={"index": "Week"})
                    fig = go.Figure()
                    _metric_colors = {
                        "OOB": "#E83F6F",
                        "OOC": "#F59E0B",
                        "OOS": "#0284C7",
                        "CPK<1.33": "#7C3AED",
                    }
                    for _metric in ["OOB", "OOC", "OOS", "CPK<1.33"]:
                        if _metric in _trend_plot_df.columns:
                            fig.add_trace(go.Scatter(
                                x=_trend_plot_df["Week"],
                                y=_trend_plot_df[_metric],
                                name=_metric,
                                mode="lines+markers",
                                line=dict(width=2, color=_metric_colors.get(_metric)),
                                marker=dict(size=6),
                            ))
                    fig.update_layout(
                        height=260,
                        margin=dict(l=20, r=12, t=8, b=24),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        hovermode="x unified",
                        template="plotly_white",
                    )
                    fig.update_xaxes(title_text=None, showgrid=False)
                    fig.update_yaxes(title_text=None, rangemode="tozero", gridcolor="rgba(0,0,0,0.08)")
                    st.plotly_chart(fig, use_container_width=True)

                    st.markdown("##### Weekly Summary")
                    _weekly_summary_display_df = _trend_df.reset_index().rename(columns={"index": "KPI"})
                    st.dataframe(
                        _weekly_summary_display_df,
                        hide_index=True,
                        use_container_width=True,
                        height=178,
                    )

                    st.markdown("##### Group Weekly Summary")
                    if not _group_week_df.empty:
                        _group_week_df = _group_week_df[sorted(_group_week_df.columns, key=_week_sort_key)]
                        _group_week_display_df = _group_week_df.reset_index()
                        _group_week_display_df.loc[
                            _group_week_display_df["Group"].duplicated(),
                            "Group",
                        ] = ""
                        st.dataframe(
                            _group_week_display_df,
                            hide_index=True,
                            use_container_width=True,
                            height=670,
                        )
                    else:
                        st.info("No group weekly summary data.")

        with oob_chart_tab:
            c_top_left, c_top_right = st.columns([1.6, 2.4], gap="small")
        
            with c_top_left:
                # st.markdown("##### Summary Table")
            
                # --- 任務摘要統計列 ---
                _mode_now = st.session_state.get("current_mode", "")
                if _mode_now == "OOB/SPC":
                    pass
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
                keep_list = ['group_name', 'gname', 'chart_name', 'cname', 'OOB_Rule', 'WE_Rule', 'group', 'Characteristics', 'characteristics', 'ooc_cnt', 'oos_cnt', 'cpk_below_133_cnt', 'abnormal_type', 'cpk', 'data_cnt']
                if _mode_now == "CPK Dashboard":
                    keep_list.extend(['cpk_l1', 'cpk_l2', 'r1', 'r2', 'cpk_violation', 'k_value', 'mean_index', 'sigma_index'])
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
                _numeric_cols = [
                    "ooc_cnt", "oos_cnt", "cpk_below_133_cnt", "data_cnt",
                    "cpk", "cpk_l1", "cpk_l2", "r1", "r2", "k_value",
                    "mean_index", "sigma_index",
                ]

                # Keep numeric columns numeric so Streamlit/Arrow can serialize them safely.
                for _col in [c for c in _numeric_cols if c in display_df.columns]:
                    display_df[_col] = pd.to_numeric(display_df[_col], errors="coerce")

                for _col in display_df.columns:
                    if _col in ['WE_Rule', 'OOB_Rule'] or _col in _numeric_cols:
                        continue
                    if display_df[_col].dtype == object:
                        display_df[_col] = display_df[_col].fillna("-").replace(_na_vals, "-")
                    else:
                        display_df[_col] = display_df[_col].where(display_df[_col].notna(), other="-")

                # --- 建立 AgGrid ---
                gb = GridOptionsBuilder.from_dataframe(display_df)
                gb.configure_selection(selection_mode="single", use_checkbox=False)
                gb.configure_default_column(
                    resizable=True,
                    wrapHeaderText=True,
                    autoHeaderHeight=True,
                    suppressSizeToFit=False,
                )
            
                col_settings = {
                    "gname": {"header_name": "Group", "minWidth": 95},
                    "cname": {"header_name": "Chart", "minWidth": 150},
                    "group": {"header_name": "Group", "minWidth": 95},
                    "group_name": {"header_name": "Group", "minWidth": 110},
                    "chart_name": {"header_name": "Chart", "minWidth": 140},
                    "Characteristics": {"header_name": "Char.", "minWidth": 85},
                    "characteristics": {"header_name": "Char.", "minWidth": 85},
                    "WE_Rule": {"header_name": "WE Rule", "minWidth": 70},
                    "OOB_Rule": {"header_name": "OOB Rule", "minWidth": 210},
                    "abnormal_type": {"header_name": "Abnormal", "minWidth": 105},
                    "cpk": {"header_name": "Cpk", "minWidth": 75},
                    "cpk_l1": {"header_name": "Cpk L1", "minWidth": 85},
                    "cpk_l2": {"header_name": "Cpk L2", "minWidth": 85},
                    "r1": {"header_name": "R1", "minWidth": 70},
                    "r2": {"header_name": "R2", "minWidth": 70},
                    "cpk_violation": {"header_name": "Violation", "minWidth": 100},
                    "k_value": {"header_name": "K Value", "minWidth": 90},
                    "ooc_cnt": {"header_name": "OOC", "minWidth": 70},
                    "oos_cnt": {"header_name": "OOS", "minWidth": 70},
                    "cpk_below_133_cnt": {"header_name": "CPK<1.33", "minWidth": 105},
                    "mean_index": {"header_name": "Mean Idx", "minWidth": 100},
                    "sigma_index": {"header_name": "Sigma Idx", "minWidth": 105},
                    "data_cnt": {"header_name": "N", "minWidth": 70}
                }
            
                for col in display_df.columns:
                    if col in col_settings:
                        gb.configure_column(col, **col_settings[col])
            
                gridOptions = gb.build()
            
                grid_response = AgGrid(
                    display_df,
                    gridOptions=gridOptions,
                    update_on=["selectionChanged"],
                    fit_columns_on_grid_load=True,
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
                        st.markdown(
                            f"""
                            <div style="height: 530px; display: flex; flex-direction: column; justify-content: flex-start;">
                                <div style="flex: 1; display: flex; align-items: center; justify-content: center;">
                                    <img
                                        src="data:image/png;base64,{item['chart_image']}"
                                        style="max-width: 100%; max-height: 100%; width: 100%; height: 100%; object-fit: contain;"
                                    />
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
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

            # ==========================================
            # Per-Row 明細表（OOC / OOS / CPK<1.33）
            # ==========================================
            if st.session_state.get("current_mode") == "OOB/SPC":
                _row_details = res.get("row_level_details") or []
                with st.expander(f"📋 Per-Row 明細（OOC / OOS / CPK<1.33）— 共 {len(_row_details)} 筆違規", expanded=False):
                    if _row_details:
                        _rd_df = pd.DataFrame(_row_details)
                        # 欄位排序與重命名
                        _col_order = ["group_name", "chart_name", "Batch_ID", "point_time", "point_val",
                                      "OOC", "OOS", "cpk", "cpk_lt_133", "chart_oob_rule"]
                        _rd_df = _rd_df[[c for c in _col_order if c in _rd_df.columns]]
                        _rd_df = _rd_df.rename(columns={
                            "group_name": "Group", "chart_name": "Chart",
                            "point_val": "Value", "cpk": "CPK",
                            "cpk_lt_133": "CPK<1.33", "chart_oob_rule": "Chart OOB Rule",
                        })
                        st.dataframe(_rd_df, hide_index=True, use_container_width=True)
                    else:
                        st.info("本次執行無 OOC / OOS / CPK<1.33 違規資料（或 input rawdata 無 cpk 欄位）。")

else:
    if st.session_state.status == "idle":
        st.markdown("<h3 style='text-align: center; color: #888; padding-top: 100px;'>點擊左上角 Settings 開始分析</h3>", unsafe_allow_html=True)
