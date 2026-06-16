"""
generate_unified_test_data.py
產生符合新 Unified Vertical 格式的模擬 rawdata CSV，共 10 張圖表。

欄位格式：
    GroupName, ChartName, USL, LSL, UCL, LCL, Target,
    Mean, Std, point_time, Batch_ID,
    Site1, Site2, ... Site17

規則：
    - Mean / Std 是該列 17 個 site 量測值的平均 / 標準差
    - 同一組 (GroupName, ChartName) 有多筆（時間序列）
    - Characteristics 由 USL/LSL 決定（程式不需填，系統自動推算）
"""

import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# 圖表設定：10 張不同的 GroupName + ChartName
# ─────────────────────────────────────────────
CHART_CONFIGS = [
    # (GroupName, ChartName, center, sigma, USL,   LSL,   UCL,   LCL,   Target, n_rows, pattern)
    ("PKG_A",  "Thickness",     5.0,  0.30,  6.0,   4.0,   5.9,   4.1,   5.0,   80, "normal"),
    ("PKG_A",  "Warpage",       2.0,  0.20,  2.8,   None,  2.7,   None,  2.0,   80, "trend_up"),
    ("PKG_B",  "Bond_Strength", 30.0, 1.50,  35.0,  25.0,  33.0,  27.0,  30.0,  100, "normal"),
    ("PKG_B",  "Void_Rate",     1.5,  0.10,  2.0,   None,  1.9,   None,  1.5,   60, "skew_right"),
    ("PKG_C",  "Resistance",    100.0,3.00,  109.0, 91.0,  106.0, 94.0,  100.0, 90, "shift"),
    ("PKG_C",  "Leakage",       0.05, 0.005, None,  None,  0.065, None,  0.05,  70, "normal"),
    ("TEST_D", "Yield_Loss",    500.0,20.0,  560.0, 440.0, 540.0, 460.0, 500.0, 100, "ooc_spike"),
    ("TEST_D", "Vt_Shift",      -0.2, 0.02,  0.0,  -0.4,  -0.05, -0.35, -0.2,  80, "normal"),
    ("TEST_E", "Contact_Res",   5.5,  0.25,  6.5,   4.5,   6.2,   4.8,   5.5,   90, "bimodal"),
    ("TEST_E", "Solder_Height", 0.3,  0.02,  0.38,  0.22,  0.36,  0.24,  0.30,  80, "trend_down"),
]

N_SITES = 17
START_DATE = datetime(2025, 1, 1, 8, 0, 0)


def _make_site_values(center: float, sigma: float, pattern: str,
                      row_idx: int, rng: np.random.Generator) -> np.ndarray:
    """根據 pattern 為單列產生 N_SITES 個量測值。"""
    base = center + (row_idx * 0.001 if pattern == "trend_up"
                     else -row_idx * 0.001 if pattern == "trend_down"
                     else 0)

    if pattern in ("normal", "trend_up", "trend_down"):
        vals = rng.normal(base, sigma, N_SITES)

    elif pattern == "skew_right":
        # gamma 分佈偏右
        k = 2.0
        vals = rng.gamma(k, sigma / np.sqrt(k), N_SITES) + base - sigma * np.sqrt(k)

    elif pattern == "shift":
        # 前半段正常，後半段均值上移 2σ
        shift = 2 * sigma if row_idx >= 45 else 0
        vals = rng.normal(base + shift, sigma, N_SITES)

    elif pattern == "ooc_spike":
        # 每 ~15 筆有一個 spike
        vals = rng.normal(base, sigma, N_SITES)
        if row_idx % 15 == 0:
            vals += rng.choice([-1, 1]) * rng.uniform(3, 4) * sigma

    elif pattern == "bimodal":
        # 一半高、一半低
        low  = rng.normal(base - 1.5 * sigma, sigma * 0.5, N_SITES // 2)
        high = rng.normal(base + 1.5 * sigma, sigma * 0.5, N_SITES - N_SITES // 2)
        vals = np.concatenate([low, high])
        rng.shuffle(vals)

    else:
        vals = rng.normal(base, sigma, N_SITES)

    return vals


def _calc_cpk_row(mean: float, std: float, usl, lsl) -> float | None:
    """計算單列 Cpk（依 USL/LSL 存在與否自動判斷 Nominal/Smaller/Bigger）。"""
    if std is None or std <= 0:
        return None
    cpu = (usl - mean) / (3 * std) if usl is not None else None
    cpl = (mean - lsl) / (3 * std) if lsl is not None else None
    if cpu is not None and cpl is not None:
        return round(min(cpu, cpl), 4)
    if cpu is not None:
        return round(cpu, 4)
    if cpl is not None:
        return round(cpl, 4)
    return None


def generate_unified_rawdata(output_path: str = "input/unified_rawdata_sample.csv") -> pd.DataFrame:
    rng = np.random.default_rng(seed=42)
    rows = []

    for (gname, cname, center, sigma, usl, lsl, ucl, lcl, target, n_rows, pattern) in CHART_CONFIGS:
        for i in range(n_rows):
            ts = START_DATE + timedelta(hours=i * 12)
            batch_id = f"LOT{i+1:04d}"
            site_vals = _make_site_values(center, sigma, pattern, i, rng)
            mean_val  = float(np.mean(site_vals))
            std_val   = float(np.std(site_vals, ddof=1)) if len(site_vals) > 1 else 0.0

            cpk_val = _calc_cpk_row(mean_val, std_val, usl, lsl)

            row = {
                "GroupName":  gname,
                "ChartName":  cname,
                "USL":        usl,
                "LSL":        lsl,
                "UCL":        ucl,
                "LCL":        lcl,
                "Target":     target,
                "Mean":       round(mean_val, 4),
                "Std":        round(std_val,  4),
                "point_time": ts.strftime("%Y/%m/%d %H:%M"),
                "Batch_ID":   batch_id,
                "cpk":        cpk_val,
            }
            for s_idx, val in enumerate(site_vals, start=1):
                row[f"Site{s_idx}"] = round(float(val), 4)

            rows.append(row)

    df = pd.DataFrame(rows)
    os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else ".", exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8-sig")

    print("Generated Unified Vertical test data")
    print(f"   路徑  : {os.path.abspath(output_path)}")
    print(f"   列數  : {len(df):,}  ({len(CHART_CONFIGS)} 張圖表)")
    print(f"   欄位數: {len(df.columns)}  (GroupName/ChartName/spec欄 + Mean/Std + Site1~Site{N_SITES} + point_time/Batch_ID)")
    print()
    print(df.groupby(["GroupName", "ChartName"]).size().rename("rows").to_string())
    return df


def generate_historical_weekly_summary(output_path: str = "input/historical_weekly_summary_sample.csv") -> pd.DataFrame:
    """Generate a sample historical Weekly Summary CSV for UI import testing."""
    rng = np.random.default_rng(seed=2026)
    weeks = ["W103", "W110", "W117", "W124", "W131", "W207", "W214", "W221"]

    values = {
        "Metric": ["OOB", "OOC", "OOS", "CPK<1.33"],
        "W103": [2, 5, 1, 3],
        "W110": [3, 4, 0, 2],
        "W117": [1, 6, 2, 4],
        "W124": [4, 3, 1, 3],
        "W131": [2, 5, 0, 2],
        "W207": [5, 7, 2, 5],
        "W214": [3, 4, 1, 4],
        "W221": [4, 6, 1, 3],
    }
    df = pd.DataFrame(values)

    # Add small deterministic variation while keeping readable demo numbers.
    for week in weeks:
        df[week] = (df[week] + rng.integers(-1, 2, size=len(df))).clip(lower=0)

    os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else ".", exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8-sig")

    print("Generated Historical Weekly Summary test CSV")
    print(f"   路徑  : {os.path.abspath(output_path)}")
    print(f"   欄位  : Metric + {len(weeks)} weekly columns")
    return df


if __name__ == "__main__":
    generate_unified_rawdata()
    generate_historical_weekly_summary()
