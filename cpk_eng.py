"""
cpk_eng.py — SPC CPK Dashboard 計算與繪圖模組
包含 CPK 計算、時間窗口分析、SPC 圖表產生及 Excel 匯出等純函式。
"""
import os
import math
import base64
import tempfile
from io import BytesIO
from datetime import date, datetime
from typing import Optional, List, Any

import pandas as pd
import numpy as np
from scipy import stats

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.transforms as mtransforms
import matplotlib.gridspec as gridspec

import xlsxwriter


# ==========================================
# CPK 計算函式
# ==========================================

def calculate_cpk_dashboard(raw_df: pd.DataFrame, chart_info: dict) -> dict:
    if raw_df.empty: return {'Cpk': None}
    mean = raw_df['point_val'].mean()
    std = raw_df['point_val'].std()
    # 防呆：標準差接近 0 時 Cpk 會爆掉，直接回傳 None
    if pd.isna(std) or std < 1e-6:
        return {'Cpk': None}
    characteristic = chart_info.get('Characteristics', '')
    usl = chart_info.get('USL', None)
    lsl = chart_info.get('LSL', None)
    
    cpk = None
    if characteristic == 'Nominal':
        if usl is not None and lsl is not None:
            cpk = min((usl - mean) / (3 * std), (mean - lsl) / (3 * std))
    elif characteristic in ['Smaller', 'Sigma']:
        if usl is not None: cpk = (usl - mean) / (3 * std)
    elif characteristic == 'Bigger':
        if lsl is not None: cpk = (mean - lsl) / (3 * std)
    
    return {'Cpk': round(cpk, 3) if cpk is not None else None}


def compute_cpk_windows(raw_df: pd.DataFrame, chart_info: dict, end_time: pd.Timestamp) -> dict:
    result = {
        'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None,
        'mean_current': None, 'sigma_current': None, 'mean_last_month': None, 
        'sigma_last_month': None, 'mean_last2_month': None, 'sigma_last2_month': None,
    }

    if raw_df is None or raw_df.empty: return result

    if 'point_time' not in raw_df.columns:
        result['Cpk'] = calculate_cpk_dashboard(raw_df, chart_info)['Cpk']
        result['mean_current'] = raw_df['point_val'].mean()
        result['sigma_current'] = raw_df['point_val'].std()
        return result

    df = raw_df.copy()
    try:
        df['point_time'] = pd.to_datetime(df['point_time'])
    except Exception:
        result['Cpk'] = calculate_cpk_dashboard(df, chart_info)['Cpk']
        result['mean_current'] = df['point_val'].mean()
        result['sigma_current'] = df['point_val'].std()
        return result

    df = df[df['point_time'] <= end_time]
    if df.empty: return result

    start1 = end_time - pd.DateOffset(months=1)
    start2 = end_time - pd.DateOffset(months=2)
    start3 = end_time - pd.DateOffset(months=3)

    mask1 = (df['point_time'] > start1) & (df['point_time'] <= end_time)
    mask2 = (df['point_time'] > start2) & (df['point_time'] <= start1)
    mask3 = (df['point_time'] > start3) & (df['point_time'] <= start2)

    if mask1.any():
        seg = df[mask1]
        result['Cpk'] = calculate_cpk_dashboard(seg, chart_info)['Cpk']
        result['mean_current'] = seg['point_val'].mean()
        result['sigma_current'] = seg['point_val'].std()
    else:
        # 當月無資料時，以最新進點的時間為基準，往前取 1 個月作為 fallback 視窗
        df_sorted = df.sort_values('point_time')
        if not df_sorted.empty:
            latest_time = df_sorted['point_time'].iloc[-1]
            fallback_start = latest_time - pd.DateOffset(months=1)
            fallback = df_sorted[(df_sorted['point_time'] > fallback_start) & (df_sorted['point_time'] <= latest_time)]
            if fallback.empty:
                fallback = df_sorted  # 若仍為空則使用全部資料
            result['Cpk'] = calculate_cpk_dashboard(fallback, chart_info)['Cpk']
            result['mean_current'] = fallback['point_val'].mean()
            result['sigma_current'] = fallback['point_val'].std()
    if mask2.any():
        seg = df[mask2]
        result['Cpk_last_month'] = calculate_cpk_dashboard(seg, chart_info)['Cpk']
        result['mean_last_month'] = seg['point_val'].mean()
        result['sigma_last_month'] = seg['point_val'].std()
    if mask3.any():
        seg = df[mask3]
        result['Cpk_last2_month'] = calculate_cpk_dashboard(seg, chart_info)['Cpk']
        result['mean_last2_month'] = seg['point_val'].mean()
        result['sigma_last2_month'] = seg['point_val'].std()

    for k, v in list(result.items()):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            result[k] = None
    return result


def _get_target_value(chart_info) -> Optional[float]:
    for key in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
        if key in chart_info and pd.notna(chart_info.get(key)):
            return chart_info[key]
    return None


def _calculate_period_statistics(raw_df: pd.DataFrame, end_date: date, custom_mode: bool, start_date: Optional[date] = None) -> dict:
    stats_dict = {
        'mean_current': None, 'sigma_current': None, 'mean_last_month': None, 'sigma_last_month': None,
        'mean_last2_month': None, 'sigma_last2_month': None, 'mean_all': None, 'sigma_all': None
    }
    if raw_df is None or raw_df.empty: return stats_dict
    
    stats_dict['mean_all'] = raw_df['point_val'].mean()
    stats_dict['sigma_all'] = raw_df['point_val'].std()
    
    if 'point_time' not in raw_df.columns: return stats_dict
    
    try:
        df = raw_df.copy()
        df['point_time'] = pd.to_datetime(df['point_time'])
        end_time = pd.to_datetime(end_date)
        
        if custom_mode and start_date:
            start_time = pd.to_datetime(start_date)
            end_time = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
            current_df = df[(df['point_time'] >= start_time) & (df['point_time'] <= end_time)]
            if not current_df.empty:
                stats_dict['mean_current'] = current_df['point_val'].mean()
                stats_dict['sigma_current'] = current_df['point_val'].std()
        else:
            start1 = end_time - pd.DateOffset(months=1)
            start2 = end_time - pd.DateOffset(months=2)
            start3 = end_time - pd.DateOffset(months=3)
            
            current_df = df[(df['point_time'] > start1) & (df['point_time'] <= end_time)]
            if not current_df.empty:
                stats_dict['mean_current'] = current_df['point_val'].mean()
                stats_dict['sigma_current'] = current_df['point_val'].std()
            
            last_month_df = df[(df['point_time'] > start2) & (df['point_time'] <= start1)]
            if not last_month_df.empty:
                stats_dict['mean_last_month'] = last_month_df['point_val'].mean()
                stats_dict['sigma_last_month'] = last_month_df['point_val'].std()
            
            last2_month_df = df[(df['point_time'] > start3) & (df['point_time'] <= start2)]
            if not last2_month_df.empty:
                stats_dict['mean_last2_month'] = last2_month_df['point_val'].mean()
                stats_dict['sigma_last2_month'] = last2_month_df['point_val'].std()
    except Exception:
        pass
    return stats_dict


def _compute_cpk_custom_range(raw_df: pd.DataFrame, chart_info: dict, start_time: pd.Timestamp, end_time: pd.Timestamp) -> dict:
    result = {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
    if raw_df is None or raw_df.empty: return result
    if 'point_time' not in raw_df.columns:
        result['Cpk'] = calculate_cpk_dashboard(raw_df, chart_info)['Cpk']
        return result
    
    df = raw_df.copy()
    df['point_time'] = pd.to_datetime(df['point_time'])
    filtered_df = df[(df['point_time'] >= start_time) & (df['point_time'] <= end_time)]
    if not filtered_df.empty:
        result['Cpk'] = calculate_cpk_dashboard(filtered_df, chart_info)['Cpk']
    return result


def _calculate_k_value(raw_df: pd.DataFrame, chart_info: dict, start_date: date, end_date: date, custom_mode: bool) -> Optional[float]:
    try:
        usl = chart_info.get('USL')
        lsl = chart_info.get('LSL')
        target = _get_target_value(chart_info)
        if target is None or usl is None or lsl is None: return None
        
        if custom_mode and 'point_time' in raw_df.columns:
            start_ts = pd.to_datetime(start_date)
            end_ts = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
            filtered_df = raw_df[(pd.to_datetime(raw_df['point_time']) >= start_ts) & 
                               (pd.to_datetime(raw_df['point_time']) <= end_ts)]
            mean_val = filtered_df['point_val'].mean() if not filtered_df.empty else raw_df['point_val'].mean()
        else:
            mean_val = raw_df['point_val'].mean()
        
        rng = (usl - lsl) / 2 if (usl - lsl) != 0 else None
        if mean_val is not None and rng:
            return round(abs(mean_val - target) / rng, 3)
    except Exception: pass
    return None


# ==========================================
# SPC 圖表繪圖函式
# ==========================================

def _detect_tool_col(df: pd.DataFrame) -> Optional[str]:
    """偵測 DataFrame 中代表機台/工具的欄位（支援 EQP_id、ByTool 等多種命名）。
    回傳第一個有效欄位名稱，無則回傳 None。"""
    # 👇 這裡加入了 'Matching' 與 'matching_group'
    for col in ['EQP_id', 'ByTool', 'Tool', 'tool_id', 'TOOL_ID', 'Matching', 'matching_group']:
        if col in df.columns:
            valid = df[col].dropna().astype(str).str.strip()
            valid = valid[valid != '']
            if valid.nunique() > 1:
                return col
    return None


def _draw_main_spc_chart_api(ax, plot_df, chart_info, start_date, end_date, custom_mode):
    y = plot_df['point_val'].values
    x = range(1, len(y) + 1)
    
    if 'point_time' in plot_df.columns:
        try:
            plot_df = plot_df.sort_values('point_time').reset_index(drop=True)
            y = plot_df['point_val'].values
            
            if not plot_df.empty:
                times = pd.to_datetime(plot_df['point_time']).to_numpy()
                tmin, tmax = times.min(), times.max()
                
                if custom_mode and start_date and end_date:
                    start_time = pd.to_datetime(start_date)
                    end_time = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    windows = [(start_time, end_time, 'Custom', '#dbeafe')]
                else:
                    # 用當天最後一毫秒，確保當天所有時間點都被底色涵蓋（原本用 00:00:00 會漏掉同天的下午資料點）
                    end_sel = (pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)) if end_date else pd.Timestamp(tmax)
                    if end_sel > pd.Timestamp(tmax): end_sel = pd.Timestamp(tmax)
                    start1 = end_sel - pd.DateOffset(months=1)
                    start2 = end_sel - pd.DateOffset(months=2)
                    start3 = end_sel - pd.DateOffset(months=3)
                    windows = [
                        (start1, end_sel, 'L0', '#dbeafe'),
                        (start2, start1, 'L1', '#fef9c3'),
                        (start3, start2, 'L2', '#ede9fe'),
                    ]
                
                text_trans = mtransforms.blended_transform_factory(ax.transData, ax.transAxes)
                n = len(times)
                def t2ix_left(t): return float(np.searchsorted(times, np.datetime64(t), side='left')) + 0.5
                def t2ix_right(t): return float(np.searchsorted(times, np.datetime64(t), side='right')) + 0.5
                
                x_min, x_max = 0.5, n + 0.5
                for s, e, lab, col in windows:
                    s_clip = max(pd.Timestamp(s), pd.Timestamp(tmin))
                    e_clip = min(pd.Timestamp(e), pd.Timestamp(tmax))
                    if e_clip <= s_clip: continue
                    xl = max(x_min, t2ix_left(s_clip))
                    xr = min(x_max, t2ix_right(e_clip))
                    if xr <= xl: continue
                    ax.axvspan(xl, xr, color=col, alpha=0.25, zorder=0)
                    ax.text((xl + xr) / 2.0, 1.04, lab, transform=text_trans, ha='center', va='top', fontsize=8, color='#374151', alpha=0.9)
        except Exception: pass
    
    ax.plot(x, y, linestyle='-', marker='o', color='#2563eb', markersize=4, linewidth=1.0)
    
    usl = chart_info.get('USL', None)
    lsl = chart_info.get('LSL', None)
    target = _get_target_value(chart_info)
    mean_val = float(np.mean(y)) if len(y) else None
    
    if usl is not None:
        ax.scatter([xi for xi, yi in zip(x, y) if yi > usl], [yi for yi in y if yi > usl], color='#dc2626', s=25, zorder=5)
    if lsl is not None:
        ax.scatter([xi for xi, yi in zip(x, y) if yi < lsl], [yi for yi in y if yi < lsl], color='#dc2626', marker='s', s=25, zorder=5)
    
    extra_vals = [v for v in [usl, lsl, target, mean_val] if v is not None and not (isinstance(v, float) and np.isnan(v))]
    if len(y) > 0:
        ymin_sel, ymax_sel = float(np.min(y)), float(np.max(y))
    else:
        ymin_sel, ymax_sel = (0.0, 1.0)
    if extra_vals:
        ymin_sel, ymax_sel = min(ymin_sel, min(extra_vals)), max(ymax_sel, max(extra_vals))
    rng = ymax_sel - ymin_sel
    margin = 0.05 * rng if rng > 0 else 1.0
    ax.set_ylim(ymin_sel - margin, ymax_sel + margin)
    
    trans = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)
    def segment_with_label(val, name, color, va='center'):
        if val is None or (isinstance(val, float) and np.isnan(val)): return
        ax.plot([0.0, 0.96], [val, val], transform=trans, color=color, linestyle='--', linewidth=1.0)
        ax.text(0.96, val, name, transform=trans, color=color, va=va, ha='left', fontsize=8)
    
    segment_with_label(usl, 'USL', '#ef4444', va='center')
    segment_with_label(lsl, 'LSL', '#ef4444', va='center')
    segment_with_label(target, 'Target', '#f59e0b', va='center')
    segment_with_label(mean_val, 'Mean', '#16a34a', va='center')
    
    if 'point_time' in plot_df.columns and not plot_df.empty:
        times = plot_df['point_time'].tolist()
        total = len(times)
        if total <= 8:
            tick_idx = list(range(1, total + 1))
        else:
            step = max(1, total // 6)
            tick_idx = list(range(1, total + 1, step))
            if tick_idx[-1] != total: tick_idx.append(total)
        labels = [times[i-1].strftime('%Y-%m-%d') for i in tick_idx]
        ax.set_xticks(tick_idx)
        ax.set_xticklabels(labels, rotation=90, ha='right', fontsize=8)
    
    ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.5)


def _draw_box_plot_api(ax, plot_df, chart_info):
    if plot_df.empty:
        ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
        return
    tool_col = _detect_tool_col(plot_df)
    if tool_col is not None:
        grouped = plot_df.groupby(tool_col)
        box_data, labels = [], []
        for grp_id in sorted(grouped.groups.keys(), key=str):
            group_data = grouped.get_group(grp_id)['point_val'].values
            if len(group_data) > 0:
                box_data.append(group_data)
                labels.append(str(grp_id))
        
        if len(box_data) == 0:
            ax.text(0.5, 0.5, "No Valid Data", ha='center', va='center', transform=ax.transAxes)
            return
        
        box_plot = ax.boxplot(box_data, patch_artist=True, notch=False)
        colors = ['#87CEEB', '#98FB98', '#FFB6C1', '#F0E68C', '#DDA0DD', '#F5DEB3', '#B0E0E6']
        for i, patch in enumerate(box_plot['boxes']):
            patch.set_facecolor(colors[i % len(colors)])
            patch.set_alpha(0.8)
        ax.set_xticklabels(labels, rotation=0, ha='center', fontsize=9)
        ax.set_xlabel('')
    else:
        y = plot_df['point_val'].values
        if len(y) == 0:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
            return
        box_plot = ax.boxplot(y, patch_artist=True, notch=False)
        box_plot['boxes'][0].set_facecolor('#87CEEB')
        box_plot['boxes'][0].set_alpha(0.8)
        ax.set_xticks([])
        ax.set_xlabel('')


def _draw_qq_plot_api(ax, plot_df, chart_info):
    if plot_df.empty or plot_df['point_val'].dropna().empty:
        ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
        return
    tool_col = _detect_tool_col(plot_df)
    palette = plt.cm.tab10(np.linspace(0, 1, 10))
    if tool_col is not None:
        # 每台機台獨立畫一條常態機率線
        groups = sorted(plot_df[tool_col].dropna().unique(), key=str)
        for i, grp in enumerate(groups):
            y = plot_df[plot_df[tool_col] == grp]['point_val'].dropna().values
            if len(y) < 3:
                continue
            try:
                (osm, osr), (slope, intercept, r) = stats.probplot(y, dist="norm", plot=None)
                color = palette[i % len(palette)]
                ax.scatter(osm, osr, alpha=0.6, color=color, s=15, zorder=3)
                line_x = np.array([osm.min(), osm.max()])
                ax.plot(line_x, slope * line_x + intercept, '-', color=color,
                        linewidth=1.5, alpha=0.9, label=f'{grp} R²={r**2:.3f}')
            except Exception:
                continue
        if ax.get_legend_handles_labels()[0]:
            ax.legend(fontsize=6, loc='lower right', ncol=1)
    else:
        y = plot_df['point_val'].dropna().values
        if len(y) == 0:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax.transAxes)
            return
        try:
            (osm, osr), (slope, intercept, r) = stats.probplot(y, dist="norm", plot=None)
            ax.scatter(osm, osr, alpha=0.7, color='blue', s=20)
            line_x = np.array([osm.min(), osm.max()])
            ax.plot(line_x, slope * line_x + intercept, 'r-', linewidth=1.5, alpha=0.8,
                    label=f'R²={r**2:.3f}')
            ax.legend(fontsize=7, loc='lower right')
        except Exception as e:
            ax.text(0.5, 0.5, f"Calculation Error:\n{str(e)}",
                    ha='center', va='center', transform=ax.transAxes, fontsize=8)
    ax.set_xlabel('Theoretical Quantiles', fontsize=8)
    ax.set_ylabel('Sample Quantiles', fontsize=8)
    ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.3)
    ax.tick_params(axis='both', which='major', labelsize=8)


def generate_spc_chart_base64(raw_df: pd.DataFrame, chart_info: dict, start_date: Optional[date] = None, end_date: Optional[date] = None, custom_mode: bool = False, metrics: Optional[dict] = None) -> str:
    fig = plt.figure(figsize=(12, 6))
    gs = gridspec.GridSpec(2, 2, width_ratios=[3, 1], height_ratios=[1, 1], hspace=0.3, wspace=0.25)
    ax_main = fig.add_subplot(gs[:, 0])
    ax_box = fig.add_subplot(gs[0, 1])
    ax_qq = fig.add_subplot(gs[1, 1])
    
    if raw_df is None or raw_df.empty:
        ax_main.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_main.transAxes)
        ax_box.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_box.transAxes)
        ax_qq.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_qq.transAxes)
    else:
        plot_df = raw_df.copy()
        if 'point_time' in plot_df.columns:
            try:
                plot_df['point_time'] = pd.to_datetime(plot_df['point_time'])
                if custom_mode and start_date and end_date:
                    # custom 模式：用使用者指定的日期範圍
                    start_ts = pd.to_datetime(start_date)
                    end_ts = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    filtered = plot_df[(plot_df['point_time'] >= start_ts) & (plot_df['point_time'] <= end_ts)]
                    if not filtered.empty:
                        plot_df = filtered
                elif end_date:
                    # 非 custom 模式：以 min(end_date, tmax) 為基準倒推 3M
                    # 確保 L0/L1/L2 三個色塊都落在繪圖資料範圍內
                    end_ts = pd.to_datetime(end_date) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    end_actual = min(end_ts, plot_df['point_time'].max())
                    start_actual = end_actual - pd.DateOffset(months=3)
                    filtered = plot_df[plot_df['point_time'] >= start_actual]
                    if not filtered.empty:
                        plot_df = filtered
            except Exception:
                pass
        
        if not plot_df.empty:
            _draw_main_spc_chart_api(ax_main, plot_df, chart_info, start_date, end_date, custom_mode)
            _draw_box_plot_api(ax_box, plot_df, chart_info)
            _draw_qq_plot_api(ax_qq, plot_df, chart_info)
    
    group_name = chart_info.get('GroupName', '')
    chart_name = chart_info.get('ChartName', '')
    characteristics = chart_info.get('Characteristics', '')

    # --- 組合多行 title ---
    def _sf(v):
        try:
            f = float(v)
            return None if (f != f) else f  # NaN guard
        except (TypeError, ValueError):
            return None

    _line1 = f"{group_name} / {chart_name}" + (f" / {characteristics}" if characteristics else "")

    _m = metrics or {}
    _cpk   = _sf(_m.get('cpk'))
    _cpk_l1 = _sf(_m.get('cpk_l1'))
    _cpk_l2 = _sf(_m.get('cpk_l2'))
    _r1    = _sf(_m.get('r1'))
    _r2    = _sf(_m.get('r2'))
    _viol  = str(_m.get('violation', '') or '').strip()

    _line2_parts = []
    if _cpk    is not None: _line2_parts.append(f"Cpk(L0): {_cpk:.3f}")
    if _cpk_l1 is not None: _line2_parts.append(f"Cpk(L1): {_cpk_l1:.3f}")
    if _cpk_l2 is not None: _line2_parts.append(f"Cpk(L2): {_cpk_l2:.3f}")

    _line3_parts = []
    if _r1 is not None: _line3_parts.append(f"R1: {_r1:.1f}%")
    if _r2 is not None: _line3_parts.append(f"R2: {_r2:.1f}%")
    if _viol and _viol not in ['-', 'nan', 'None']: _line3_parts.append(f"⚠ {_viol}")

    _has_violation = bool(_line3_parts and _viol and _viol not in ['-', 'nan', 'None'])

    if _line2_parts or _line3_parts:
        # 用 ax.text + transform 分行控制顏色
        ax_main.set_title("")
        _y = 1.0
        ax_main.text(0, _y + 0.115, _line1,
                     transform=ax_main.transAxes, fontsize=12, fontweight='bold',
                     va='bottom', ha='left', clip_on=False)
        if _line2_parts:
            ax_main.text(0, _y + 0.068, "  |  ".join(_line2_parts),
                         transform=ax_main.transAxes, fontsize=10, color='#444444',
                         va='bottom', ha='left', clip_on=False)
        if _line3_parts:
            ax_main.text(0, _y + 0.022, "  |  ".join(_line3_parts),
                         transform=ax_main.transAxes, fontsize=10,
                         color='#c0392b' if _has_violation else '#444444',
                         va='bottom', ha='left', clip_on=False)
        fig.subplots_adjust(top=0.77)
    else:
        ax_main.set_title(_line1, pad=18, fontsize=12)
    
    _tc = _detect_tool_col(raw_df) if raw_df is not None else None
    ax_box.set_title(f"Box Plot (by {_tc})" if _tc else "Box Plot", fontsize=10)
    ax_qq.set_title("Q-Q Plot", fontsize=10)
    
    fig.tight_layout()
    buffer = BytesIO()
    fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
    plt.close(fig)
    return image_base64


# ==========================================
# Excel 匯出函式
# ==========================================

def _export_spc_cpk_to_excel(chart_results: List[Any], summary: dict, start_date: date, end_date: date) -> str:
    try:
        temp_dir = "temp_uploads"
        os.makedirs(temp_dir, exist_ok=True)
        output_path = os.path.join(temp_dir, f"spc_cpk_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        excel_data, chart_images = [], []
        
        for chart in chart_results:
            if chart.chart_image:
                temp_img_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                temp_img_file.close()
                try:
                    img_data = base64.b64decode(chart.chart_image)
                    with open(temp_img_file.name, 'wb') as f:
                        f.write(img_data)
                    chart_images.append(temp_img_file.name)
                except Exception:
                    chart_images.append(None)
            else:
                chart_images.append(None)
            
            excel_data.append({
                'ChartImage': '', 
                'ChartKey': f"{chart.group_name}@{chart.chart_name}@{chart.characteristics}",
                'GroupName': chart.group_name,
                'ChartName': chart.chart_name,
                'Characteristics': chart.characteristics,
                'USL': chart.usl, 'LSL': chart.lsl, 'Target': chart.target, 'K': chart.metrics.k_value,
                'Cpk_Curr': chart.metrics.cpk, 'Cpk_L1': chart.metrics.cpk_l1, 'Cpk_L2': chart.metrics.cpk_l2,
                'Custom_Cpk': chart.metrics.custom_cpk, 'R1(%)': chart.metrics.r1, 'R2(%)': chart.metrics.r2,
                'Mean_Curr': chart.mean_current, 'Sigma_CurrentMonth': chart.sigma_current,
                'Mean_LastMonth': chart.mean_last_month, 'Sigma_LastMonth': chart.sigma_last_month,
                'Mean_Last2Month': chart.mean_last2_month, 'Sigma_Last2Month': chart.sigma_last2_month,
                'Mean_All': chart.mean_all, 'Sigma_All': chart.sigma_all
            })
        
        df = pd.DataFrame(excel_data)
        columns = ['ChartImage'] + [c for c in df.columns if c != 'ChartImage']
        
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet()
        worksheet.set_column(0, 0, 100) 
        for i in range(1, len(columns)):
            worksheet.set_column(i, i, 15) 
        
        bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        
        for col_idx, col_name in enumerate(columns):
            worksheet.write(0, col_idx, col_name, bold)
        
        for row_idx, (row_data, img_path) in enumerate(zip(df.to_dict('records'), chart_images), 1):
            if img_path and os.path.exists(img_path):
                worksheet.set_row(row_idx, 200)
                worksheet.insert_image(row_idx, 0, img_path, {'x_scale': 0.6, 'y_scale': 0.4, 'object_position': 1, 'y_offset': 10})
            for col_idx, col_name in enumerate(columns[1:], 1):
                val = row_data.get(col_name, '')
                if val is None: val = ''
                elif isinstance(val, float) and (math.isnan(val) or math.isinf(val)): val = 'N/A'
                worksheet.write(row_idx, col_idx, val, cell_format)
        
        workbook.close()
        for img_path in chart_images:
            if img_path and os.path.exists(img_path):
                try: os.unlink(img_path)
                except: pass
        return output_path
    except Exception as e:
        print(f"Excel export failed: {e}")
        return None
