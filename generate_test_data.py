import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from datetime import datetime, timedelta
import random

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def generate_pattern_data(pattern, n_samples, base_value=10, sigma=1, decimals=2, n_categories=None):
    """
    根據 pattern 生成對應的測試數據
    decimals 參數：控制小數點位數
    n_categories 參數：Attribute pattern 的離散類別數量 (None 表示隨機 2~8)
    """
    np.random.seed(42 + hash(pattern) % 1000 + n_samples)
    
    n_samples = max(1, int(n_samples))
    
    if pattern == "Normal":
        data = np.random.normal(base_value, sigma, n_samples)
        
    elif pattern == "Skew-Right":
        data = np.random.gamma(2, sigma, n_samples) + base_value - 2*sigma
        
    elif pattern == "Skew-Left":
        data = -(np.random.gamma(2, sigma, n_samples) - 2*sigma) + base_value
        
    elif pattern == "Bimodal":
        n1 = n_samples // 2
        n2 = n_samples - n1
        data1 = np.random.normal(base_value - 2*sigma, sigma*0.5, n1)
        data2 = np.random.normal(base_value + 2*sigma, sigma*0.5, n2)
        data = np.concatenate([data1, data2])
        
    elif pattern == "Attribute":
        # 離散型數據 - 類別數量可變
        if n_categories is None:
            n_categories = np.random.randint(2, 9)  # 隨機 2~8 種
        n_categories = max(2, int(n_categories))
        # 在 base_value ± 2*sigma 範圍內均勻分配類別值
        step = (4 * sigma) / (n_categories - 1)
        categories = list({round(base_value - 2*sigma + k * step, decimals) for k in range(n_categories)})
        categories.sort()
        n_categories = len(categories)  # 去重後重新計算
        # 使用 Dirichlet 產生不均勻隨機權重
        raw_weights = np.random.dirichlet(np.ones(n_categories) * 0.7)
        data = np.random.choice(categories, n_samples, p=raw_weights)
        
    elif pattern == "Constant":
        data = np.full(n_samples, base_value)
        
    elif pattern == "Near Constant":
        n_constant = int(n_samples * 0.95)
        n_variant = n_samples - n_constant
        data_constant = np.full(n_constant, base_value)
        step = 10**(-int(decimals))
        data_variant = np.full(n_variant, base_value + step) 
        data = np.concatenate([data_constant, data_variant])
        
    elif pattern == "Step":
        n_steps = random.randint(3, 6)
        step_size = n_samples // n_steps
        data = []
        for i in range(n_steps):
            level = base_value + (i - n_steps//2) * sigma * 0.8
            step_data = np.random.normal(level, sigma * 0.15, step_size)
            data.extend(step_data)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.random.normal(base_value, sigma * 0.15, remaining))
        data = np.array(data)
        
    elif pattern == "Step-Up":
        n_steps = random.randint(3, 5)
        step_size = n_samples // n_steps
        data = []
        for i in range(n_steps):
            level = base_value + i * sigma * 0.6
            step_data = np.random.normal(level, sigma * 0.15, step_size)
            data.extend(step_data)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.random.normal(base_value + (n_steps-1) * sigma * 0.6, sigma * 0.15, remaining))
        data = np.array(data)
        
    elif pattern == "Step-Down":
        n_steps = random.randint(3, 5)
        step_size = n_samples // n_steps
        data = []
        for i in range(n_steps):
            level = base_value - i * sigma * 0.6
            step_data = np.random.normal(level, sigma * 0.15, step_size)
            data.extend(step_data)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.random.normal(base_value - (n_steps-1) * sigma * 0.6, sigma * 0.15, remaining))
        data = np.array(data)
        
    elif pattern == "Cyclic":
        t = np.linspace(0, 4*np.pi, n_samples)
        amplitude = sigma * 2
        data = base_value + amplitude * np.sin(t) + np.random.normal(0, sigma*0.2, n_samples)
        
    elif pattern == "Trending-Up":
        slope = sigma * 2 / n_samples
        trend = np.arange(n_samples) * slope
        data = base_value - sigma + trend + np.random.normal(0, sigma*0.3, n_samples)
        
    elif pattern == "Trending-Down":
        slope = sigma * 2 / n_samples
        trend = np.arange(n_samples) * slope
        data = base_value + sigma - trend + np.random.normal(0, sigma*0.3, n_samples)
        
    elif pattern == "Outliers":
        data = np.random.normal(base_value, sigma*0.5, n_samples)
        n_outliers = max(1, min(int(n_samples * random.uniform(0.05, 0.1)), n_samples))
        outlier_indices = np.random.choice(n_samples, n_outliers, replace=False)
        outlier_direction = np.random.choice([-1, 1], n_outliers)
        data[outlier_indices] += outlier_direction * sigma * random.uniform(4, 6)
        
    elif pattern == "Multimodal":
        n_modes = random.randint(3, 4)
        data = []
        samples_per_mode = n_samples // n_modes
        for i in range(n_modes):
            center = base_value + (i - n_modes//2) * sigma * 1.5
            mode_data = np.random.normal(center, sigma*0.4, samples_per_mode)
            data.extend(mode_data)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.random.normal(base_value, sigma*0.4, remaining))
        data = np.array(data)
        
    elif pattern == "Random-Walk":
        data = [base_value]
        for _ in range(n_samples - 1):
            step = np.random.normal(0, sigma*0.3)
            data.append(data[-1] + step)
        data = np.array(data)
        
    elif pattern == "Spike":
        data = np.random.normal(base_value, sigma*0.5, n_samples)
        n_spikes = min(random.randint(2, 5), n_samples)
        spike_indices = np.random.choice(n_samples, n_spikes, replace=False)
        spike_magnitude = np.random.choice([-1, 1], n_spikes) * sigma * random.uniform(5, 8)
        data[spike_indices] += spike_magnitude
        
    elif pattern == "Exponential":
        data = np.random.exponential(sigma, n_samples) + base_value - sigma
        
    elif pattern == "Uniform":
        data = np.random.uniform(base_value - 2*sigma, base_value + 2*sigma, n_samples)
        
    elif pattern == "U-Shape":
        middle_samples = max(1, int(n_samples * 0.1))
        side_samples = (n_samples - middle_samples) // 2
        remaining = n_samples - middle_samples - 2 * side_samples
        data1 = np.random.normal(base_value - 2*sigma, sigma*0.5, side_samples)
        data2 = np.random.normal(base_value + 2*sigma, sigma*0.5, side_samples)
        data_middle = np.random.normal(base_value, sigma*0.3, middle_samples)
        if remaining > 0:
            data_extra = np.random.normal(base_value + 2*sigma, sigma*0.5, remaining)
            data = np.concatenate([data1, data2, data_middle, data_extra])
        else:
            data = np.concatenate([data1, data2, data_middle])
        
    elif pattern == "Sawtooth":
        n_cycles = random.randint(3, 6)
        samples_per_cycle = n_samples // n_cycles
        data = []
        for _ in range(n_cycles):
            cycle = np.linspace(base_value - sigma, base_value + sigma, samples_per_cycle)
            cycle += np.random.normal(0, sigma*0.1, samples_per_cycle)
            data.extend(cycle)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.linspace(base_value - sigma, base_value + sigma, remaining))
        data = np.array(data)
        
    elif pattern == "Chaos":
        parts = random.randint(3, 5)
        part_size = n_samples // parts
        data = []
        sub_patterns = ["Normal", "Uniform", "Exponential", "Spike"]
        for i in range(parts):
            sub_pattern = random.choice(sub_patterns)
            # Chaos 內部也使用相同的 decimals
            sub_data = generate_pattern_data(sub_pattern, part_size, base_value, sigma, decimals)
            data.extend(sub_data)
        remaining = n_samples - len(data)
        if remaining > 0:
            data.extend(np.random.normal(base_value, sigma, remaining))
        data = np.array(data)
        
    else:
        data = np.random.normal(base_value, sigma, n_samples)
    
    data = np.array(data)
    
    # 長度調整
    if len(data) < n_samples:
        shortage = n_samples - len(data)
        extra = np.random.normal(base_value, sigma, shortage)
        data = np.concatenate([data, extra])
    elif len(data) > n_samples:
        data = data[:n_samples]
    
    # 打亂順序 (除非是有序的時間序列 Pattern)
    ordered_patterns = ["Step", "Step-Up", "Step-Down", "Cyclic", "Trending-Up", "Trending-Down", "Random-Walk", "Sawtooth"]
    if pattern not in ordered_patterns and len(data) > 1:
        np.random.shuffle(data)
        
    # 數值上做 Rounding
    data = np.round(data, decimals)
    
    return data

def generate_test_charts():
    """生成 200 張測試圖表的配置和數據"""
    
    patterns = ["Normal", "Skew-Right", "Skew-Left", "Bimodal", "Attribute", "Constant", "Near Constant", 
                "Step", "Step-Up", "Step-Down", "Cyclic", "Trending-Up", "Trending-Down", 
                "Outliers", "Multimodal", "Random-Walk", "Spike", "Exponential", "Uniform", 
                "U-Shape", "Sawtooth", "Chaos"]
    
    sample_ranges = [
        (100, 300), (300, 600), (600, 1000), (1000, 2000)
    ]
    characteristics = ["Nominal", "Smaller", "Bigger"]
    
    # === 修改 1: 設定 1~5 位小數點的權重 ===
    possible_decimals = [1, 2, 3, 4, 5] 
    # 權重分配 (您可以依需求調整，這裡假設 2,3 位最常見)
    decimal_weights = [0.1, 0.3, 0.3, 0.2, 0.1]
    
    charts_info = []
    
    for i in range(50):
        pattern = "Attribute"
        sample_range = random.choice(sample_ranges)
        n_samples = random.randint(sample_range[0], sample_range[1])
        characteristic = "Nominal"
        n_cat = random.randint(5, 15)  # 每張圖隨機不同類別數
        
        # 隨機決定這張 Chart 的小數點位數 (1~5)
        n_decimals = np.random.choice(possible_decimals, p=decimal_weights)
        resolution_value = 10 ** (-int(n_decimals))
        
        # 基礎參數
        base_value = random.uniform(8, 12)
        sigma = random.uniform(0.5, 2.0)
        
        # 生成數據 (帶入 decimals 與 n_categories 參數)
        data = generate_pattern_data(pattern, n_samples, base_value, sigma, decimals=n_decimals, n_categories=n_cat)
        
        # 設定 Target 和管制線 (也要 round 到相同位數)
        target = round(base_value, n_decimals)
        
        if characteristic == "Nominal":
            ori_ucl = round(target + 4.5 * sigma, n_decimals)
            ori_lcl = round(target - 4.5 * sigma, n_decimals)
        elif characteristic == "Smaller":
            ori_ucl = round(target + 4 * sigma, n_decimals)
            ori_lcl = round(target - 5.5 * sigma, n_decimals)
        else: # Bigger
            ori_ucl = round(target + 5.5 * sigma, n_decimals)
            ori_lcl = round(target - 4 * sigma, n_decimals)
            
        usl = round(ori_ucl + 2 * sigma, n_decimals)
        lsl = round(ori_lcl - 2 * sigma, n_decimals)
        
        chart_info = {
            'GroupName': f'TestGroup_{i//10 + 1}',
            'ChartName': f'Chart_{i+1:03d}',
            'ChartID': f'TC{i+1:03d}',
            'Material_no': f'MAT_{i+1:03d}',
            'Target': target,
            'UCL': ori_ucl,
            'LCL': ori_lcl,
            'USL': usl,
            'LSL': lsl,
            'Characteristics': characteristic,
            'DetectionLimit': round(target - 3.5 * sigma, n_decimals) if characteristic == 'Smaller' else None,
            'ExpectedPattern': pattern,
            'SampleCount': n_samples,
            'Resolution': resolution_value
        }
        
        charts_info.append(chart_info)
        
        start_date = datetime.now() - timedelta(days=365*2)
        dates = [start_date + timedelta(days=random.randint(0, 730)) for _ in range(n_samples)]
        dates.sort()
        
        # --- 新增 ByTool 邏輯 ---
        # 隨機決定這張 Chart 是由哪幾台 Tool 生產的 (假設每張圖表由 2~4 台機器輪替)
        num_tools = random.randint(2, 4)
        available_tools = [f"TOOL_{random.randint(101, 150):03d}" for _ in range(num_tools)]
        
        # 為每一筆數據隨機指派一個機台 (也可以用循環指派，這裡採隨機指派模擬真實輪替)
        tools_col = [random.choice(available_tools) for _ in range(n_samples)]
        
        batch_ids = []
        for date in dates:
            date_str = date.strftime('%Y%m%d')
            sequence = random.randint(1, 999)
            batch_id = f"BATCH-{date_str}-{sequence:03d}"
            batch_ids.append(batch_id)
        
        csv_data = pd.DataFrame({
            'point_time': dates,
            'point_val': data,
            'Batch_ID': batch_ids,
            'Matching': tools_col  # 新增的機台欄位
        })

        # --- 計算 per-row cpk（使用滑動視窗 std 估算過程能力）---
        _win = min(20, max(5, len(data) // 10))
        _vals_s = pd.Series(data)
        _rolling_std = _vals_s.rolling(window=_win, min_periods=3).std().bfill()
        _rolling_std = _rolling_std.replace(0, np.nan)
        _usl = chart_info.get('USL')
        _lsl = chart_info.get('LSL')
        def _cpk(pv, rs):
            if rs is None or (isinstance(rs, float) and np.isnan(rs)) or rs <= 0:
                return None
            cpu = ((_usl - pv) / (3 * rs)) if _usl is not None else None
            cpl = ((pv - _lsl) / (3 * rs)) if _lsl is not None else None
            if cpu is not None and cpl is not None:
                return round(min(cpu, cpl), 4)
            return round(cpu, 4) if cpu is not None else (round(cpl, 4) if cpl is not None else None)
        csv_data['cpk'] = [_cpk(pv, rs) for pv, rs in zip(data, _rolling_std)]
        # -----------------------
        
        os.makedirs('input/raw_charts', exist_ok=True)
        csv_filename = f"input/raw_charts/{chart_info['GroupName']}_{chart_info['ChartName']}.csv"
        
        # === 修改 2: 寫入 CSV 時強制指定 float_format ===
        # 這樣可以確保該檔案內的數值都有統一的小數位數 (包含補零)
        csv_data.to_csv(csv_filename, index=False, float_format=f'%.{n_decimals}f')
    
    charts_df = pd.DataFrame(charts_info)
    os.makedirs('input', exist_ok=True)
    charts_df.to_excel('input/All_Chart_Information.xlsx', sheet_name='Chart', index=False)
    
    print(f"✅ 已生成 200 張測試圖表 (Resolution 模擬範圍: 1~5 位小數)")
    print(f"📊 Resolution 分佈:\n{charts_df['Resolution'].value_counts().sort_index()}")
    
    return charts_df

if __name__ == '__main__':
    print("🚀 開始生成測試數據...")
    generate_test_charts()
    print("✅ 測試數據生成完成！請執行主程式進行測試。")