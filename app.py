import re

def extract_number(text):
    """將文字中的數字挖出來，支援小數、逗號與單位"""
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0
    if specs:
                    # 1. 計算機齡與篩選
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # 2. 【核心計算公式】在此強制執行
                    # 實際銅損 = 滿載銅損 * (負載率/100)^2
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    
                    # 改善前耗能 (kWh/年) = (鐵損 + 實際銅損) * 8760 / 1000
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    # 將計算結果存入總表
                    all_transformer_data.append({"specs": specs, "analysis": d, "capacity": d["容量"], "usage_rate": d["負載率"]})
        # 修改抓取損耗的標籤識別
                    if any(k in label for k in ["無載損", "鐵損", "Wi"]):
                        d["鐵損"] = extract_number(val)
                    if any(k in label for k in ["負載損", "銅損", "Wc", "全載損"]):
                        d["滿載銅損"] = extract_number(val)
