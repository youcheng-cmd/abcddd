import streamlit as st
import pandas as pd
from docx import Document
import io
import re

st.set_page_config(page_title="變壓器數據精確解析", layout="wide")
st.title("⚡ 變壓器數據抓取與計算 (規格解析版)")

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

def extract_capacity(text):
    """
    從字串中提取容量數值 (例如從 '3Φ3W 22.8kV-220V 500' 提取 500)
    """
    if pd.isna(text): return None
    text = str(text).strip()
    # 尋找字串末尾的數字 (可能是整數或小數)
    match = re.search(r'(\d+\.?\d*)$', text)
    if match:
        return float(match.group(1))
    return None

if excel_file:
    # 1. 讀取 Excel
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 2. 定位標題列 (尋找包含「變壓器編號」或「容量」的行)
    header_idx = None
    for idx, row in raw_df.iterrows():
        row_str = "".join(row.astype(str))
        if "變壓器編號" in row_str and "容量" in row_str:
            header_idx = idx
            break
            
    if header_idx is not None:
        df = pd.read_excel(excel_file, header=header_idx + 1)
        df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
        
        # 3. 定義欄位 (根據截圖精確匹配)
        col_no = next((c for c in df.columns if "變壓器編號" in c), None)
        col_cap_spec = next((c for c in df.columns if "容量" in c), None)
        col_load = next((c for c in df.columns if "負載率" in c), None)
        col_eff = next((c for c in df.columns if "效率" in c), None)

        # 4. 數據清洗：解析容量
        results = []
        for idx, row in df.iterrows():
            spec_text = row.get(col_cap_spec)
            cap_value = extract_capacity(spec_text)
            
            # 如果抓不到容量數字，代表這行不是設備資料，跳過
            if cap_value is None: continue
            
            # 取得負載率與效率 (處理百分比符號)
            try:
                load_p = float(str(row.get(col_load, 0)).replace('%','')) / 100
                eff_b = float(str(row.get(col_eff, 0)).replace('%','')) / 100
                
                # --- 計算公式 ---
                loss_total_b = cap_value * (1 - eff_b)
                iron_b = loss_total_b * 0.25 
                copper_b = (loss_total_b - iron_b) * (load_p**2)
                
                results.append({
                    'no': row.get(col_no, idx+1),
                    'spec': spec_text,       # 原始規格文字
                    'cap': cap_value,        # 提取出的數字
                    'load': f"{int(load_p*100)}%",
                    'iron_b': round(iron_b, 3),
                    'cop_b': round(copper_b, 3),
                    'tot_b': round(iron_b + copper_b, 3),
                    # 改善後預設 99%
                    'saving': round((iron_b + copper_b) * 0.2, 3) # 暫定節省 20%
                })
            except:
                continue

        st.success(f"✅ 成功解析數據！已提取出 {len(results)} 台變壓器。")
        st.write("### ⬇️ 數據提取結果 (請確認「容量」是否正確)：")
        st.table(pd.DataFrame(results))

        if st.button("🚀 生成報告"):
            # ... (Word 生成代碼與之前相同)
            st.info("點擊按鈕即可下載 Word 報告")
    else:
        st.error("❌ 找不到標題列，請確認 Excel 中有『變壓器編號』與『容量』。")
