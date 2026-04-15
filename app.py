import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器專業報告-彙總統計版", layout="wide")
st.title("📑 變壓器自動化報告 (數據精確修正版)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 設定基準與篩選")
base_year = st.sidebar.number_input("請輸入基準年份：", min_value=1900, max_value=2100, value=2026)
age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年 (汰換參考)", "超過 15 年 (優先汰換)", "超過 20 年 (屆齡汰換)"]
)

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機", "電能系統"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_text = str(raw_df.iloc[r, c]).replace(' ', '').replace('\n', '')
                if cell_text == "序號":
                    anchors.append((r, c))
        
        if anchors:
            for r_start, c_start in anchors:
                for offset in range(1, 10):
                    target_col = c_start + offset
                    if target_col >= len(raw_df.columns): break
                    sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                    if sn_val == "nan" or sn_val == "": continue
                    
                    specs = []
                    mfg_year = None 
                    capacity = 0
                    usage_rate = None
                    
                    for r_offset in range(0, 50):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        
                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        label_clean = label.replace(' ', '')
                        if r_offset > 0 and label_clean == "序號": break
                        if any(k in label_clean for k in filter_keywords): continue
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label: continue
                        
                        # --- 數據提取優化 ---
                        if any(k in label for k in ["製造年份", "製造日期", "出廠", "年份"]):
                            digits = ''.join(filter(str.isdigit, value))
                            if digits:
                                y = int(digits)
                                mfg_year = y + 1911 if y < 200 else y
                        
                        if "容量" in label:
                            cap_digits = ''.join(filter(str.isdigit, value))
                            if cap_digits: capacity = int(cap_digits)
                        
                        # 利用率處理：防止小數與百分比混淆
                        if any(k in label for k in ["利用率", "負載率"]):
                            u_raw = value.replace('%', '').strip()
                            try:
                                u_val = float(u_raw)
                                # 判斷是否為小數格式 (如 0.32 -> 32)
                                if 0 < u_val < 1:
                                    usage_rate = u_val * 100
                                else:
                                    usage_rate = u_val
                            except:
                                usage_rate = None
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        age = base_year - mfg_year if mfg_year else 0
                        should_add = False
                        if age_filter == "顯示全部": should_add = True
                        elif "10" in age_filter and age >= 10: should_add = True
                        elif "15" in age_filter and age >= 15: should_add = True
                        elif "20" in age_filter and age >= 20: should_add = True
                        
                        if should_add:
                            all_transformer_data.append({
                                "specs": specs,
                                "capacity": capacity,
                                "usage_rate": usage_rate,
                                "age": age
                            })

    if all_transformer_data:
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        
        # 只取有意義的利用率進行平均 (排除 0 或 None)
        valid_usages = [t["usage_rate"] for t in all_transformer_data if t["usage_rate"] is not None and t["usage_rate"] > 0]
        avg_usage = sum(valid_usages) / len(valid_usages) if valid_usages else 0

        st.success(f"✅ 修正完成！平均負載利用率已更新。")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("總裝置容量", f"{total_capacity} kVA")
        col2.metric("平均負載利用率", f"{avg_usage:.1f} %") # 這裡應該會顯示正常的 2x.x % 了
        col3.metric("篩選後總台數", f"{len(all_transformer_data)} 台")

        if st.button("🚀 下載修正版報告"):
            doc = Document()
            # (省略 Word 生成部分，保持與之前一致)
            # ... (Word 生成代碼)
            st.write("報告產出中...")
