import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (全數據對齊版)")

# --- 側邊欄設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", min_value=1, max_value=100, value=95)
pf_after = pf_after_input / 100

age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"]
)

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_number(text):
    if pd.isna(text): return 0.0
    clean_text = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", clean_text)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    
    for sheet_name, raw_df in all_sheets.items():
        # 轉成字串並清洗，方便搜尋
        df_str = raw_df.astype(str).apply(lambda x: x.str.replace(' ', '').replace('\n', ''))
        
        anchors = []
        for r in range(len(df_str)):
            for c in range(len(df_str.columns)):
                if "序號" in df_str.iloc[r, c]:
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            for offset in range(1, 10):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                # 取得編號 (例如 TR-7)
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                if sn_val == "nan" or sn_val == "": continue
                
                d = {"建築物": "-", "編號": sn_val, "年份": 0, "廠牌": "-", "容量": 0, 
                     "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8}
                specs = []
                
                # 往下掃描 50 列
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # --- 核心優化：多重標籤探測 ---
                    # 同時看原本的 c_start 跟左邊一格 c_start-1
                    label_1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').strip()
                    label_2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '').strip() if c_start > 0 else ""
                    
                    # 決定哪一個才是有效的標籤
                    label = label_1 if (label_1 != "nan" and label_1 != "") else label_2
                    
                    if r_offset > 0 and "序號" in label: break
                    
                    val_cell = raw_df.iloc[curr_r, target_col]
                    val_str = str(val_cell).strip()
                    
                    if label == "nan" or not label: continue

                    # 數據精確提取
                    if any(k in label for k in ["建築", "位置"]): d["建築物"] = val_str
                    if any(k in label for k in ["年份", "出廠"]):
                        num = extract_number(val_str)
                        d["年份"] = int(num) + 1911 if 0 < num < 200 else int(num)
                    if "廠牌" in label: d["廠牌"] = val_str
                    if "型式" in label: d["型式"] = val_str
                    if "容量" in label: d["容量"] = int(extract_number(val_str))
                    if any(k in label for k in ["利用率", "負載率"]):
                        num = extract_number(val_str)
                        d["負載率"] = num * 100 if 0 < num < 1 else num
                    if any(k in label for k in ["功因", "PF"]):
                        num = extract_number(val_str)
                        d["現況功因"] = num / 100 if num > 1 else num
                    if any(k in label for k in ["鐵損", "無載損"]): d["鐵損"] = extract_number(val_str)
                    if any(k in label for k in ["銅損", "負載損", "全載損"]): d["滿載銅損"] = extract_number(val_str)

                    specs.append((label, val_str))
                
                if d["容量"] > 0: # 確保真的有抓到資料
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                    d["改善前耗能"] = (d["鐵損"] +
