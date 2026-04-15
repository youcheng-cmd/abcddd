import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (終極修正版)")

# --- 1. 側邊欄參數設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", value=95)
age_filter = st.sidebar.selectbox("選擇變壓器齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

# --- 通用工具函數 ---
def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_number(text):
    """強力數字提取器：處理逗號、單位、百分比"""
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取 Excel 內容
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    seen_device_keys = set() # 核心：用來防止台數重複抓取

    # 遍歷所有工作表
    for sheet_name, raw_df in all_sheets.items():
        # 尋找「序號」錨點
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_val = str(raw_df.iloc[r, c]).replace(' ', '')
                if cell_val == "序號":
                    # 確認下方是否有相關標籤，避免誤抓非規格表的「序號」
                    if r + 1 < len(raw_df):
                        check_label = str(raw_df.iloc[r+1, c]).replace(' ', '')
                        if any(k in check_label for k in ["建築", "位置", "編號"]):
                            anchors.append((r, c))
        
        for r_start, c_start in anchors:
            # 橫向掃描數據欄位 (通常是 TR-1 ~ TR-7)
            for offset in range(1, 12):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                # 檢查該欄是否有內容（序號數字）
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                if sn_val in ["nan", ""] or not sn_val.isdigit(): continue
                
                # 初始化單台設備字典
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0.0, 
                     "型式": "-", "負載率": 0.0, "現況功因": 0.0, "鐵損": 0.0, "實際銅損": 0.0, "改善前耗能": 0.0}
                specs = []
                
                # 垂直掃描該設備的所有屬性
                for r_offset in range(0, 55):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip() if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    label_p = label.replace(' ', '')
                    
                    # 碰到下一個序號標籤就停止掃描此台設備
                    if r_offset > 0 and label_p == "序號": break
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label_p == "nan" or not label_p: continue

                    # 數據精準對標抓取
                    if any(k in label_p for k in ["建築", "位置"]): d["建築物"] = val
                    if "編號" in label_p: d["編號"] = val
                    if "廠牌" in label_p: d["廠牌"] = val
                    if "型式" in label_p: d["型式"] = val
                    if any(k in label_p for k in ["年份", "出廠"]):
                        n
