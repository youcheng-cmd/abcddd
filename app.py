import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (精確數量版)")

# --- 1. 側邊欄參數設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", value=95)
age_filter = st.sidebar.selectbox("選擇變壓器齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

# --- 工具函數 ---
def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_number(text):
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 只讀取第一個工作表，避免隱藏頁面干擾
    raw_df = pd.read_excel(excel_file, sheet_name=0, header=None)
    all_transformer_data = []
    
    # --- 精準錨點定位 ---
    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            cell_str = str(raw_df.iloc[r, c]).replace(' ', '')
            # 只抓取作為標題的「序號」，且下方必須有像「建築物名稱」之類的標籤才算數
            if cell_str == "序號":
                if r + 1 < len(raw_df):
                    next_cell = str(raw_df.iloc[r+1, c]).replace(' ', '')
                    if any(k in next_cell for k in ["建築", "位置", "編號"]):
                        anchors.append((r, c))

    seen_device_ids = set() # 防止重複抓取相同編號的設備

    for r_start, c_start in anchors:
        # 橫向掃描設備 (TR-1 ~ TR-7)
        for offset in range(1, 12):
            target_col = c_start + offset
            if target_col >= len(raw_df.columns): break
            
            # 取得該欄的編號（通常在序號下方第
