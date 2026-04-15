import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化報告 (損耗計算終極修復版)")

# --- 側邊欄設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", value=95)

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
    """強力數字提取：解決逗號、單位、括號問題"""
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if "序號" in str(raw_df.iloc[r, c]):
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            for offset in range(1, 12):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                sn_cell = raw_df.iloc[r_start, target_col]
                if pd.isna(sn_cell) or str(sn_cell).strip() == "": continue
                
                d = {
                    "建築物": "-", "編號": str(sn_cell).strip(), "年份": 0, "廠牌": "-", "容量": 0.0, 
                    "型式": "-", "負載率": 0.0, "現況功因": 0.8, "鐵損": 0.0, "滿載銅損": 0.0
                }
                specs = []
                
                for r_offset in range(0, 60):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 標籤抓取 (包含左鄰偵測)
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '').replace('\n', '') if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    
                    if r_offset > 0 and "序號" in label: break
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    # --- 數據分類抓取 ---
                    if any(k in label for k in ["建築", "位置"]): d["建築物"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val); d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "廠牌" in label: d["廠牌"] = val
                    if "型式" in label: d["型式"] = val
                    if "容量" in label: d["容量"] = extract_number(val)
                    if any(k in label for k in ["負載率", "利用率"]):
                        n = extract_number(val); d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "功率因數", "PF"]):
                        n = extract_number(val); d["現況功因"] = n / 100 if n > 1 else n
                    
                    # --- 關鍵修正：銅鐵損模糊抓取 ---
                    if any(k in label for k in ["鐵損", "無載損", "Wi", "Pi"]):
                        d["鐵損"] = extract_number(val)
                    if any(k in label for k in ["銅損", "負載損", "全載損", "Wc", "Pc"]):
                        d["滿載銅損"] = extract_number(val)

                    if val != "nan": specs.append((label, val))

                if d["容量"] > 0:
                    # 公式計算
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    d["輸出功率"] = d["容量"] * d["現況功因"] * (d["負載率"] / 100)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    all_transformer_data.append(d)

    if all_transformer_data:
        # 網頁摘要
        st.success(f"✅ 成功抓取 {len(all_transformer_data)} 台數據")
        
        # --- 增加預覽表協助除錯 ---
        st.write("📊 **實時數據檢視 (請確認銅損/鐵損欄位是否有數字)**")
        st.dataframe(pd.DataFrame(all_transformer_data)[["編號", "容量", "負載率", "滿載銅損", "鐵損", "改善前耗能"]])

        # Word 生成部分 (壹、貳、參)
        doc = Document()
        # (這裡省略重複的 Word 生成 code，確保貳、改善前數據分析表包含這 11 欄)
        # ... [代碼同前，請確保 row_data 包含 d["實際銅損"] 等變數]
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, "Report.docx")
