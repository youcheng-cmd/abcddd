import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化報告 (公式全量計算版)")

# --- 側邊欄設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
age_filter = st.sidebar.selectbox("機齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_number(text):
    """提取純數字，排除單位與符號"""
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    
    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if "序號" in str(raw_df.iloc[r, c]):
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            for offset in range(1, 10):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                if pd.isna(raw_df.iloc[r_start, target_col]): continue
                
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                     "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8}
                specs = []
                
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 左右鄰居標籤探測
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '').replace('\n', '') if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    
                    if r_offset > 0 and "序號" in label: break
                    val_str = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label == "nan" or not label: continue

                    # 提取關鍵數值
                    if any(k in label for k in ["建築", "位置"]): d["建築物"] = val_str
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val_str)
                        d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "編號" in label: d["編號"] = val_str
                    if "廠牌" in label: d["廠牌"] = val_str
                    if "型式" in label: d["型式"] = val_str
                    if "容量" in label: d["容量"] = int(extract_number(val_str))
                    if any(k in label for k in ["利用率", "負載率"]):
                        n = extract_number(val_str)
                        d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "PF"]):
                        n = extract_number(val_str)
                        d["現況功因"] = n / 100 if n > 1 else n
                    if any(k in label for k in ["鐵損", "無載損"]): d["鐵損"] = extract_number(val_str)
                    if any(k in label for k in ["銅損", "負載損", "全載損"]): d["滿載銅損"] = extract_number(val_str)

                    if val_str != "nan": specs.append((label, val_str))
                
                if d["容量"] > 0:
                    # 篩選
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # --- 關鍵公式計算 ---
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})

    if all_transformer_data:
        st.success(f"✅ 解析完成，符合條件共 {len(all_transformer_data)} 台。")

        # 生成 Word
        doc = Document()
        
        # 1. 改善前分析大表 (11欄)
        doc.add_heading('壹、 變壓器設備改善前數據分析表', 1)
        ana_t = doc.add_table(rows=1, cols=11); ana_t.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量(kVA)", "型式", "負載率(%)", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能(kWh/年)"]
        for i, h in enumerate(headers): set_font_kai(ana_t.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        
        for item in all_transformer_data:
            d = item["analysis"]
            row = ana_t.add_row().cells
            # 填入計算後的數值
            row_data = [
                d["建築物"], d["編號"], d["年份"], d["廠牌"], d["容量"], d["型式"],
                f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}",
                f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"
            ]
            for i, val in enumerate(row_data):
                set_font_kai(row[i].paragraphs[0].add_run(str(val)), 8)
        
        doc.add_page_break()
        # (後續章節...)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整分析報告", output, "Energy_Report.docx")
