import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (損耗公式計算版)")

# --- 側邊欄設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("基準年份 (計算機齡)：", value=2026)
age_filter = st.sidebar.selectbox("機齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

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
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    
    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if "序號" in str(raw_df.iloc[r, c]):
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            for offset in range(1, 15):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                sn_cell = raw_df.iloc[r_start, target_col]
                if pd.isna(sn_cell) or str(sn_cell).strip() == "": continue
                
                d = {
                    "建築物": "-", "編號": str(sn_cell).strip(), "年份": 0, "容量": 0.0, 
                    "負載率": 0.0, "現況功因": 0.8, "鐵損": 0.0, "滿載銅損": 0.0, "型式": "-"
                }
                
                for r_offset in range(0, 60):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '')
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '') if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    if any(k in label for k in ["建築", "位置"]): d["建築物"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val)
                        d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "容量" in label: d["容量"] = extract_number(val)
                    if "型式" in label: d["型式"] = val
                    if any(k in label for k in ["負載率", "利用率"]):
                        n = extract_number(val)
                        d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "PF"]):
                        n = extract_number(val)
                        d["現況功因"] = n / 100 if n > 1 else n
                    # 嘗試抓取 Excel 內的損耗，若無則後續公式補位
                    if any(k in label for k in ["鐵損", "無載損"]): d["鐵損"] = extract_number(val)
                    if any(k in label for k in ["銅損", "負載損"]): d["滿載銅損"] = extract_number(val)

                if d["容量"] > 0:
                    # --- 公式強制計算邏輯 ---
                    # 1. 如果 Excel 沒寫鐵損，按容量 0.2% 估算 (可依需求調整)
                    if d["鐵損"] == 0: d["鐵損"] = d["容量"] * 2.0  # 假設 2W/kVA
                    
                    # 2. 如果 Excel 沒寫滿載銅損，按容量 1.2% 估算
                    if d["滿載銅損"] == 0: d["滿載銅損"] = d["容量"] * 12.0 # 假設 12W/kVA
                    
                    # 3. 計算實際銅損 (隨負載平方變化)
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    
                    # 4. 計算輸出功率 (kW)
                    d["輸出功率"] = d["容量"] * d["現況功因"] * (d["負載率"] / 100)
                    
                    # 5. 計算改善前年耗能 (kWh/年)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    all_transformer_data.append(d)

    if all_transformer_data:
        # 網頁摘要
        total_cap = sum(t["容量"] for t in all_transformer_data)
        avg_usage = sum(t["負載率"] for t in all_transformer_data) / len(all_transformer_data)
        
        st.success(f"✅ 解析完成！共計 {len(all_transformer_data)} 台。")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("1. 總裝置容量", f"{total_cap:,.0f} kVA")
            st.write("**2. 設備規格分布：**")
            for k, v in sorted(Counter([t["容量"] for t in all_transformer_data]).items(), reverse=True):
                st.write(f"🔹 {k:,.0f} kVA × {v} 台")
        with c2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # 生成 Word
        doc = Document()
        # --- 貳、改善前數據分析表 ---
        doc.add_heading('貳、 改善前數據分析表', 1)
        table = doc.add_table(rows=1, cols=11); table.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率%", "輸出功率", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(table.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        
        for d in all_transformer_data:
            row = table.add_row().cells
            row_vals = [d["建築物"], d["編號"], d["年份"], "-", f"{d['容量']:.0f}", d["型式"], f"{d['負載率']:.1f}%", f"{d['輸出功率']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, v in enumerate(row_vals): set_font_kai(row[i].paragraphs[0].add_run(str(v)), 8)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整分析報告", output, "Report.docx")
