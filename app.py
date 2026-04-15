import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化報告 (功能全數找回版)")

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
                
                d = {
                    "建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                    "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8
                }
                
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '')
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '') if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    
                    if r_offset > 0 and "序號" in label: break
                    val_str = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    if "建築" in label: d["建築物"] = val_str
                    if "編號" in label: d["編號"] = val_str
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val_str)
                        d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "容量" in label: d["容量"] = int(extract_number(val_str))
                    if any(k in label for k in ["負載率", "利用率"]):
                        n = extract_number(val_str)
                        d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "PF"]):
                        n = extract_number(val_str)
                        d["現況功因"] = n / 100 if n > 1 else n
                    if any(k in label for k in ["鐵損", "無載損"]): d["鐵損"] = extract_number(val_str)
                    if any(k in label for k in ["銅損", "負載損", "全載損"]): d["滿載銅損"] = extract_number(val_str)

                if d["容量"] > 0:
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # --- 強制公式計算 ---
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    all_transformer_data.append(d)

    if all_transformer_data:
        # --- 找回的功能 1, 2, 3 ---
        total_capacity = sum(t["容量"] for t in all_transformer_data)
        cap_counts = Counter(t["容量"] for t in all_transformer_data)
        avg_usage = sum(t["負載率"] for t in all_transformer_data) / len(all_transformer_data)

        st.success(f"✅ 解析完成！符合條件共 {len(all_transformer_data)} 台。")
        
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("1. 總裝置容量")
            st.metric(label="", value=f"{total_capacity} kVA")
            st.subheader("2. 設備規格分布")
            for k, v in sorted(cap_counts.items(), reverse=True):
                st.write(f"🔹 {k} kVA × {v} 台")
        with c2:
            st.subheader("3. 平均負載利用率")
            st.metric(label="", value=f"{avg_usage:.2f} %")
        
        st.divider()
        st.write("🔍 **計算除錯預覽 (請確認下方數據非 0)**")
        st.dataframe(pd.DataFrame(all_transformer_data)[["編號", "容量", "負載率", "實際銅損", "鐵損", "改善前耗能"]])

        # --- Word 產出邏輯 ---
        doc = Document()
        # 壹、統計總表
        doc.add_heading('壹、 設備統計總表', 1)
        stb = doc.add_table(rows=0, cols=2); stb.style = 'Table Grid'
        def add_r(l, v):
            row = stb.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(v), 12)
        add_r("總裝置容量", f"{total_capacity} kVA")
        add_r("設備規格分布", "、".join([f"{k}kVAx{v}台" for k, v in sorted(cap_counts.items(), reverse=True)]))
        add_r("平均負載利用率", f"{avg_usage:.2f} %")
        
        # 貳、分析大表
        doc.add_page_break()
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_t = doc.add_table(rows=1, cols=11); ana_t.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_t.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        for d in all_transformer_data:
            row = ana_t.add_row().cells
            vals = [d["建築物"], d["編號"], d["年份"], d["廠牌"], d["容量"], d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, v in enumerate(vals): set_font_kai(row[i].paragraphs[0].add_run(str(v)), 8)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, "Transformer_Final_Report.docx")
