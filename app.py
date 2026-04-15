import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (損耗計算回歸版)")

# --- 1. 側邊欄參數設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("基準年份 (計算機齡)：", value=2026)
age_filter = st.sidebar.selectbox("機齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])
pf_after_input = st.sidebar.number_input("改善後目標功率因數 (%)：", value=95)

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
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                if sn_val in ["nan", ""]: continue
                
                d = {
                    "建築物": "-", "編號": sn_val, "年份": 0, "廠牌": "-", "容量": 0.0, 
                    "型式": "-", "負載率": 0.0, "現況功因": 0.8, "鐵損": 0.0, "滿載銅損": 0.0
                }
                specs = []
                
                for r_offset in range(0, 60):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace(' ', '')
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace(' ', '') if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    if r_offset > 0 and "序號" in label: break
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label == "nan" or not label: continue

                    # 數據抓取
                    if any(k in label for k in ["建築", "位置"]): d["建築物"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val); d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "廠牌" in label: d["廠牌"] = val
                    if "型式" in label: d["型式"] = val
                    if "容量" in label: d["容量"] = extract_number(val)
                    if any(k in label for k in ["負載率", "利用率"]):
                        n = extract_number(val); d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "PF"]):
                        n = extract_number(val); d["現況功因"] = n / 100 if n > 1 else n
                    
                    # 抓取銅損鐵損基礎值
                    if any(k in label for k in ["鐵損", "無載損", "Wi"]): d["鐵損"] = extract_number(val)
                    if any(k in label for k in ["銅損", "負載損", "Wc"]): d["滿載銅損"] = extract_number(val)

                    specs.append((label, val))

                if d["容量"] > 0:
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # --- 重點：銅損鐵損與耗能公式計算 ---
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})

    if all_transformer_data:
        # --- 2. 網頁即時統計摘要 ---
        total_cap = sum(t["analysis"]["容量"] for t in all_transformer_data)
        cap_counts = Counter(t["analysis"]["容量"] for t in all_transformer_data)
        avg_usage = sum(t["analysis"]["負載率"] for t in all_transformer_data) / len(all_transformer_data)
        
        st.success(f"✅ 解析完成！符合條件共 {len(all_transformer_data)} 台")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("1. 總裝置容量", f"{total_cap:,.0f} kVA")
            st.write("**2. 規格台數分布：**")
            for k, v in sorted(cap_counts.items(), reverse=True): st.write(f"🔹 {k:,.0f} kVA × {v} 台")
        with c2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # --- 4. Word 報告產出 ---
        doc = Document()
        
        # 壹、設備統計總表
        doc.add_heading('壹、 設備統計總表', 1)
        st_table = doc.add_table(rows=0, cols=2); st_table.style = 'Table Grid'
        def add_sum_row(l, v):
            row = st_table.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(v), 12)
        add_sum_row("總裝置容量", f"{total_cap:,.0f} kVA")
        add_sum_row("設備規格分布", "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True)]))
        add_sum_row("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、改善前數據分析表 (11欄)
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_t = doc.add_table(rows=1, cols=11); ana_t.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_t.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        for item in all_transformer_data:
            d = item["analysis"]
            row = ana_t.add_row().cells
            row_data = [d["建築物"], d["編號"], d["年份"], d["廠牌"], f"{d['容量']:.0f}", d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, val in enumerate(row_data): set_font_kai(row[i].paragraphs[0].add_run(str(val)), 8)
        doc.add_page_break()

        # 參、詳細資料
        doc.add_heading('參、 詳細設備數據', 1)
        for item in all_transformer_data:
            doc.add_paragraph().add_run(f"設備詳細資料 (編號：{item['analysis']['編號']})").bold = True
            t_dt = doc.add_table(rows=0, cols=2); t_dt.style = 'Table Grid'
            for l, v in item["specs"]:
                cells = t_dt.add_row().cells
                set_font_kai(cells[0].paragraphs[0].add_run(l), 10)
                set_font_kai(cells[1].paragraphs[0].add_run(v), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, f"Transformer_Report.docx")
