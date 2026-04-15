import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (基準穩定版)")

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
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if str(raw_df.iloc[r, c]).replace(' ', '') == "序號":
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            # 橫向掃描 TR-1 ~ TR-7
            for offset in range(1, 12):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                if pd.isna(raw_df.iloc[r_start, target_col]) or str(raw_df.iloc[r_start, target_col]).strip() == "": continue
                
                # 初始化設備字典
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0.0, 
                     "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8}
                specs = []
                
                # 垂直掃描數據
                for r_offset in range(0, 60):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 標籤偵測 (含左鄰補位，解決 TR-7 問題)
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip() if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    
                    if r_offset > 0 and label.replace(' ', '') == "序號": break
                    if any(k in label for k in filter_keywords): continue
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label == "nan" or not label: continue

                    # 3. 數據自動抓取邏輯 (核心)
                    if any(k in label for k in ["位置", "建築"]): d["建築物"] = val
                    if "編號" in label: d["編號"] = val
                    if "廠牌" in label: d["廠牌"] = val
                    if "型式" in label: d["型式"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        n = extract_number(val)
                        d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "容量" in label: d["容量"] = extract_number(val)
                    if any(k in label for k in ["利用率", "負載率"]):
                        n = extract_number(val)
                        d["負載率"] = n * 100 if 0 < n < 1 else n
                    if any(k in label for k in ["功因", "PF"]):
                        n = extract_number(val)
                        d["現況功因"] = n / 100 if n > 1 else n
                    
                    # 抓取損耗基礎值以利計算
                   
                    specs.append((label, val))
                
                if d["容量"] > 0:
                    # 機齡過濾
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                   # --- 修改後的損耗計算公式 ---
                    # 1. 自動計算基礎損耗 (根據容量估算)
                    base_iron_loss = d["容量"] * 2.5    # 鐵損估算值 (W)
                    base_copper_loss = d["容量"] * 13.0 # 滿載銅損估算值 (W)
                    
                    # 2. 實際銅損 = 滿載銅損 * (負載率/100)^2
                    d["實際銅損"] = base_copper_loss * ((d["負載率"] / 100) ** 2)
                    
                    # 3. 鐵損 (固定值)
                    d["鐵損"] = base_iron_loss
                    
                    # 4. 改善前耗能 (kWh/年) = (鐵損 + 實際銅損) * 8760 / 1000
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})

    if all_transformer_data:
        # --- 2. 網頁即時統計摘要 ---
        caps = [t["analysis"]["容量"] for t in all_transformer_data]
        total_cap = sum(caps)
        cap_dist = Counter(caps)
        avg_usage = sum(t["analysis"]["負載率"] for t in all_transformer_data) / len(all_transformer_data)

        st.success(f"✅ 解析完成！符合篩選條件共 {len(all_transformer_data)} 台")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("1. 總裝置容量", f"{total_cap:,.0f} kVA")
            st.write("**2. 規格台數分布：**")
            for k, v in sorted(cap_dist.items(), reverse=True):
                st.write(f"　🔹 {k:,.0f} kVA × {v} 台")
        with c2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # --- 4. Word 報告產出 ---
        doc = Document()
        
        # 壹、設備統計總表
        doc.add_heading('壹、 設備統計總表', 1)
        st_table = doc.add_table(rows=0, cols=2); st_table.style = 'Table Grid'
        def add_sum(l, v):
            r = st_table.add_row().cells
            set_font_kai(r[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(r[1].paragraphs[0].add_run(v), 12)
        add_sum("總裝置容量", f"{total_cap:,.0f} kVA")
        add_sum("設備規格分布", "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_dist.items(), reverse=True)]))
        add_sum("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、改善前數據分析表 (11欄)
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_table = doc.add_table(rows=1, cols=11); ana_table.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_table.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        for t in all_transformer_data:
            d = t["analysis"]
            row = ana_table.add_row().cells
            row_vals = [d["建築物"], d["編號"], d["年份"], d["廠牌"], f"{d['容量']:.0f}", d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, v in enumerate(row_vals): set_font_kai(row[i].paragraphs[0].add_run(str(v)), 8)
        doc.add_page_break()

        # 參、詳細設備數據 (每台一頁)
        doc.add_heading('參、 詳細設備數據', 1)
        for t in all_transformer_data:
            doc.add_paragraph().add_run(f"設備詳細資料 (編號：{t['analysis']['編號']})").bold = True
            dt = doc.add_table(rows=0, cols=2); dt.style = 'Table Grid'
            for l, v in t["specs"]:
                row = dt.add_row().cells
                set_font_kai(row[0].paragraphs[0].add_run(l), 10)
                set_font_kai(row[1].paragraphs[0].add_run(v), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, f"Transformer_Report.docx")
