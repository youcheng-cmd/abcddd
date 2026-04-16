import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (數量精確版)")

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
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取 Excel
    raw_df = pd.read_excel(excel_file, sheet_name=0, header=None)
    all_transformer_data = []
    seen_sn = set() # 儲存已抓取的變壓器編號

    # 定位「序號」座標
    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if str(raw_df.iloc[r, c]).replace(' ', '') == "序號":
                # 檢查下方是否為規格表的特徵標籤
                if r + 1 < len(raw_df):
                    next_v = str(raw_df.iloc[r+1, c]).replace(' ', '')
                    if any(k in next_v for k in ["建築", "編號", "位置"]):
                        anchors.append((r, c))
    
    # 遍歷所有找到的規格表區塊
    for r_start, c_start in anchors:
        # 橫向掃描設備 (精確掃描 TR-1 到 TR-7)
        for offset in range(1, 10): 
            target_col = c_start + offset
            if target_col >= len(raw_df.columns): break
            
            # 初始化設備字典
            d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0.0, 
                 "型式": "-", "負載率": 0.0, "現況功因": 0.0, "鐵損": 0.0, "實際銅損": 0.0, "改善前耗能": 0.0}
            specs = []
            
            # 垂直掃描該欄位的內容
            is_valid_device = False
            for r_offset in range(0, 45):
                curr_r = r_start + r_offset
                if curr_r >= len(raw_df): break
                
                # 標籤偵測
                l1 = str(raw_df.iloc[curr_r, c_start]).strip()
                l2 = str(raw_df.iloc[curr_r, c_start-1]).strip() if c_start > 0 else ""
                label = l1 if (l1 != "nan" and l1 != "") else l2
                lp = label.replace(' ', '').replace('\n', '')
                
                if r_offset > 0 and lp == "序號": break
                
                val = str(raw_df.iloc[curr_r, target_col]).strip()
                if lp == "nan" or not lp or val == "nan": continue

                # 資料分類抓取
                if any(k in lp for k in ["建築", "位置"]): d["建築物"] = val
                if "編號" in lp: 
                    d["編號"] = val
                    is_valid_device = True # 只要有編號就視為潛在設備
                if "廠牌" in lp: d["廠牌"] = val
                if "型式" in lp: d["型式"] = val
                if any(k in lp for k in ["年份", "出廠"]):
                    n = extract_number(val)
                    d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                if "容量" in lp: d["容量"] = extract_number(val)
                if any(k in lp for k in ["利用率", "負載率"]):
                    n = extract_number(val)
                    d["負載率"] = n * 100 if 0 < n < 1 else n
                if lp == "功因" or any(k in lp for k in ["功率因數", "PF", "P.F"]):
                    n_pf = extract_number(val)
                    if n_pf > 0: d["現況功因"] = n_pf / 100 if n_pf > 1 else n_pf

                specs.append((label, val))

            # 最終儲存判斷：必須有容量、有編號，且沒重複抓過
            if is_valid_device and d["容量"] > 0:
                # 建立唯一 Key (建築物+編號)
                unique_key = f"{d['建築物']}_{d['編號']}"
                if unique_key not in seen_sn:
                    # 功因預設
                    if d["現況功因"] <= 0: d["現況功因"] = 0.8
                    
                    # 計算公式
                    d["鐵損"] = d["容量"] * 3
                    d["實際銅損"] = (d["容量"] * 13.0) * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    # 篩選邏輯
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})
                    seen_sn.add(unique_key)

    if all_transformer_data:
        # --- 數據摘要顯示 ---
        total_cap = sum(t["analysis"]["容量"] for t in all_transformer_data)
        cap_counts = Counter(t["analysis"]["容量"] for t in all_transformer_data)
        avg_usage = sum(t["analysis"]["負載率"] for t in all_transformer_data) / len(all_transformer_data)

        st.success(f"✅ 解析完成！符合篩選條件：共 {len(all_transformer_data)} 台")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("1. 總裝置容量", f"{total_cap:,.0f} kVA")
            st.write("**2. 規格台數分布：**")
            for k, v in sorted(cap_counts.items(), reverse=True):
                st.write(f"　🔹 {k:,.0f} kVA × {v} 台")
        with c2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # --- Word 產出邏輯 ---
        doc = Document()
        # 壹、總表
        doc.add_heading('壹、 設備統計總表', 1)
        st_t = doc.add_table(rows=0, cols=2); st_t.style = 'Table Grid'
        def add_sum(l, v):
            r = st_t.add_row().cells
            set_font_kai(r[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(r[1].paragraphs[0].add_run(v), 12)
        add_sum("總裝置容量", f"{total_cap:,.0f} kVA")
        add_sum("設備規格分布", "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True)]))
        add_sum("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、數據分析表
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_t = doc.add_table(rows=1, cols=11); ana_t.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_t.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        for t in all_transformer_data:
            d = t["analysis"]
            row = ana_t.add_row().cells
            row_vals = [d["建築物"], d["編號"], d["年份"], d["廠牌"], f"{d['容量']:.0f}", d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, v in enumerate(row_vals): set_font_kai(row[i].paragraphs[0].add_run(str(v)), 8)
        
        doc.add_page_break()
        # 參、詳細資料
        doc.add_heading('參、 詳細設備數據', 1)
        for t in all_transformer_data:
            doc.add_paragraph().add_run(f"設備詳細資料 (編號：{t['analysis']['編號']})").bold = True
            dt = doc.add_table(rows=0, cols=2); dt.style = 'Table Grid'
            for l, v in t["specs"]:
                row = dt.add_row().cells
                set_font_kai(row[0].paragraphs[0].add_run(l), 10); set_font_kai(row[1].paragraphs[0].add_run(v), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, "Transformer_Report_Final.docx")
