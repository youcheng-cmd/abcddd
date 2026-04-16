import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (最終修正版)")

# --- 1. 參數與查表設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", value=95)
age_filter = st.sidebar.selectbox("選擇變壓器齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

# 鐵損 (Wfe) 與 滿載銅損 (Wcu) 查表 (單位: kW)
IRON_MAP = {150: 0.9, 200: 1.15, 300: 1.57, 400: 1.99, 500: 2.36, 600: 2.75, 750: 3.34, 1000: 4.2, 1250: 5.25, 1500: 5.0, 2000: 6.3, 2500: 7.36, 3000: 8.83}
COPPER_MAP = {150: 2.71, 200: 3.45, 300: 4.71, 400: 5.96, 500: 7.065, 600: 8.25, 750: 10.02, 1000: 12.58, 1250: 15.75, 1500: 20.17, 2000: 25.2, 2500: 29.44, 3000: 35.31}

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
    raw_df = pd.read_excel(excel_file, sheet_name=0, header=None)
    all_transformer_data = []
    seen_sn = set()

    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if str(raw_df.iloc[r, c]).replace(' ', '') == "序號":
                if r + 1 < len(raw_df):
                    next_v = str(raw_df.iloc[r+1, c]).replace(' ', '')
                    if any(k in next_v for k in ["建築", "編號", "位置"]):
                        anchors.append((r, c))
    
    for r_start, c_start in anchors:
        for offset in range(1, 10):
            target_col = c_start + offset
            if target_col >= len(raw_df.columns): break
            
            d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0.0, "型式": "-", "負載率": 0.0, "現況功因": 0.0}
            specs = []
            is_valid = False
            
            for r_offset in range(0, 45):
                curr_r = r_start + r_offset
                if curr_r >= len(raw_df): break
                l1 = str(raw_df.iloc[curr_r, c_start]).strip()
                l2 = str(raw_df.iloc[curr_r, c_start-1]).strip() if c_start > 0 else ""
                label = l1 if (l1 != "nan" and l1 != "") else l2
                lp = label.replace(' ', '').replace('\n', '')
                if r_offset > 0 and lp == "序號": break
                val = str(raw_df.iloc[curr_r, target_col]).strip()
                if lp == "nan" or not lp or val == "nan": continue

                if any(k in lp for k in ["建築", "位置"]): d["建築物"] = val
                if "編號" in lp: d["編號"] = val; is_valid = True
                if "廠牌" in lp: d["廠牌"] = val
                if "型式" in lp: d["型式"] = val
                if any(k in lp for k in ["年份", "出廠"]):
                    n = extract_number(val)
                    d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                if "容量" in lp: d["容量"] = extract_number(val)
                if any(k in lp for k in ["利用率", "負載率"]):
                    n = extract_number(val)
                    d["負載率"] = n * 100 if 0 < n < 1 else n
                if lp == "功因" or any(k in lp for k in ["功率因數", "PF"]):
                    n_pf = extract_number(val)
                    if n_pf > 0: d["現況功因"] = n_pf / 100 if n_pf > 1 else n_pf
                specs.append((label, val))

            if is_valid and d["容量"] > 0:
                ukey = f"{d['建築物']}_{d['編號']}"
                if ukey not in seen_sn:
                    if d["現況功因"] <= 0: d["現況功因"] = 0.8
                    cap = d["容量"]
                    d["鐵損"] = IRON_MAP.get(cap, cap * 3.5) * 1000
                    full_cu = COPPER_MAP.get(cap, cap * 13.0) * 1000
                    d["實際銅損"] = full_cu * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})
                    seen_sn.add(ukey)

    if all_transformer_data:
        total_cap = sum(t["analysis"]["容量"] for t in all_transformer_data)
        cap_counts = Counter(t["analysis"]["容量"] for t in all_transformer_data)
        avg_usage = sum(t["analysis"]["負載率"] for t in all_transformer_data) / len(all_transformer_data)
        total_kwh_before = sum(t["analysis"]["改善前耗能"] for t in all_transformer_data)

        st.success(f"✅ 解析完成！符合篩選條件：共 {len(all_transformer_data)} 台")
        
        doc = Document()
        # 壹、貳、參 略 (保持原本結構)
        doc.add_heading('壹、 設備統計總表', 1)
        st_t = doc.add_table(rows=0, cols=2); st_t.style = 'Table Grid'
        def add_sum(l, v):
            r = st_t.add_row().cells
            set_font_kai(r[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(r[1].paragraphs[0].add_run(v), 12)
        add_sum("總裝置容量", f"{total_cap:,.0f} kVA")
        dist_str = "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True)])
        add_sum("設備規格分布", dist_str)
        doc.add_page_break()

        # 肆、 報告部分 (修正變數與紅字)
        doc.add_heading('肆、 節能改善建議報告', 1)
        savings_kwh = total_kwh_before * 0.65
        savings_money = savings_kwh * 3.3 / 10000 # 萬元
        invest_cost = total_cap * 1600 / 10000 # 萬元
        payback = invest_cost / savings_money if savings_money > 0 else 0

        doc.add_heading('一、 現況說明', 2)
        p1 = doc.add_paragraph()
        p1.add_run("1. 依據能源申報資料，總裝置容量達 ")
        p1.add_run(f"{total_cap:,.0f} kVA").font.color.rgb = RGBColor(255, 0, 0)
        p1.add_run("，現況使用 20 年以上。")
        
        p2 = doc.add_paragraph()
        p2.add_run("2. 評估 ")
        p2.add_run(dist_str).font.color.rgb = RGBColor(255, 0, 0)
        p2.add_run(f" 設備平均利用率 {avg_usage:.1f}%，年耗能約 ")
        p2.add_run(f"{total_kwh_before:,.0f} kWh/年").font.color.rgb = RGBColor(255, 0, 0)

        doc.add_heading('三、 預期效益', 2)
        p3 = doc.add_paragraph()
        p3.add_run(f"預估節電 {savings_kwh:,.0f} kWh/年，省下約 ")
        p3.add_run(f"{savings_money:.1f} 萬元/年").font.color.rgb = RGBColor(255, 0, 0)
        
        p4 = doc.add_paragraph()
        p4.add_run(f"回收年限約 ")
        p4.add_run(f"{payback:.1f} 年").font.color.rgb = RGBColor(255, 0, 0)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, "Transformer_Final_Report.docx")
