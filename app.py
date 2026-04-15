import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", min_value=0, max_value=100, value=95)
pf_after = pf_after_input / 100

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
            for offset in range(1, 10):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                if str(raw_df.iloc[r_start, target_col]) in ["nan", ""]: continue
                
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                     "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8}
                specs = []
                
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    if (label == "nan" or label == "") and c_start > 0:
                        label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                    
                    label_clean = label.replace(' ', '')
                    if r_offset > 0 and label_clean == "序號": break
                    if any(k in label_clean for k in filter_keywords): continue
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label == "nan" or not label: continue

                    # 自動數據抓取
                    if "位置" in label or "建築" in label: d["建築物"] = val
                    if "編號" in label: d["編號"] = val
                    if "廠牌" in label: d["廠牌"] = val
                    if "型式" in label: d["型式"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        digits = ''.join(filter(str.isdigit, val))
                        if digits:
                            y = int(digits)
                            d["年份"] = y + 1911 if y < 200 else y
                    if "容量" in label:
                        cap_digits = ''.join(filter(str.isdigit, val))
                        if cap_digits: d["容量"] = int(cap_digits)
                    if any(k in label for k in ["利用率", "負載率"]):
                        u_raw = val.replace('%', '').strip()
                        try:
                            u_val = float(u_raw)
                            d["負載率"] = u_val * 100 if 0 < u_val < 1 else u_val
                        except: pass
                    if any(k in label for k in ["功率因數", "功因", "PF"]):
                        pf_raw = val.replace('%', '').strip()
                        try:
                            pf_val = float(pf_raw)
                            d["現況功因"] = pf_val / 100 if pf_val > 1 else pf_val
                        except: pass
                    if "無載損" in label or "鐵損" in label:
                        i_num = ''.join(filter(str.isdigit, val))
                        d["鐵損"] = int(i_num) if i_num else 0
                    if "負載損" in label or "銅損" in label:
                        cu_num = ''.join(filter(str.isdigit, val))
                        d["滿載銅損"] = int(cu_num) if cu_num else 0

                    specs.append((label, val))
                
                if specs:
                    age = base_year - d["年份"] if d["年份"] else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # 損耗計算
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                    d["現況輸出功率"] = d["容量"] * d["現況功因"] * (d["負載率"]/100)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    all_transformer_data.append({"specs": specs, "analysis": d, "capacity": d["容量"], "usage_rate": d["負載率"]})

    if all_transformer_data:
        # --- 數據摘要顯示區 ---
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        valid_usages = [t["usage_rate"] for t in all_transformer_data if t["usage_rate"] > 0]
        avg_usage = sum(valid_usages) / len(valid_usages) if valid_usages else 0
        
        # 畫面摘要欄位 (找回的功能)
        st.success(f"✅ 解析完成！符合篩選條件：共 {len(all_transformer_data)} 台")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("1. 總裝置容量", f"{total_capacity} kVA")
            st.write("**2. 設備規格分布：**")
            # 顯示如 1000kVA x 5台
            for k, v in sorted(cap_counts.items(), reverse=True):
                if k > 0: st.write(f"　🔹 {k} kVA × {v} 台")
        with col2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")
            st.info(f"💡 改善後目標功因設定：{pf_after_input}%")

        # --- Word 報告生成 ---
        doc = Document()
        
        # 壹、統計總表
        doc.add_heading('壹、 設備統計總表', 1)
        sum_table = doc.add_table(rows=0, cols=2); sum_table.style = 'Table Grid'
        def add_sum_row(l, v):
            row = sum_table.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(v), 12)
        add_sum_row("總裝置容量", f"{total_capacity} kVA")
        cap_dist = "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True) if k > 0])
        add_sum_row("設備規格分布", cap_dist)
        add_sum_row("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、改善前數據分析表 (11欄)
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_table = doc.add_table(rows=1, cols=11); ana_table.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_table.rows[0].cells[i].paragraphs[0].add_run(h), 9, True)
        
        for item in all_transformer_data:
            d = item["analysis"]
            row = ana_table.add_row().cells
            row_data = [d["建築物"], d["編號"], d["年份"], d["廠牌"], d["容量"], d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能'])}"]
            for i, val in enumerate(row_data): set_font_kai(row[i].paragraphs[0].add_run(str(val)), 8)
        doc.add_page_break()

        # 參、詳細設備資料
        doc.add_heading('參、 詳細設備數據', 1)
        for item in all_transformer_data:
            specs = item["specs"]
            p = doc.add_paragraph(); set_font_kai(p.add_run(f"設備詳細資料 (序號：{specs[0][1]})"), 14, True)
            t_detail = doc.add_table(rows=0, cols=2); t_detail.style = 'Table Grid'
            for l, v in specs:
                cells = t_detail.add_row().cells
                set_font_kai(cells[0].paragraphs[0].add_run(l), 10)
                set_font_kai(cells[1].paragraphs[0].add_run(v), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        # 唯一的下載按鈕
        st.download_button(
            label="📥 下載完整報告",
            data=output,
            file_name=f"Transformer_Report_{base_year}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
