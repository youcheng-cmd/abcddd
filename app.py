import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器專業報告-穩定版", layout="wide")
st.title("📑 變壓器自動化報告 (穩定下載版)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 設定基準與篩選")
base_year = st.sidebar.number_input("請輸入基準年份：", min_value=1900, max_value=2100, value=2026)
age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年 (汰換參考)", "超過 15 年 (優先汰換)", "超過 20 年 (屆齡汰換)"]
)

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機", "電能系統"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_text = str(raw_df.iloc[r, c]).replace(' ', '').replace('\n', '')
                if cell_text == "序號":
                    anchors.append((r, c))
        
        if anchors:
            for r_start, c_start in anchors:
                for offset in range(1, 10):
                    target_col = c_start + offset
                    if target_col >= len(raw_df.columns): break
                    sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                    if sn_val == "nan" or sn_val == "": continue
                    
                    specs = []
                    mfg_year, capacity, usage_rate = None, 0, None
                    
                    for r_offset in range(0, 50):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        
                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        label_clean = label.replace(' ', '')
                        if r_offset > 0 and label_clean == "序號": break
                        if any(k in label_clean for k in filter_keywords): continue
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label: continue
                        
                        # 數據提取
                        if any(k in label for k in ["製造年份", "製造日期", "出廠", "年份"]):
                            digits = ''.join(filter(str.isdigit, value))
                            if digits:
                                y = int(digits)
                                mfg_year = y + 1911 if y < 200 else y
                        
                        if "容量" in label:
                            cap_digits = ''.join(filter(str.isdigit, value))
                            if cap_digits: capacity = int(cap_digits)
                        
                        if any(k in label for k in ["利用率", "負載率"]):
                            u_raw = value.replace('%', '').strip()
                            try:
                                u_val = float(u_raw)
                                usage_rate = u_val * 100 if 0 < u_val < 1 else u_val
                            except: pass
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        age = base_year - mfg_year if mfg_year else 0
                        should_add = False
                        if age_filter == "顯示全部": should_add = True
                        elif "10" in age_filter and age >= 10: should_add = True
                        elif "15" in age_filter and age >= 15: should_add = True
                        elif "20" in age_filter and age >= 20: should_add = True
                        
                        if should_add:
                            all_transformer_data.append({"specs": specs, "capacity": capacity, "usage_rate": usage_rate, "age": age})

    if all_transformer_data:
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        valid_usages = [t["usage_rate"] for t in all_transformer_data if t["usage_rate"] is not None and t["usage_rate"] > 0]
        avg_usage = sum(valid_usages) / len(valid_usages) if valid_usages else 0

        st.success(f"✅ 修正完成！平均負載利用率：{avg_usage:.1f} %")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("總裝置容量", f"{total_capacity} kVA")
        col2.metric("平均負載利用率", f"{avg_usage:.1f} %")
        col3.metric("篩選後總台數", f"{len(all_transformer_data)} 台")

        # --- 核心修復：先準備好 Word 資料再顯示按鈕 ---
        doc = Document()
        # 壹、統計總表
        doc.add_heading('壹、設備統計總表', 1)
        summary_table = doc.add_table(rows=0, cols=2)
        summary_table.style = 'Table Grid'
        
        def add_sum_row(t, l, v):
            row = t.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(v), 12)

        add_sum_row(summary_table, "總裝置容量", f"{total_capacity} kVA")
        cap_dist = "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True) if k > 0])
        add_sum_row(summary_table, "設備規格分布", cap_dist if cap_dist else "-")
        add_sum_row(summary_table, "平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、詳細資料
        doc.add_heading('貳、詳細設備資料', 1)
        for item in all_transformer_data:
            specs = item["specs"]
            p = doc.add_paragraph()
            run_t = p.add_run(f"變壓器設備資料 (序號：{specs[0][1]})")
            set_font_kai(run_t, 14, is_bold=True)
            table = doc.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            for label, value in specs:
                cells = table.add_row().cells
                set_font_kai(cells[0].paragraphs[0].add_run(label), 10)
                set_font_kai(cells[1].paragraphs[0].add_run(value), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        # 改用這種方式：不讓按鈕去觸發生成，而是資料一上傳就生成好，按鈕只負責下載
        st.download_button(
            label="📥 點我下載正式報告",
            data=output,
            file_name=f"Transformer_Report_{base_year}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
