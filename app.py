import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📊 變壓器節能改善分析 (改善前數據表)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 設定基準與參數")
# 基準年份
base_year = st.sidebar.number_input("請輸入基準年份：", min_value=1900, max_value=2100, value=2026)
# 功率因數設定 (預設 95)
pf_input = st.sidebar.number_input("設定功率因數 (%)：", min_value=0, max_value=100, value=95)
power_factor = pf_input / 100  # 轉為小數供計算使用

# 機齡篩選
age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"]
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
                    
                    # 初始化該台設備的分析數據
                    d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                         "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0}
                    
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
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label: continue
                        
                        # --- 精確提取分析用數據 ---
                        if "建築物" in label: d["建築物"] = value
                        if "編號" in label: d["編號"] = value
                        if "廠牌" in label: d["廠牌"] = value
                        if "型式" in label: d["型式"] = value
                        if any(k in label for k in ["年份", "出廠"]):
                            digits = ''.join(filter(str.isdigit, value))
                            if digits:
                                y = int(digits)
                                d["年份"] = y + 1911 if y < 200 else y
                        if "容量" in label:
                            cap_digits = ''.join(filter(str.isdigit, value))
                            if cap_digits: d["容量"] = int(cap_digits)
                        if any(k in label for k in ["利用率", "負載率"]):
                            u_raw = value.replace('%', '').strip()
                            try:
                                u_val = float(u_raw)
                                d["負載率"] = u_val * 100 if 0 < u_val < 1 else u_val
                            except: pass
                        if "效率" in label:
                            # 之後若要反推損耗可用
                            pass

                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        age = base_year - d["年份"] if d["年份"] else 0
                        # 篩選邏輯
                        should_add = False
                        if age_filter == "顯示全部": should_add = True
                        elif "10" in age_filter and age >= 10: should_add = True
                        elif "15" in age_filter and age >= 15: should_add = True
                        elif "20" in age_filter and age >= 20: should_add = True
                        
                        if should_add:
                            # --- 損耗與功率計算 ---
                            # 1. 鐵損估計 (假設為容量的 0.2%)
                            d["鐵損"] = d["容量"] * 2.0 
                            # 2. 滿載銅損估計 (假設為容量的 1.2%)
                            d["滿載銅損"] = d["容量"] * 12.0
                            # 3. 實際銅損 = 滿載銅損 * (負載率/100)^2
                            d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                            # 4. 輸出功率 = 容量 * 功因 * (負載率/100)
                            d["輸出功率"] = d["容量"] * power_factor * (d["負載率"]/100)
                            # 5. 改善前年耗能 = (鐵損 + 實際銅損) * 8760 / 1000
                            d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                            
                            all_transformer_data.append({"specs": specs, "analysis": d})

    if all_transformer_data:
        st.success(f"✅ 成功解析 {len(all_transformer_data)} 台設備。基準年份：{base_year}，功因：{pf_input}%")
        
        # 預覽分析表
        analysis_list = [t["analysis"] for t in all_transformer_data]
        st.write("### 📋 改善前分析數據預覽")
        st.dataframe(pd.DataFrame(analysis_list).drop(columns=['鐵損','滿載銅損']))

        if st.button("🚀 生成 Word 報告 (含 11 欄分析表)"):
            doc = Document()
            
            # 1. 改善前分析總表
            doc.add_heading('壹、 變壓器設備改善前數據分析表', 1)
            table = doc.add_table(rows=1, cols=11)
            table.style = 'Table Grid'
            
            headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "輸出功率", "銅損(W)", "鐵損(W)", "改善前耗能"]
            units = ["", "", "", "", "(kVA)", "", "(%)", "(kW)", "", "", "(kWh/年)"]
            
            hdr_cells = table.rows[0].cells
            for i, (h, u) in enumerate(zip(headers, units)):
                p = hdr_cells[i].paragraphs[0]
                set_font_kai(p.add_run(f"{h}\n{u}"), 9, True)
            
            for item in all_transformer_data:
                d = item["analysis"]
                row_cells = table.add_row().cells
                data_row = [
                    d["建築物"], d["編號"], d["年份"], d["廠牌"], d["容量"], d["型式"],
                    f"{d['負載率']:.1f}%", f"{d['輸出功率']:.2f}", f"{d['實際銅損']:.1f}", 
                    f"{d['鐵損']:.1f}", f"{int(d['改善前耗能'])}"
                ]
                for i, val in enumerate(data_row):
                    p = row_cells[i].paragraphs[0]
                    set_font_kai(p.add_run(str(val)), 8)

            doc.add_page_break()
            
            # 2. 詳細設備資料 (原本的內容)
            doc.add_heading('貳、 詳細設備數據', 1)
            for item in all_transformer_data:
                specs = item["specs"]
                p = doc.add_paragraph()
                run_t = p.add_run(f"設備詳細資料 (序號：{specs[0][1]})")
                set_font_kai(run_t, 14, is_bold=True)
                
                table_detail = doc.add_table(rows=0, cols=2)
                table_detail.style = 'Table Grid'
                for label, value in specs:
                    cells = table_detail.add_row().cells
                    set_font_kai(cells[0].paragraphs[0].add_run(label), 10)
                    set_font_kai(cells[1].paragraphs[0].add_run(value), 10)
                doc.add_page_break()

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載完整分析報告", output, f"Transformer_Analysis_{base_year}.docx")
