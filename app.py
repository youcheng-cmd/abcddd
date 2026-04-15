import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
import io
from collections import Counter

# --- 基礎設定 ---
st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📊 變壓器節能改善分析 (改善前數據表)")

st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("基準年份：", value=2026)
age_filter = st.sidebar.selectbox("機齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])
power_factor = st.sidebar.slider("設定功因 (PF)：", 0.7, 1.0, 0.8) # 預設 0.8

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("上傳 Excel 資料", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    
    # 掃描邏輯 (延用之前精確抓取版本)
    for sheet_name, raw_df in all_sheets.items():
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if str(raw_df.iloc[r, c]).replace(' ', '') == "序號":
                    # 定位到序號後，橫向抓取設備
                    for offset in range(1, 10):
                        target_col = c + offset
                        if target_col >= len(raw_df.columns): break
                        
                        # 初始化數據字典
                        d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                             "型式": "-", "負載率": 0.0, "鐵損": 0.0, "銅損": 0.0, "滿載銅損": 0.0}
                        
                        specs_list = []
                        valid_device = False
                        
                        for r_off in range(50):
                            curr_r = r + r_off
                            if curr_r >= len(raw_df): break
                            label = str(raw_df.iloc[curr_r, c]).strip()
                            val = str(raw_df.iloc[curr_r, target_col]).strip()
                            
                            if r_off == 0 and val != "nan": 
                                d["編號"] = val
                                valid_device = True
                            
                            # 關鍵欄位提取
                            if "位置" in label or "建築" in label: d["建築物"] = val
                            if "製造年份" in label or "出廠" in label:
                                num = ''.join(filter(str.isdigit, val))
                                if num: d["年份"] = int(num) + 1911 if int(num) < 200 else int(num)
                            if "廠牌" in label: d["廠牌"] = val
                            if "容量" in label:
                                c_num = ''.join(filter(str.isdigit, val))
                                d["容量"] = int(c_num) if c_num else 0
                            if "型式" in label: d["型式"] = val
                            if "負載率" in label or "利用率" in label:
                                u_num = val.replace('%', '')
                                try: d["負載率"] = float(u_num) * 100 if 0 < float(u_num) < 1 else float(u_num)
                                except: pass
                            if "無載損" in label or "鐵損" in label:
                                i_num = ''.join(filter(str.isdigit, val))
                                d["鐵損"] = int(i_num) if i_num else 0
                            if "負載損" in label or "銅損" in label:
                                cu_num = ''.join(filter(str.isdigit, val))
                                d["滿載銅損"] = int(cu_num) if cu_num else 0
                            
                            if label != "nan" and val != "nan":
                                specs_list.append((label, val))
                            if r_off > 0 and "序號" in label: break
                        
                        if valid_device:
                            # 計算計算項
                            age = base_year - d["年份"] if d["年份"] else 0
                            # 過濾邏輯
                            if (age_filter == "超過 10 年" and age < 10) or \
                               (age_filter == "超過 15 年" and age < 15) or \
                               (age_filter == "超過 20 年" and age < 20): continue
                            
                            # 損耗計算
                            d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                            d["輸出功率"] = d["容量"] * power_factor * (d["負載率"]/100)
                            d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                            d["age"] = age
                            d["full_specs"] = specs_list
                            all_transformer_data.append(d)

    if all_transformer_data:
        st.success(f"已整理 {len(all_transformer_data)} 台改善前設備數據")
        
        # 預覽表格
        df_show = pd.DataFrame(all_transformer_data).drop(columns=['full_specs'])
        st.dataframe(df_show)

        if st.button("🚀 下載包含「改善前數據表」的報告"):
            doc = Document()
            
            # 1. 插入改善前數據總表
            doc.add_heading('壹、 變壓器設備改善前數據分析表', 1)
            table = doc.add_table(rows=1, cols=11)
            table.style = 'Table Grid'
            
            headers = ["建築物", "編號", "年份", "廠牌", "容量(kVA)", "型式", "負載率%", "輸出功率(kW)", "銅損(W)", "鐵損(W)", "改善前耗能(kWh/年)"]
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                set_font_kai(hdr_cells[i].paragraphs[0].add_run(h), 9, True)
            
            for d in all_transformer_data:
                row_cells = table.add_row().cells
                row_cells[0].text = str(d["建築物"])
                row_cells[1].text = str(d["編號"])
                row_cells[2].text = str(d["年份"])
                row_cells[3].text = str(d["廠牌"])
                row_cells[4].text = str(d["容量"])
                row_cells[5].text = str(d["型式"])
                row_cells[6].text = f"{d['負載率']:.1f}%"
                row_cells[7].text = f"{d['輸出功率']:.2f}"
                row_cells[8].text = f"{d['實際銅損']:.1f}"
                row_cells[9].text = f"{d['鐵損']:.1f}"
                row_cells[10].text = f"{d['改善前耗能']:.0f}"
                for cell in row_cells:
                    for p in cell.paragraphs:
                        set_font_kai(p.runs[0] if p.runs else p.add_run(""), 8)

            doc.add_page_break()
            # (後續接原本的詳細數據表...)
            
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載分析報告", output, "Transformer_Analysis.docx")
