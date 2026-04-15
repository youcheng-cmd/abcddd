import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器專業報告-最終版", layout="wide")
st.title("📑 變壓器自動化報告 (純淨表格版)")

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

    # 設定過濾關鍵字：只要標籤包含這些字，就不抓進表格
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
                    for r_offset in range(0, 50):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        
                        label_raw = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                        
                        # 停止條件：撞到下一個序號就停
                        if r_offset > 0 and label_raw == "序號":
                            break
                        
                        # 過濾條件：撞到註解關鍵字就跳過這一行，但不一定要停止
                        if any(k in label_raw for k in filter_keywords):
                            continue

                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        
                        # 再次檢查清洗後的標籤
                        if label == "nan" or not label or any(k in label for k in filter_keywords):
                            continue
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        all_transformer_data.append(specs)

    if all_transformer_data:
        st.success(f"📊 成功生成 {len(all_transformer_data)} 台純淨設備數據！")

        if st.button("🚀 下載最終版標楷體報告"):
            doc = Document()
            for specs in all_transformer_data:
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
            st.download_button("📥 點我下載純淨報告", output, "Transformer_Final_Clean.docx")
