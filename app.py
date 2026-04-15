import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器精確報告生成器", layout="wide")
st.title("📑 變壓器自動化報告 (一台一組精確版)")

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    # 讀取 Excel 的所有分頁
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    
    all_transformer_data = []

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        # 1. 定位所有「序號」起點
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_text = str(raw_df.iloc[r, c]).replace(' ', '').replace('\n', '')
                if "序號" == cell_text: # 精確匹配
                    anchors.append((r, c))
        
        if anchors:
            for r_start, c_start in anchors:
                # 橫向處理 (1~6台)
                for offset in range(1, 10):
                    target_col = c_start + offset
                    if target_col >= len(raw_df.columns): break
                    
                    sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                    if sn_val == "nan" or sn_val == "": continue
                    
                    specs = []
                    # 2. 垂直抓取：遇到下一個「序號」字樣就停止
                    for r_offset in range(0, 50): # 最大掃描深度
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        
                        # 檢查是否撞到下一個區塊的開頭
                        label_raw = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                        if r_offset > 0 and "序號" in label_raw:
                            break # 這是別人的資料，停止抓取
                        
                        # 正常抓取標籤與數值
                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        
                        if label == "nan" or not label or "電能系統" in label:
                            continue
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        all_transformer_data.append(specs)

    if all_transformer_data:
        st.success(f"📊 成功分離出 {len(all_transformer_data)} 台獨立設備數據！")

        if st.button("🚀 下載精確版標楷體報告"):
            doc = Document()
            for specs in all_transformer_data:
                # 每個設備給一個清楚的標題
                p = doc.add_paragraph()
                run_t = p.add_run(f"變壓器設備資料 (序號：{specs[0][1]})")
                set_font_kai(run_t, 14, is_bold=True)
                
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                for label, value in specs:
                    cells = table.add_row().cells
                    set_font_kai(cells[0].paragraphs[0].add_run(label), 10)
                    set_font_kai(cells[1].paragraphs[0].add_run(value), 10)
                
                # 強制換頁，確保一組資料一個區塊，不會混在一起
                doc.add_page_break()

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 點我下載精確報告", output, "Transformer_Clean.docx")
