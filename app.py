import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器分段報告生成器", layout="wide")
st.title("📑 變壓器分能自動化報告 (支援分頁/分段抓取)")

def set_font_kai(run, size=12, is_bold=False):
    """設定標楷體、黑字"""
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取 Excel 原始資料 (不設標題)
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 找出所有「序號」出現的座標 (支援 A4 分頁跳行)
    anchor_points = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if "序號" in str(raw_df.iloc[r, c]):
                anchor_points.append((r, c))
    
    if anchor_points:
        st.success(f"✅ 偵測到 {len(anchor_points)} 個數據區塊")
        
        all_transformers = []
        
        # 2. 針對每個「序號」區塊進行掃描
        for start_row, start_col in anchor_points:
            # 橫向找序號後的機器 (通常是 6 台)
            current_block_cols = []
            for c in range(start_col + 1, len(raw_df.columns)):
                val = raw_df.iloc[start_row, c]
                if pd.notna(val) and str(val).strip() != "":
                    current_block_cols.append(c)
                else:
                    # 遇到空白代表這組 (6台) 結束
                    break
            
            # 垂直抓取該區塊內的每一台機器
            for col_idx in current_block_cols:
                transformer_info = []
                # 抓取序號下方約 35 行的內容
                for r in range(start_row, start_row + 35):
                    if r >= len(raw_df): break
                    label = str(raw_df.iloc[r, start_col]).replace('\n', '').strip()
                    value = str(raw_df.iloc[r, col_idx]).strip()
                    
                    if label == "nan" or not label: continue
                    transformer_info.append((label, value if value != "nan" else "-"))
                
                all_transformers.append(transformer_info)
        
        st.info(f"📊 總計抓取到 {len(all_transformers)} 台變壓器數據")

        if st.button("🚀 產出完整標楷體報告"):
            doc = Document()
            
            for idx, info in enumerate(all_transformers):
                # 取得該組的序號數字 (通常在 info 的第一項)
                sn_val = info[0][1] if info else str(idx+1)
                
                p = doc.add_paragraph()
                run_title = p.add_run(f"變壓器資料區塊 (序號：{sn_val})")
                set_font_kai(run_title, 14, is_bold=True)
                
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                
                for label, value in info:
                    row_cells = table.add_row().cells
                    # 左標籤
                    set_font_kai(row_cells[0].paragraphs[0].add_run(label), 11)
                    # 右數值
                    set_font_kai(row_cells[1].paragraphs[0].add_run(value), 11)
                
                doc.add_paragraph() # 組間空格

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button(f"📥 下載全部 {len(all_transformers)} 台報告", output, "Transformer_Full_Report.docx")
    else:
        st.error("❌ 無法在 Excel 中找到任何『序號』起始標記。")
