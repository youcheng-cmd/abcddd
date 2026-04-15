import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# 設定網頁標題
st.set_page_config(page_title="變壓器報告生成器", layout="wide")
st.title("📑 變壓器數據自動化報告 (標楷體/黑字版)")

def set_font_kai(run, size=12):
    """設定字體為標楷體與黑色"""
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0) # 強制黑色
    # 針對中文字體的特殊設定
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 尋找「序號」所在的座標
    sn_row, sn_col = None, None
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if "序號" in str(raw_df.iloc[r, c]):
                sn_row, sn_col = r, c
                break
        if sn_row is not None: break

    if sn_row is not None:
        # 2. 抓取所有序號欄位 (橫向往右找 1, 2, 3...)
        transformer_cols = []
        for c in range(sn_col + 1, len(raw_df.columns)):
            val = raw_df.iloc[sn_row, c]
            if pd.notna(val):
                transformer_cols.append(c)
        
        st.success(f"✅ 偵測到 {len(transformer_cols)} 組變壓器資料")

        if st.button("🚀 產出標楷體報告"):
            doc = Document()
            
            # 遍歷每一組變壓器，分別建立一個整齊的區塊
            for idx, col_idx in enumerate(transformer_cols):
                # 新增標題
                p = doc.add_paragraph()
                run = p.add_run(f"變壓器資料組 {idx + 1}")
                set_font_kai(run, 16)
                run.bold = True
                
                # 建立一個簡單的 2 欄表格 (左邊標題，右邊數據)
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                
                # 垂直抓取該序號下的所有資料
                # 我們抓取從「序號」列開始往下的 20 行資料
                for r in range(sn_row, sn_row + 25): 
                    if r >= len(raw_df): break
                    
                    label = str(raw_df.iloc[r, sn_col]).replace('\n', '').strip()
                    value = str(raw_df.iloc[r, col_idx]).strip()
                    
                    if label == "nan" or not label: continue
                    
                    # 新增表格列
                    row_cells = table.add_row().cells
                    
                    # 左側標題
                    run_l = row_cells[0].paragraphs[0].add_run(label)
                    set_font_kai(run_l, 12)
                    
                    # 右側數據
                    run_r = row_cells[1].paragraphs[0].add_run(value if value != "nan" else "-")
                    set_font_kai(run_r, 12)
                
                # 組與組之間空一行
                doc.add_paragraph()

            # 存檔下載
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載標楷體報告", output, "Transformer_Report_Kai.docx")
    else:
        st.error("❌ 找不到『序號』起始點，請檢查 Excel。")
