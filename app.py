import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="全自動變壓器報告", layout="wide")
st.title("📑 變壓器數據自動化報告 (支援無限台數)")

def set_font_kai(run, size=12, is_bold=False):
    """設定字體為標楷體、黑色、指定大小"""
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    # 中文字體特殊關鍵詞設定
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 尋找「序號」所在的起始座標
    sn_row, sn_col = None, None
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if "序號" in str(raw_df.iloc[r, c]):
                sn_row, sn_col = r, c
                break
        if sn_row is not None: break

    if sn_row is not None:
        # 2. 自動偵測「所有」序號欄位 (從序號格往右一直找，直到沒數字為止)
        transformer_cols = []
        for c in range(sn_col + 1, len(raw_df.columns)):
            val = raw_df.iloc[sn_row, c]
            # 只要格子裡有內容，就認定是一台變壓器
            if pd.notna(val) and str(val).strip() != "":
                transformer_cols.append(c)
        
        st.success(f"✅ 系統已自動偵測到：共 {len(transformer_cols)} 組變壓器數據")

        if st.button("🚀 產出全數標楷體報告"):
            doc = Document()
            
            # 遍歷偵測到的所有組數 (無論是 6 組還是 15 組)
            for idx, col_idx in enumerate(transformer_cols):
                # 取得該組的序號名稱 (例如 1, 2, 3...)
                sn_name = str(raw_df.iloc[sn_row, col_idx])
                
                # 新增區塊標題
                p = doc.add_paragraph()
                run_title = p.add_run(f"變壓器序號：{sn_name}")
                set_font_kai(run_title, 14, is_bold=True)
                
                # 建立 2 欄表格
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                
                # 垂直抓取該序號下的參數（設定往下抓 30 行，確保所有規格都抓到）
                for r in range(sn_row + 1, sn_row + 35): 
                    if r >= len(raw_df): break
                    
                    label = str(raw_df.iloc[r, sn_col]).replace('\n', '').strip()
                    value = str(raw_df.iloc[r, col_idx]).strip()
                    
                    # 跳過空白標籤
                    if label == "nan" or not label: continue
                    
                    row_cells = table.add_row().cells
                    # 設定左側標題
                    set_font_kai(row_cells[0].paragraphs[0].add_run(label), 12)
                    # 設定右側數據
                    set_font_kai(row_cells[1].paragraphs[0].add_run(value if value != "nan" else "-"), 12)
                
                # 分隔每一組變壓器
                doc.add_paragraph()

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button(f"📥 下載 {len(transformer_cols)} 組數據報告", output, "Full_Report.docx")
    else:
        st.error("❌ 找不到『序號』起始點，請確認 Excel。")
