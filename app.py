import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器 41 台全量報告生成器", layout="wide")
st.title("📑 變壓器全自動化報告 (41台全量修正版)")

def set_font_kai(run, size=11, is_bold=False):
    """設定標楷體與純黑字"""
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    # 讀取全部資料
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 超級模糊搜尋起點：只要儲存格包含「序」且「號」在附近，或者直接包含「序號」
    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            cell_text = str(raw_df.iloc[r, c])
            if "序" in cell_text and r < len(raw_df) - 5: # 確保後面還有資料
                # 再次確認這是不是標題行（右邊格子通常會有 1, 2, 3...）
                next_val = str(raw_df.iloc[r, c+1]).strip()
                if next_val.isdigit() or (next_val != "nan" and len(next_val) < 4):
                    # 避免重複抓取同一個位置
                    if not any(a[0] == r for a in anchors):
                        anchors.append((r, c))
    
    if not anchors:
        st.error("❌ 還是找不到起點，請確認 Excel 第一欄是否有『序號』字樣。")
    else:
        all_transformer_data = []

        # 2. 處理每一個偵測到的區塊
        for r_start, c_start in anchors:
            # 橫向掃描：從序號格子往右看，最多看 10 欄 (確保 6 台都能抓到)
            for offset in range(1, 11):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                # 取得設備序號 (例如 37, 38...)
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                
                # 跳過空白或不具代表性的格子
                if sn_val == "nan" or sn_val == "" or len(sn_val) > 5:
                    continue
                
                # 垂直抓取該台變壓器的所有標籤與參數
                specs = []
                # 往下抓 40 行，確保涵蓋到「裝置電容器容量」
                for r_offset in range(0, 40):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 抓取標籤 (例如：變壓器容量)
                    label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    # 如果標籤是 nan，嘗試往左看一格（處理合併儲存格標籤偏左的問題）
                    if label == "nan" or label == "":
                        if c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()

                    value = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    if label == "nan" or not label or "電能系統資料" in label: 
                        continue
                        
                    specs.append((label, value if value != "nan" else "-"))
                
                if len(specs) > 5: # 確保抓到的是有意義的數據
                    all_transformer_data.append(specs)

        st.success(f"✅ 掃描完成！本次共抓取到 {len(all_transformer_data)} 組變壓器數據。")

        # 3. 產出 Word
        if len(all_transformer_data) > 0:
            if st.button(f"🚀 生成這 {len(all_transformer_data)} 台的專業報告"):
                doc = Document()
                for specs in all_transformer_data:
                    # 找序號數字
                    current_sn = specs[0][1]
                    p = doc.add_paragraph()
                    run_t = p.add_run(f"變壓器設備資料 - 序號 {current_sn}")
                    set_font_kai(run_t, 14, is_bold=True)
                    
                    table = doc.add_table(rows=0, cols=2)
                    table.style = 'Table Grid'
                    
                    for label, value in specs:
                        row_cells = table.add_row().cells
                        set_font_kai(row_cells[0].paragraphs[0].add_run(label), 10)
                        set_font_kai(row_cells[1].paragraphs[0].add_run(value), 10)
                    
                    doc.add_paragraph()

                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                st.download_button("📥 下載全量標楷體報告", output, "Transformer_Full_Report.docx")
