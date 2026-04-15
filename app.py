import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器全自動報告-終極版", layout="wide")
st.title("📑 變壓器全自動化報告 (全分頁掃描版)")

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    # 1. 讀取 Excel 的「所有」分頁
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    
    found_data = False
    all_transformer_data = []

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        # 2. 全域掃描該分頁
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                # 強制清洗字元：去掉空格、換行、括號
                cell_text = str(raw_df.iloc[r, c]).replace(' ', '').replace('\n', '').replace('\r', '')
                if "序號" in cell_text:
                    # 確認右邊有資料
                    if c + 1 < len(raw_df.columns):
                        anchors.append((r, c))
        
        if anchors:
            found_data = True
            st.info(f"📍 在分頁 [{sheet_name}] 找到 {len(anchors)} 個資料區塊")
            
            for r_start, c_start in anchors:
                # 橫向抓取 (1~6台)
                for offset in range(1, 10):
                    target_col = c_start + offset
                    if target_col >= len(raw_df.columns): break
                    
                    sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                    if sn_val == "nan" or sn_val == "": continue
                    
                    specs = []
                    # 垂直抓取 (向下抓 45 行)
                    for r_offset in range(0, 45):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        
                        # 標籤定位 (處理合併儲存格)
                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').replace(' ', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').replace(' ', '').strip()
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label or "電能系統" in label: continue
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if len(specs) > 5:
                        all_transformer_data.append(specs)

    if not found_data:
        st.error("❌ 系統翻遍了所有分頁還是找不到『序號』儲存格。")
        st.warning("💡 診斷建議：請確認 Excel 裡的『序號』這兩個字沒有寫錯，或試著把該分頁移到最前面。")
    else:
        st.success(f"📊 總計跨分頁解析出 {len(all_transformer_data)} 台變壓器數據！")

        if len(all_transformer_data) > 0:
            if st.button("🚀 下載標楷體報告"):
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
                    doc.add_paragraph()

                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                st.download_button("📥 點我下載報告", output, "Transformer_Final.docx")
