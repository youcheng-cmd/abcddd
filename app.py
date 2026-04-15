import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器全量報告生成器", layout="wide")
st.title("📑 變壓器全自動化報告 (全域掃描版)")

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
    # 讀取全部資料，確保連空白格都讀進來
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 全域掃描：尋找包含「序」和「號」的任何儲存格
    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            cell_text = str(raw_df.iloc[r, c])
            # 只要格子裡有「序」和「號」兩個字，不論位置
            if "序" in cell_text and "號" in cell_text:
                # 確認右邊那格是不是數字 1 或其他數字，確保這真的是資料起點
                next_val = str(raw_df.iloc[r, c+1]).strip()
                if next_val != "nan" and next_val != "":
                    anchors.append((r, c))
    
    if not anchors:
        st.error("❌ 系統掃描了整張表還是找不到『序號』儲存格，請檢查 Excel 文字是否正確。")
    else:
        st.success(f"🔍 成功定位到 {len(anchors)} 個資料區塊起點！")
        
        all_transformer_data = []

        # 2. 遍歷起點抓取資料
        for r_start, c_start in anchors:
            # 往右抓取 6 欄 (1~6 台)
            for offset in range(1, 10):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                
                # 如果這格是空的代表這排結束了
                if sn_val == "nan" or sn_val == "":
                    continue
                
                specs = []
                # 垂直抓取標籤與數據 (向下抓 40 行)
                for r_offset in range(0, 40):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 優先抓取與「序號」同欄的標籤
                    label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    # 如果該格沒字，往左看一格 (處理合併儲存格)
                    if (label == "nan" or label == "") and c_start > 0:
                        label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                    
                    value = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    if label == "nan" or not label or "電能系統資料" in label:
                        continue
                    
                    specs.append((label, value if value != "nan" else "-"))
                
                if len(specs) > 5:
                    all_transformer_data.append(specs)

        st.info(f"📊 總計成功解析出 {len(all_transformer_data)} 台變壓器數據！")

        if len(all_transformer_data) > 0:
            if st.button("🚀 下載標楷體專業報告"):
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
                st.download_button("📥 下載報告", output, "Transformer_Report.docx")
