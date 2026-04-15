import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="變壓器大數據報告生成器", layout="wide")
st.title("📑 變壓器全自動化報告 (支援 15+ 台數)")

def set_font_kai(run, size=12, is_bold=False):
    """設定中文字體為標楷體，英文字體為標楷體，顏色為純黑"""
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    # 讀取全部資料，不設標題
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 尋找所有包含「序號」字樣的座標作為「區塊起點」
    anchors = []
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if "序號" in str(raw_df.iloc[r, c]):
                anchors.append((r, c))
    
    if not anchors:
        st.error("❌ 找不到『序號』起始點，請確認 Excel 內容。")
    else:
        all_transformer_data = []

        # 2. 處理每一個偵測到的區塊
        for r_start, c_start in anchors:
            # 在每個區塊中，往右看 6 欄 (1~6 台)
            for offset in range(1, 7):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                # 取得該欄位的序號值 (例如 1, 2, 3...)
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                
                # 如果這格是空的代表這組結束了
                if sn_val == "nan" or sn_val == "":
                    continue
                
                # 垂直抓取該台變壓器的參數
                specs = []
                # 往下抓 35 行內容
                for r_offset in range(0, 35):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    value = str(raw_df.iloc[curr_r, target_col]).strip()
                    
                    if label == "nan" or not label: continue
                    specs.append((label, value if value != "nan" else "-"))
                
                if specs:
                    all_transformer_data.append(specs)

        st.success(f"✅ 掃描完成！總共偵測到 {len(all_transformer_data)} 組變壓器數據。")

        # 3. 產出 Word
        if len(all_transformer_data) > 0:
            if st.button("🚀 生成標楷體專業報告"):
                doc = Document()
                
                # 遍歷所有機器資料
                for i, specs in enumerate(all_transformer_data):
                    # 標題
                    p = doc.add_paragraph()
                    run_t = p.add_run(f"設備數據報告 - 序號 {specs[0][1]}")
                    set_font_kai(run_t, 14, is_bold=True)
                    
                    # 建立 2 欄表格
                    table = doc.add_table(rows=0, cols=2)
                    table.style = 'Table Grid'
                    
                    for label, value in specs:
                        row_cells = table.add_row().cells
                        # 左側標籤
                        set_font_kai(row_cells[0].paragraphs[0].add_run(label), 10)
                        # 右側數值
                        set_font_kai(row_cells[1].paragraphs[0].add_run(value), 10)
                    
                    # 組間隔
                    doc.add_paragraph()

                # 輸出
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                st.download_button(f"📥 下載共 {len(all_transformer_data)} 台報告", output, "Full_Report.docx")
