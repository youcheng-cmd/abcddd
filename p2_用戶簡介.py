import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io

# 設定字體為標楷體的功能
def set_font_kai(run, size=12, is_bold=False):
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

st.title("📄 案場基本資料與用戶簡介")

# 1. 用戶簡介輸入
st.header("2-1. 用戶簡介")
col1, col2 = st.columns(2)
with col1:
    company = st.text_input("公司名稱", "誠友開發股份有限公司")
    address = st.text_input("建物地址", "台中市南屯區...")
with col2:
    building = st.text_input("建物名稱", "大墩食衣生活廣場")
    area = st.text_input("總建物面積 (m²)", "43,614")

# 2. 電力系統表格 (對應你提供的圖片內容)
st.header("1. 電力系統規格")
power_df = pd.DataFrame([
    {"項目": "台電電號", "規格內容": "07699050101"},
    {"項目": "契約容量 (kW)", "規格內容": "1,100"},
    {"項目": "年用電量 (kWh)", "規格內容": "4,239,400"},
    {"項目": "平均電費單價", "規格內容": "4.48"}
])
edited_power = st.data_editor(power_df, num_rows="dynamic", key="power_table")

# 3. 生成 Word 邏輯
if st.button("🚀 生成基本資料 Word 報告"):
    doc = Document()
    
    # 標題
    h = doc.add_paragraph()
    set_font_kai(h.add_run("二、 基本資料"), size=16, is_bold=True)
    
    # 用戶簡介文字
    p1 = doc.add_paragraph()
    set_font_kai(p1.add_run("2-1. 用戶簡介"), size=14, is_bold=True)
    
    p2 = doc.add_paragraph()
    intro_text = f"{company}({building})總建物面積約 {area} 平方公尺，能源使用主要以電力為主。"
    set_font_kai(p2.add_run(intro_text))
    
    # 寫入電力系統表格
    doc.add_paragraph().add_run("1. 電力系統：").bold = True
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '項目'
    hdr_cells[1].text = '規格內容'
    
    for index, row in edited_power.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['項目'])
        row_cells[1].text = str(row['規格內容'])

    # 下載
    buffer = io.BytesIO()
    doc.save(buffer)
    st.download_button("💾 下載基本資料報告", buffer.getvalue(), "用戶簡介.docx")
