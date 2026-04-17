import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io

# --- 字體設定函數 ---
def set_font_kai(run, size=12, is_bold=False):
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.name = '標楷體'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

st.title("📄 案場基本資料修正")

# --- 1. 用戶簡介 ---
st.header("2-1. 用戶簡介")
col1, col2 = st.columns(2)
with col1:
    company = st.text_input("公司名稱", "誠友開發股份有限公司")
    area = st.text_input("總建物面積 (m²)", "43,614")
with col2:
    building = st.text_input("建物名稱", "大墩食衣生活廣場")
    staff = st.text_input("員工編制 (人)", "10")

# --- 2. 電力系統表格 (對照你的照片 1) ---
st.header("1. 電力系統規格")
power_data = {
    "台電電號": ["07699050101"],
    "契約容量 (kW)": ["1,100"],
    "供電電壓 (kV)": ["22.8"],
    "平均單價 (元/kWh)": ["4.48"]
}
df_power = pd.DataFrame(power_data)
st.data_editor(df_power, key="power_edit")

# --- 3. 照明系統表格 (對照你的照片 2) ---
st.header("2. 照明系統設備")
# 預設一些你照片中的數據
lighting_data = [
    {"燈具種類": "日光燈", "規格": "14W*4", "數量": 40, "運轉時數": 4380},
    {"燈具種類": "日光燈", "規格": "28W*2", "數量": 645, "運轉時數": 4380},
    {"燈具種類": "LED", "規格": "12W*1", "數量": 385, "運轉時數": 4380},
]
df_lighting = pd.DataFrame(lighting_data)
edited_lighting = st.data_editor(df_lighting, num_rows="dynamic", key="light_edit")

# --- 4. 生成 Word 報告 ---
if st.button("📝 產出基本資料 Word"):
    doc = Document()
    
    # 章節標題
    h = doc.add_paragraph()
    set_font_kai(h.add_run("二、 基本資料"), size=16, is_bold=True)
    
    # 用戶簡介
    p = doc.add_paragraph()
    set_font_kai(p.add_run("2-1. 用戶簡介"), size=14, is_bold=True)
    
    p2 = doc.add_paragraph()
    set_font_kai(p2.add_run(f"{company}({building}) 總建物面積 {area} 平方公尺，員工約 {staff} 人。"))

    # 產出照明表格
    doc.add_paragraph().add_run("照明系統明細：").bold = True
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '種類'
    hdr_cells[1].text = '規格'
    hdr_cells[2].text = '數量'
    hdr_cells[3].text = '時數'

    for _, row in edited_lighting.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['燈具種類'])
        row_cells[1].text = str(row['規格'])
        row_cells[2].text = str(row['數量'])
        row_cells[3].text = str(row['運轉時數'])

    buffer = io.BytesIO()
    doc.save(buffer)
    st.download_button("💾 下載 Word", buffer.getvalue(), "Basic_Info.docx")
