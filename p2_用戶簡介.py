import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 通用工具函數：標楷體設定 ---
def set_font_kai(run, size=14, is_bold=False, color=RGBColor(0, 0, 0)):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = color
    # 強制中文字型
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 自動抓取函數 ---
def fetch_basic_info():
    info = {
        "company_name": "", 
        "cid": "",
        "area": "0", 
        "employees": "0",
        "air_area": "0"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            sheet_name = next((s for s in xl.sheet_names if "用戶基本資料" in s), None)
            
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                for r in range(len(df)):
                    for c in range(len(df.columns)):
                        val = str(df.iloc[r, c]).replace(" ", "").replace("\n", "")
                        
                        if "能源用戶名稱" in val:
                             info["company_name"] = str(df.iloc[r, c+2]).split('(')[0].strip()
                        elif "11.電號" in val:
                            info["cid"] = str(df.iloc[r, c+1]).strip()
                        elif "19.總樓地板面積" in val:
                            info["area"] = str(df.iloc[r, c+1]).strip()
                        elif "20.總空調使用面積" in val:
                            info["air_area"] = str(df.iloc[r, c+1]).strip()
                        elif "17.員工人數" in val:
                            info["employees"] = str(df.iloc[r, c+1]).strip()
                
                for k in info:
                    if info[k] == "nan" or info[k] == "None": info[k] = ""
                    
                st.success(f"✅ 已從【{sheet_name}】自動帶入資料")
        except Exception as e:
            st.error(f"自動讀取失敗：{e}")
            
    return info

# --- 3. 顯示介面 ---
st.title("📋 用戶基本資料設定")
basic_data = fetch_basic_info()

col1, col2 = st.columns(2)
with col1:
    comp_name = st.text_input("公司名稱", value=basic_data["company_name"], key="p2_comp")
    area_val = st.text_input("建物總面積 (m2)", value=basic_data["area"], key="p2_area")
    air_area_val = st.text_input("空調使用面積 (m2)", value=basic_data["air_area"], key="p2_air")
with col2:
    cid_val = st.text_input("台電電號", value=basic_data["cid"], key="p2_cid")
    emp_val = st.text_input("員工人數", value=basic_data["employees"], key="p2_emp")

# --- 4. 生成 Word 邏輯 ---
if st.button("📝 生成基本資料報告"):
    doc = Document()
    
    # 【標題】黑色 14 號 加粗
    title = doc.add_paragraph()
    set_font_kai(title.add_run("二、能源用戶概述"), size=14, is_bold=True)
    
    sub_title = doc.add_paragraph()
    set_font_kai(sub_title.add_run("2-1. 用戶簡介"), size=14, is_bold=True)
    
    # 【內文】黑色 14 號
    p = doc.add_paragraph()
    set_font_kai(p.add_run(comp_name), size=14)
    set_font_kai(p.add_run(" 總建物面積 "), size=14)
    set_font_kai(p.add_run(area_val), size=14)
    set_font_kai(p.add_run(" 平方公尺，空調使用面積 "), size=14)
    set_font_kai(p.add_run(air_area_val), size=14)
    set_font_kai(p.add_run(" 平方公尺。員工約有 "), size=14)
    set_font_kai(p.add_run(emp_val), size=14)
    set_font_kai(p.add_run(" 人。"), size=14)

    # --- 範例表格設定 (電力系統 10號 / 其他 11號) ---
    doc.add_paragraph()
    set_font_kai(doc.add_paragraph().add_run("1. 電力系統："), size=14, is_bold=True)
    
    # 電力系統表 (10 號字)
    p_table = doc.add_table(rows=1, cols=2)
    p_table.style = 'Table Grid'
    cells = p_table.rows[0].cells
    set_font_kai(cells[0].paragraphs[0].add_run("台電電號"), size=10)
    set_font_kai(cells[1].paragraphs[0].add_run(cid_val), size=10)

    doc.add_paragraph()
    set_font_kai(doc.add_paragraph().add_run("2. 照明系統："), size=14, is_bold=True)
    
    # 其他系統表 (11 號字)
    o_table = doc.add_table(rows=1, cols=2)
    o_table.style = 'Table Grid'
    o_cells = o_table.rows[0].cells
    set_font_kai(o_cells[0].paragraphs[0].add_run("燈具種類"), size=11)
    set_font_kai(o_cells[1].paragraphs[0].add_run("LED"), size=11)

    # 存入記憶體與全域倉庫
    buffer = io.BytesIO()
    doc.save(buffer)
    report_data = buffer.getvalue()
    
    if 'report_warehouse' in st.session_state:
        st.session_state['report_warehouse']["2. 用戶基本資料"] = report_data
    
    st.success("📂 報告已存入輸出中心！")
    st.download_button("💾 下載此份用戶資料", report_data, "User_Info.docx")
