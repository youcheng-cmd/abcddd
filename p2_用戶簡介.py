import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 通用工具函數 ---
def set_font_kai(run, size=14, is_bold=False, color=RGBColor(0, 0, 0)):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = color
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

# --- 2. 強化版抓取函數 (解決 nan 問題) ---
def fetch_basic_info():
    info = {"company_name": "", "cid": "", "area": "0", "employees": "0", "air_area": "0"}
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            sheet_name = next((s for s in xl.sheet_names if "能源用戶基本資料" in s), None)
            
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                for r in range(len(df)):
                    for c in range(len(df.columns)):
                        val = str(df.iloc[r, c]).replace(" ", "").replace("\n", "")
                        
                        # 依照截圖 C 欄關鍵字抓取 D 欄或 G 欄的值
                        if "01.總公司名稱" in val:
                            info["company_name"] = str(df.iloc[r, c+2]).split('(')[0].strip()
                        elif "11.電號" in val:
                            info["cid"] = str(df.iloc[r, c+1]).strip()
                        elif "19.總樓地板面積" in val:
                            info["area"] = str(df.iloc[r, c+1]).strip()
                        elif "20.總空調使用面積" in val:
                            info["air_area"] = str(df.iloc[r, c+1]).strip()
                        elif "17.員工人數" in val:
                            info["employees"] = str(df.iloc[r, c+1]).strip()
                
                # 清除 nan
                for k in info:
                    if info[k] in ["nan", "None", "0.0"]: info[k] = ""
            st.success(f"✅ 已從【{sheet_name}】帶入資料")
        except Exception as e:
            st.error(f"自動讀取失敗：{e}")
    return info

# --- 3. 介面與資料預填 ---
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

# --- 4. 準備產出 Word ---
doc = Document()

# [內容產出]
title = doc.add_paragraph()
set_font_kai(title.add_run("二、能源用戶概述"), size=14, is_bold=True)
sub_title = doc.add_paragraph()
set_font_kai(sub_title.add_run("2-1. 用戶簡介"), size=14, is_bold=True)

p = doc.add_paragraph()
set_font_kai(p.add_run(f"{comp_name} 總建物面積為 "), size=14)
set_font_kai(p.add_run(area_val), size=14)
set_font_kai(p.add_run(" 平方公尺，空調使用面積 "), size=14)
set_font_kai(p.add_run(air_area_val), size=14)
set_font_kai(p.add_run(" 平方公尺。員工約有 "), size=14)
set_font_kai(p.add_run(emp_val), size=14)
set_font_kai(p.add_run(" 人。"), size=14)

# [電力系統表] (10號字)
doc.add_paragraph()
set_font_kai(doc.add_paragraph().add_run("1. 電力系統："), size=14, is_bold=True)
p_table = doc.add_table(rows=1, cols=2)
p_table.style = 'Table Grid'
cells = p_table.rows[0].cells
set_font_kai(cells[0].paragraphs[0].add_run("台電電號"), size=10)
set_font_kai(cells[1].paragraphs[0].add_run(cid_val), size=10)

# [存入 Buffer]
buffer = io.BytesIO()
doc.save(buffer)
report_data = buffer.getvalue()

# 更新全域倉庫供一鍵打包
if 'report_warehouse' in st.session_state:
    st.session_state['report_warehouse']["2. 用戶基本資料"] = report_data

# --- 5. 輸出按鈕 ---
st.download_button(
    label="💾 生成報告並直接下載",
    data=report_data,
    file_name=f"能源用戶簡介_{comp_name}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
