import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io

# --- 1. 自動抓取函數 ---
def fetch_basic_info():
    info = {
        "company_name": "", 
        "cid": "",
        "area": "0", 
        "employees": "0",
        "air_area": "0" # 新增空調面積
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
                        
                        # 根據截圖座標調整位移
                        if "09.能源用戶負責人" in val: # 負責人上方通常是名稱
                             info["company_name"] = str(df.iloc[r-1, c+1]).split('(')[0].strip()
                        elif "11.電號" in val:
                            info["cid"] = str(df.iloc[r, c+1]).strip()
                        elif "19.總樓地板面積" in val:
                            info["area"] = str(df.iloc[r, c+1]).strip()
                        elif "20.總空調使用面積" in val:
                            info["air_area"] = str(df.iloc[r, c+1]).strip()
                        elif "17.員工人數" in val:
                            info["employees"] = str(df.iloc[r, c+1]).strip()
                
                # 清除 nan 字眼
                for k in info:
                    if info[k] == "nan" or info[k] == "None": info[k] = ""
                    
                st.success(f"✅ 已從【{sheet_name}】自動帶入資料")
        except Exception as e:
            st.error(f"自動讀取失敗：{e}")
            
    return info

# --- 2. 顯示介面 ---
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

# --- 3. 產出 Word 與存入倉庫 ---
if st.button("📝 生成基本資料報告"):
    doc = Document()
    
    # 設定標體字型工具
    def add_red_run(paragraph, text):
        run = paragraph.add_run(text)
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True
        return run

    doc.add_heading('二、能源用戶概述', 1)
    doc.add_heading('2-1. 用戶簡介', 2)
    
    p = doc.add_paragraph()
    # 這裡示範如何拼湊 Word 文字與紅字
    add_red_run(p, comp_name)
    p.add_run(f" 總建物面積 ")
    add_red_run(p, f"{area_val}")
    p.add_run(" 平方公尺，空調使用面積 ")
    add_red_run(p, f"{air_area_val}")
    p.add_run(" 平方公尺。員工約有 ")
    add_red_run(p, f"{emp_val}")
    p.add_run(" 人。")

    # 存入記憶體與全域倉庫
    buffer = io.BytesIO()
    doc.save(buffer)
    report_data = buffer.getvalue()
    
    if 'report_warehouse' in st.session_state:
        st.session_state['report_warehouse']["2. 用戶基本資料"] = report_data
    
    st.success("📂 報告已存入輸出中心！您可以在左側打包下載。")
    st.download_button("💾 下載此份用戶資料", report_data, "User_Info.docx")
