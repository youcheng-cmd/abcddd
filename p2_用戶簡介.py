import streamlit as st
import pandas as pd
from docx import Document
import io

# --- 1. 自動抓取函數 ---
def fetch_basic_info():
    # 預設值，萬一沒抓到就顯示這個
    info = {
        "company_name": "家福股份有限公司", 
        "cid": "0000000000",
        "area": "0", 
        "employees": "0"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 1. 抓取「能源用戶基本資料」工作表 ---
            sheet_name = next((s for s in xl.sheet_names if "用戶基本資料" in s), None)
            if sheet_name:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                # 遍歷整個表格找關鍵字
                for r in range(len(df)):
                    for c in range(len(df.columns)):
                        val = str(df.iloc[r, c])
                        if "能源用戶名稱" in val:
                            info["company_name"] = str(df.iloc[r, c+1]).split('(')[0].strip()
                        elif "電號" in val:
                            info["cid"] = str(df.iloc[r, c+1]).strip()
                        elif "總樓地板面積" in val:
                            info["area"] = str(df.iloc[r, c+1]).strip()
                        elif "員工人數" in val:
                            info["employees"] = str(df.iloc[r, c+1]).strip()
                st.success(f"✅ 已從【{sheet_name}】自動帶入資料")
            else:
                st.warning("⚠️ 找不到名為「能源用戶基本資料」的工作表")
                
        except Exception as e:
            st.error(f"自動讀取失敗：{e}")
            
    return info

# --- 2. 顯示介面 ---
st.title("📋 用戶基本資料設定")
basic_data = fetch_basic_info()

# 讓使用者可以手動修改抓到的資料
col1, col2 = st.columns(2)
with col1:
    comp_name = st.text_input("公司名稱", value=basic_data["company_name"])
    area_val = st.text_input("建物面積", value=basic_data["area"])
with col2:
    cid_val = st.text_input("電號", value=basic_data["cid"])
    emp_val = st.text_input("員工人數", value=basic_data["employees"])

# --- 3. 產出 Word 與存入倉庫 ---
if st.button("📝 生成基本資料報告"):
    doc = Document()
    doc.add_heading(f"{comp_name} 能源使用概述", 1)
    # 這裡把文字串起來，紅字部分用 run 加顏色
    p = doc.add_paragraph()
    p.add_run(f"{comp_name} 總建物面積 ").add_run(f"{area_val}").font.color.rgb = RGBColor(255, 0, 0)
    # ... 剩下的文字邏輯 ...
    
    # 【關鍵】存入全域倉庫，讓一鍵打包能抓到
    buffer = io.BytesIO()
    doc.save(buffer)
    if 'report_warehouse' in st.session_state:
        st.session_state['report_warehouse']["2. 用戶基本資料"] = buffer.getvalue()
    st.success("📂 報告已存入輸出中心！")
