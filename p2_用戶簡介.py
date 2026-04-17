import streamlit as st
import pandas as pd
from docx import Document
import io

# --- 1. 自動抓取函數 ---
def fetch_basic_info():
    info = {
        "company_name": "誠友開發股份有限公司", # 預設值
        "area": "0", "employees": "0", "cid": "0000000000",
        "contract_cap": "0", "total_kwh": "0", "total_fee": "0"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            # 讀取基本資料表
            df_basic = pd.read_excel(file, sheet_name="三、能源用戶基本資料", header=None)
            # 這裡用座標或搜尋文字抓取
            info["company_name"] = df_basic.iloc[4, 3] # D5
            info["area"] = df_basic.iloc[15, 11]       # L16
            
            # 讀取電能表
            df_power = pd.read_excel(file, sheet_name=0, header=None) # 假設第一個是電能表
            # ... 依此類推抓取電號、電費 ...
            st.success("✅ 已自動從 Excel 帶入基本資料")
        except:
            st.warning("⚠️ 自動抓取部分失敗，請手動校對")
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
