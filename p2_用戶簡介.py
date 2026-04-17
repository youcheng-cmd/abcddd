import streamlit as st
import pandas as pd

st.title("📄 用戶基本資料抓取")

# --- 區域上傳區 ---
st.info("💡 提示：若已在主畫面左側上傳完整檔案，此處會自動讀取。")
uploaded_local = st.file_uploader("或：上傳單張工作表 (例如表五之二)", type=["xlsx"], key="local_p2")

target_file = None

# 判斷優先權：1. 區域上傳 2. 全域上傳
if uploaded_local:
    target_file = uploaded_local
    st.success("正在使用：單張上傳版本")
elif st.session_state['global_excel']:
    target_file = st.session_state['global_excel']
    st.success("正在使用：全域資料庫版本")

# --- 開始撈料邏輯 ---
if target_file:
    try:
        # 獲取所有工作表名稱
        xl = pd.ExcelFile(target_file)
        sheet_names = xl.sheet_names
        
        # 範例：抓取「三、能源用戶基本資料」
        if "三、能源用戶基本資料" in sheet_names:
            df_basic = pd.read_excel(target_file, sheet_name="三、能源用戶基本資料")
            # 這裡寫你的抓取座標邏輯，例如：
            # company_name = df_basic.iloc[2, 1] 
            st.write("已偵測到基本資料工作表！")
            
        # 範例：抓取「表五之二」
        if any("表五之二" in s for s in sheet_names):
            target_sheet = [s for s in sheet_names if "表五之二" in s][0]
            df_elec = pd.read_excel(target_file, sheet_name=target_sheet)
            st.write(f"已偵測到電能統計表：{target_sheet}")
            # 這裡可以自動撈出電號...
            
    except Exception as e:
        st.error(f"撈取資料時發生錯誤：{e}")
else:
    st.warning("請在左側或此處上傳 Excel 檔案以開始。")
