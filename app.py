import streamlit as st

# --- 1. 網頁頁面配置 ---
st.set_page_config(page_title="節能診斷工具箱", layout="wide")

# --- 2. 側邊欄選單 ---
st.sidebar.title("🛠️ 節能診斷工具箱")
mode = st.sidebar.radio("請選擇分析項目：", ["1. 變壓器效益分析", "2. 用戶基本資料"])

# --- 3. 轉接器邏輯 ---
if mode == "1. 變壓器效益分析":
    # 呼叫第一個檔案
    exec(open("p1_變壓器分析.py", encoding="utf-8").read())

elif mode == "2. 用戶基本資料":
    # 呼叫第二個檔案
    exec(open("p2_用戶簡介.py", encoding="utf-8").read())
