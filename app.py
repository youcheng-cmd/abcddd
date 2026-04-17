import streamlit as st
import pandas as pd

st.set_page_config(page_title="節能診斷工具箱", layout="wide")

# --- 初始化暫存記憶體 ---
if 'global_excel' not in st.session_state:
    st.session_state['global_excel'] = None

st.sidebar.title("🛠️ 節能診斷工具箱")

# --- A. 主畫面：全域上傳區 ---
st.sidebar.markdown("---")
st.sidebar.subheader("📂 全域資料庫 (全部工作表)")
uploaded_global = st.sidebar.file_uploader("上傳完整能源查核 Excel", type=["xlsx"], key="global_uploader")

if uploaded_global:
    st.session_state['global_excel'] = uploaded_global
    st.sidebar.success("✅ 全域檔案已就緒")

st.sidebar.markdown("---")
mode = st.sidebar.radio("請選擇分析項目：", ["1. 變壓器效益分析", "2. 用戶基本資料"])

# --- B. 轉接器邏輯 ---
if mode == "1. 變壓器效益分析":
    exec(open("p1_變壓器分析.py", encoding="utf-8").read())

elif mode == "2. 用戶基本資料":
    exec(open("p2_用戶簡介.py", encoding="utf-8").read())
