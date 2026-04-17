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

# --- 2. 修正版數據抓取邏輯 ---
# --- 修正版：只抓用戶名字測試 ---
def fetch_exact_data():
    info = {"comp": "", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": ""}
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 1. 定位表五之二 (抓用戶名稱) ---
            sheet_p = next((s for s in xl.sheet_names if "五之二" in s), None)
            if sheet_p:
                df_p = pd.read_excel(file, sheet_name=sheet_p, header=None)
                # 確保 E6 (5, 4) 存在
                if len(df_p) > 5 and len(df_p.columns) > 4:
                    raw_comp = str(df_p.iloc[5, 4]).strip()
                    info["comp"] = raw_comp.replace("\n", "").split('(')[0] if raw_comp != "nan" else "未抓到名稱"
            
            # --- 2. 定位表三 (抓取數值數據) ---
            # 改用模糊搜尋，找包含 "三" 或 "基本資料" 的表
            sheet_b = next((s for s in xl.sheet_names if "三" in s or "基本資料" in s), None)
            
            if sheet_b:
                df_b = pd.read_excel(file, sheet_name=sheet_b, header=None)
                num_rows = len(df_b)
                num_cols = len(df_b.columns)

                for r in range(num_rows):
                    # 讀取該列所有文字進行比對
                    row_str = "".join(map(str, df_b.iloc[r, :]))
                    
                    # 抓員工人數 (J15 或 L15)
                    if "員工人數" in row_str:
                        # 安全檢查：如果有 L 欄 (11) 就抓 L，否則抓 J (9)
                        idx = 11 if num_cols > 11 else (9 if num_cols > 9 else -1)
                        if idx != -1:
                            info["emp"] = str(df_b.iloc[r, idx]).replace(".0", "").strip()

                    # 抓工作時數 (D16)
                    if "工作時數" in row_str:
                        if num_cols > 3:
                            info["hours"] = str(df_b.iloc[r, 3]).replace(".0", "").strip()

                    # 抓總面積 (L16 或 J16)
                    if "總樓地板面積" in row_str:
                        idx = 11 if num_cols > 11 else (9 if num_cols > 9 else -1)
                        if idx != -1:
                            info["area"] = str(df_b.iloc[r, idx]).strip()

                    # 抓空調面積 (D17)
                    if "空調使用面積" in row_str:
                        if num_cols > 3:
                            info["air_area"] = str(df_b.iloc[r, 3]).strip()
                
                # 抓診斷日期 (通常在第 3 列附近)
                if num_rows > 2 and num_cols > 8:
                    info["date"] = str(df_b.iloc[2, 8]).replace("填表日期：", "").strip()
            
            else:
                st.error("Excel 中找不到包含 '三' 或 '基本資料' 的工作表")

        except Exception as e:
            st.error(f"解析發生錯誤: {e}")
            
    # 最後清理資料，如果是 nan 就歸 0
    for k in info:
        if info[k] == "nan" or not info[k]: 
            info[k] = "0" if k != "comp" else "未抓到名稱"
            
    return info

# --- 3. 介面 ---
st.title("📋 用戶簡介自動化")
d = fetch_exact_data()

c1, c2 = st.columns(2)
with c1:
    v_comp = st.text_input("用戶名稱 (紅字1)", d["comp"])
    v_area = st.text_input("總面積 (紅字2)", d["area"])
    v_air = st.text_input("空調面積 (紅字3)", d["air_area"])
with c2:
    v_emp = st.text_input("員工人數 (紅字4)", d["emp"])
    v_hours = st.text_input("工作時數 (紅字5)", d["hours"])
    v_date = st.text_input("診斷日期 (紅字6)", d["date"])

# --- 4. 生成 Word 並下載 ---
doc = Document()

# 設定標題 (黑色 14號 加粗)
p_t1 = doc.add_paragraph(); set_font_kai(p_t1.add_run("二、能源用戶概述"), is_bold=True)
p_t2 = doc.add_paragraph(); set_font_kai(p_t2.add_run("  2-1. 用戶簡介"), is_bold=True)

# 第一段內文 (標楷體 14號)
p = doc.add_paragraph()
p.paragraph_format.first_line_indent = Pt(28) # 首行縮排兩格

# 拼湊紅黑文字
set_font_kai(p.add_run(v_comp), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("總建物面積"))
set_font_kai(p.add_run(v_area), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("平方公尺，空調使用面積"))
set_font_kai(p.add_run(v_air), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("平方公尺，能源使用主要以"))
set_font_kai(p.add_run("電力"), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("為主，員工約有"))
set_font_kai(p.add_run(v_emp), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("人，全年使用時間約"))
set_font_kai(p.add_run(v_hours), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("小時，"))
set_font_kai(p.add_run(v_date), color=RGBColor(255, 0, 0)) # 紅
set_font_kai(p.add_run("經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

# --- 生成下載按鈕 ---
buffer = io.BytesIO()
doc.save(buffer)
st.download_button(
    label="💾 生成並下載用戶簡介 Word",
    data=buffer.getvalue(),
    file_name=f"能源用戶簡介_{v_comp}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
