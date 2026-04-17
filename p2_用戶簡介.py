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
def fetch_exact_data():
    d = {
        "comp": "", "area": "0", "air_area": "0", 
        "emp": "0", "hours": "0", "date": "115年2月26日", "cid": ""
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 抓取用戶名字 (從 表五之二) ---
            p_sheet = next((s for s in xl.sheet_names if "五之二" in s), None)
            if p_sheet:
                df_p = pd.read_excel(file, sheet_name=p_sheet, header=None)
                # 通常名字在第 5 列左右的 E 欄 (iloc[5, 4])，或是搜尋「戶名」
                for r in range(len(df_p)):
                    row_str = "".join(map(str, df_p.iloc[r, :]))
                    if "戶名" in row_str:
                        # 抓取戶名下方的內容，並去掉「公司」之後的贅字（如果有需要）
                        d["comp"] = str(df_p.iloc[r+1, 4]).strip() 
                        d["cid"] = str(df_p.iloc[r+1, 1]).strip() # 電號通常在 B 欄
                        break

            # --- 抓取其他數據 (從 三、能源用戶基本資料) ---
            b_sheet = next((s for s in xl.sheet_names if "能源用戶基本資料" in s), None)
            if b_sheet:
                df_b = pd.read_excel(file, sheet_name=b_sheet, header=None)
                for r in range(len(df_b)):
                    label = str(df_b.iloc[r, 1]) # B 欄標籤
                    if "16.員工人數" in label:
                        d["emp"] = str(df_b.iloc[r, 9]).strip() # J15
                    if "17.全年工作時數" in label:
                        d["hours"] = str(df_b.iloc[r, 3]).replace(".0", "").strip() # D16
                    if "18.總樓地板面積" in label:
                        d["area"] = f"{df_b.iloc[r, 9]:,.0f}" # J16
                    if "19.總空調使用面積" in label:
                        d["air_area"] = f"{df_b.iloc[r, 3]:,.0f}" # D17
            
        except Exception as e:
            st.error(f"抓取失敗，請手動校對。錯誤訊息: {e}")
            
    return d

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
