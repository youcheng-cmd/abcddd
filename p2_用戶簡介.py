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

# --- 2. 強化版數據抓取 ---
def fetch_intro_data():
    # 預設值 (凱格範例)
    d = {
        "comp": "能源用戶名稱", "area": "0", "air_area": "0", 
        "emp": "0", "hours": "0", "date": "115年2月26日"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            # 讀取「三、能源用戶基本資料」
            df = pd.read_excel(file, sheet_name="三、能源用戶基本資料", header=None)
            
            # 定位掃描
            for r in range(len(df)):
                row_val = str(df.iloc[r, 1]) # 檢查 B 欄標籤
                if "01.總公司名稱" in row_val:
                    d["comp"] = str(df.iloc[r, 3]).strip() # D5
                if "16.員工人數" in row_val:
                    d["emp"] = str(df.iloc[r, 9]).strip() # J15
                if "17.全年工作時數" in row_val:
                    d["hours"] = str(df.iloc[r, 3]).replace(".0", "").strip() # D16
                if "18.總樓地板面積" in row_val:
                    d["area"] = f"{df.iloc[r, 9]:,.0f}" # J16
                if "19.總空調使用面積" in row_val:
                    d["air_area"] = f"{df.iloc[r, 3]:,.0f}" # D17
            
            # 抓取填表日期作為診斷日期 (I3)
            date_val = str(df.iloc[2, 8]) 
            if "年" in date_val: d["date"] = date_val.replace("填表日期：", "").strip()

        except Exception as e:
            st.error(f"第一段資料抓取失敗: {e}")
    return d
# --- 3. 介面 ---
st.title("📋 能源用戶概述設定")
d = fetch_all_data()

# 讓使用者可以微調抓到的數字
st.subheader("📝 內文數據微調")
c1, c2, c3 = st.columns(3)
with c1:
    v_comp = st.text_input("公司名稱", d["comp"])
    v_area = st.text_input("總面積", d["area"])
with c2:
    v_air = st.text_input("空調面積", d["air_area"])
    v_emp = st.text_input("人數", d["emp"])
with c3:
    v_hours = st.text_input("工作時數", d["hours"])
    v_date = st.text_input("診斷日期", d["date"])

# --- 4. 生成 Word ---
doc = Document()

# --- 用戶簡介段落 ---
doc.add_heading('', 1) # 二、能源用戶概述
set_font_kai(doc.paragraphs[-1].add_run("二、能源用戶概述"), is_bold=True)

p_intro = doc.add_paragraph()
set_font_kai(p_intro.add_run("2-1. 用戶簡介"), is_bold=True)

p_desc = doc.add_paragraph()
# 凱格運動事業股份有限公司 (紅)
set_font_kai(p_desc.add_run(v_comp), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 總建物面積 "))
# 23,666 (紅)
set_font_kai(p_desc.add_run(v_area), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 平方公尺，空調使用面積 "))
# 20,353 (紅)
set_font_kai(p_desc.add_run(v_air), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 平方公尺，能源使用主要以 "))
set_font_kai(p_desc.add_run("電力"), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 為主，員工約有 "))
# 48 (紅)
set_font_kai(p_desc.add_run(v_emp), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 人，全年使用時間約 "))
# 5,780 (紅)
set_font_kai(p_desc.add_run(v_hours), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 小時，"))
# 診斷日期 (紅)
set_font_kai(p_desc.add_run(v_date), color=RGBColor(255, 0, 0))
set_font_kai(p_desc.add_run(" 經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

# --- 1. 電力系統表格 (黑色 10 號) ---
doc.add_paragraph()
set_font_kai(doc.add_paragraph().add_run("1.電力系統："), is_bold=True)
t1 = doc.add_table(rows=4, cols=6)
t1.style = 'Table Grid'
# 範例：填寫表格內容並設為 10 號字
def fill_cell(table, r, c, text, is_red=False):
    cell = table.rows[r].cells[c]
    run = cell.paragraphs[0].add_run(text)
    color = RGBColor(255, 0, 0) if is_red else RGBColor(0, 0, 0)
    set_font_kai(run, size=10, color=color)

fill_cell(t1, 0, 0, "台電電號：")
fill_cell(t1, 0, 1, d["cid"], is_red=True)
# ... 表格其他欄位依此類推 ...

# --- 下載按鈕 ---
buffer = io.BytesIO()
doc.save(buffer)
st.download_button("💾 生成並下載診斷報告", buffer.getvalue(), f"能源用戶概述_{v_comp}.docx")
