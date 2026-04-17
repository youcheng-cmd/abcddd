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
def fetch_all_data():
    # 初始化所有紅字變數的預設值
    data = {
        "comp": "凱格運動事業", "area": "0", "air_area": "0", "emp": "0", "hours": "0",
        "date": "115年1月21日", "cid": "000000000", "type": "高壓3段式", "contract": "0",
        "kv": "22.8", "tr_cap": "0", "cap_cap": "0", "v_type": "380/220",
        "kwh": "0", "money": "0", "avg_price": "0", "pf": "0", "peak": "0", "off_peak": "0"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            # 讀取基本資料表
            df_basic = pd.read_excel(file, sheet_name="三、能源用戶基本資料", header=None)
            # 讀取電費統計表 (假設名稱包含五之二)
            xl = pd.ExcelFile(file)
            p_sheet = next((s for s in xl.sheet_names if "五之二" in s), xl.sheet_names[0])
            df_power = pd.read_excel(file, sheet_name=p_sheet, header=None)

            # 定位抓取 (這裡示範關鍵字抓取)
            for r in range(len(df_basic)):
                row_str = "".join(map(str, df_basic.iloc[r, :]))
                if "能源用戶名稱" in row_str: data["comp"] = str(df_basic.iloc[r, 3]).split('(')[0]
                if "總樓地板面積" in row_str: data["area"] = f"{df_basic.iloc[r, 11]:,.0f}"
                if "總空調使用面積" in row_str: data["air_area"] = f"{df_basic.iloc[r, 3]:,.0f}"
                if "員工人數" in row_str: data["emp"] = str(df_basic.iloc[r, 11])
                if "全年工作時數" in row_str: data["hours"] = f"{df_basic.iloc[r, 3]:,.0f}"
            
            # 電力系統數據 (從表五之二抓取最後一列的合計/平均)
            data["cid"] = str(df_power.iloc[5, 2]) # 範例座標
            data["kwh"] = f"{df_power.iloc[21, 11]:,.0f}"
            data["money"] = f"{df_power.iloc[21, 14]:,.0f}"
            data["avg_price"] = f"{df_power.iloc[22, 14]:.2f}"
            
        except: pass
    return data

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
