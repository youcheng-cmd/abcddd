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
    info = {
        "comp": "", "area": "0", "air_area": "0", 
        "emp": "0", "hours": "0", "date": "115年2月26日"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 1. 抓取用戶名稱 (掃描所有工作表) ---
            # 有時候名字出現在表五之二，有時候在基本資料，我們全找一遍
            for s_name in xl.sheet_names:
                df_tmp = pd.read_excel(file, sheet_name=s_name, header=None)
                for r in range(len(df_tmp)):
                    for c in range(len(df_tmp.columns)):
                        cell_str = str(df_tmp.iloc[r, c])
                        # 如果看到「戶名」或「能源用戶名稱」
                        if "戶名" in cell_str or "用戶名稱" in cell_str:
                            # 檢查下方與右方共 6 格，抓第一個有字且不是 nan 的
                            search_coords = [(r+1, c), (r+1, c+1), (r, c+1), (r, c+2)]
                            for sr, sc in search_coords:
                                if sr < len(df_tmp) and sc < len(df_tmp.columns):
                                    candidate = str(df_tmp.iloc[sr, sc]).strip()
                                    if candidate != "nan" and len(candidate) > 4: # 名字通常大於4格字
                                        info["comp"] = candidate.split('(')[0].split('（')[0]
                                        break
                    if info["comp"]: break
                if info["comp"]: break

            # --- 2. 抓取面積、人數、工時 (針對基本資料表) ---
            b_sheet = next((s for s in xl.sheet_names if "基本資料" in s), None)
            if b_sheet:
                df_b = pd.read_excel(file, sheet_name=b_sheet, header=None)
                for r in range(len(df_b)):
                    for c in range(len(df_b.columns)):
                        val = str(df_b.iloc[r, c])
                        
                        # 員工人數 (找關鍵字 "16." 或 "員工人數")
                        if "員工人數" in val:
                            info["emp"] = str(df_b.iloc[r, c+1]).strip()
                        
                        # 總面積 (找關鍵字 "18." 或 "總樓地板面積")
                        if "總樓地板面積" in val:
                            info["area"] = str(df_b.iloc[r, c+1]).strip()

                        # 全年工作時數 (找關鍵字 "17." 或 "工作時數")
                        if "工作時數" in val:
                            # 數字可能在右方 c+1
                            res = str(df_b.iloc[r, c+1]).replace("小時", "").strip()
                            info["hours"] = res if res != "nan" else "0"

        except Exception as e:
            st.error(f"解析失敗: {e}")
            
    # 清理所有 nan 或空白
    for k in info:
        if info[k] == "nan" or not info[k]: info[k] = "0" if k != "comp" else ""
            
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
