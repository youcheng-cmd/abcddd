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
        "comp": "", # 沒抓到就留白，不要預設值
        "area": "0", 
        "air_area": "0", 
        "emp": "0", 
        "hours": "0", 
        "date": "115年2月26日"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 鎖定【表五之二】 ---
            p_sheet = next((s for s in xl.sheet_names if "五之二" in s), None)
            if p_sheet:
                # 讀取整張表，不設標題
                df_p = pd.read_excel(file, sheet_name=p_sheet, header=None)
                
                # 開始地毯式搜尋「戶名」這兩個字
                for r in range(len(df_p)):
                    for c in range(len(df_p.columns)):
                        cell_val = str(df_p.iloc[r, c]).replace(" ", "").replace("\n", "")
                        
                        if "戶名" in cell_val:
                            # 座標對接：根據截圖，名字通常在「戶名」標籤的【正下方】(r+1)
                            # 但因為有合併儲存格，我們要往右掃描 (c 到 c+5) 找到第一個有字的人
                            for target_c in range(c, c + 6):
                                candidate = str(df_p.iloc[r + 1, target_c]).strip()
                                if candidate != "nan" and candidate != "":
                                    # 抓到了！去掉括號贅字
                                    info["comp"] = candidate.split('(')[0].split('（')[0]
                                    break
                            break # 抓到名字就跳出迴圈
            
            # --- 鎖定【三、能源用戶基本資料】 (抓取其他數值) ---
            b_sheet = next((s for s in xl.sheet_names if "能源用戶基本資料" in s), None)
            if b_sheet:
                df_b = pd.read_excel(file, sheet_name=b_sheet, header=None)
                for r in range(len(df_b)):
                    label = str(df_b.iloc[r, 1]) # 檢查 B 欄
                    if "18.總樓地板面積" in label:
                        info["area"] = str(df_b.iloc[r, 9]).strip() # J16
                    if "19.總空調使用面積" in label:
                        info["air_area"] = str(df_b.iloc[r, 3]).strip() # D17
                    # 這裡先只修名字，如果你發現員工人數 (紅字4) 還是 0，等等我們再修它
                        
        except Exception as e:
            st.error(f"解析過程中發生錯誤: {e}")
            
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
