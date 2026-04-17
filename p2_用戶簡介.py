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
    info = {"comp": "", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": "115年2月26日"}
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # 1. 定位【表五之二】
            p_sheet = next((s for s in xl.sheet_names if "五之二" in s), None)
            if p_sheet:
                df_p = pd.read_excel(file, sheet_name=p_sheet, header=None)
                
                for r in range(len(df_p)):
                    for c in range(len(df_p.columns)):
                        cell_val = str(df_p.iloc[r, c])
                        
                        # 找到「戶名」標籤
                        if "戶名" in cell_val:
                            # 關鍵修正：在「戶名」正下方那一列 (r+1)，往右搜尋 5 格
                            # 這樣不論它是合併在 E, F 還是 G 欄，都能抓到
                            for offset_c in range(0, 5):
                                target_c = c + offset_c
                                if target_c < len(df_p.columns):
                                    candidate = str(df_p.iloc[r + 1, target_c]).strip()
                                    if candidate != "nan" and candidate != "":
                                        # 成功抓到「凱格運動事業股份有限公司」
                                        # 同時過濾掉換行符號（因為 Excel 裡可能按了 Alt+Enter）
                                        info["comp"] = candidate.replace("\n", "").split('(')[0]
                                        break
                            break 
            
            # 2. 定位【三、能源用戶基本資料】 (抓取面積、人數、工作時數)
            # 這裡我們一個一個來檢視：
            b_sheet = next((s for s in xl.sheet_names if "能源用戶基本資料" in s), None)
            if b_sheet:
                df_b = pd.read_excel(file, sheet_name=b_sheet, header=None)
                for r in range(len(df_b)):
                    label = str(df_b.iloc[r, 1]) # B 欄標籤
                    
                    # 紅字 4：員工人數 (J15) -> Python 索引 [r, 9]
                    if "16.員工人數" in label:
                        info["emp"] = str(df_b.iloc[r, 9]).strip()
                    
                    # 紅字 5：全年工作時數 (D16) -> Python 索引 [r, 3]
                    if "17.全年工作時數" in label:
                        info["hours"] = str(df_b.iloc[r, 3]).replace(".0", "").strip()
                        
                    # 面積部分 (J16 與 D17)
                    if "18.總樓地板面積" in label:
                        info["area"] = str(df_b.iloc[r, 9]).strip()
                    if "19.總空調使用面積" in label:
                        info["air_area"] = str(df_b.iloc[r, 3]).strip()

        except Exception as e:
            st.error(f"抓取失敗: {e}")
            
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
