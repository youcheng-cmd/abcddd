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
            # 直接讀取工作表
            df_b = pd.read_excel(file, sheet_name="三、能源用戶基本資料", header=None)
            df_p = pd.read_excel(file, sheet_name="表五之二、電能使用量統計表", header=None)

            # --- 1. 用戶名稱 (鎖定：表五之二 第 6 列，第 5 欄起往右找第一個有字的) ---
            # 在 Excel 裡是第 5 列 (索引 4) 或第 6 列 (索引 5)
            row_candidates = [4, 5] 
            for r in row_candidates:
                # E 到 H 欄 (索引 4 到 7) 掃描合併儲存格
                for c in range(4, 8):
                    val = str(df_p.iloc[r, c]).strip()
                    if val != "nan" and "戶名" not in val and len(val) > 2:
                        info["comp"] = val.replace("\n", "").split('(')[0]
                        break
                if info["comp"]: break

            # --- 2. 員工人數 (鎖定：基本資料 B15 是標籤，L15 是資料) ---
            # J 欄或 L 欄常變動，我們改用關鍵字搜尋整列
            for r in range(13, 17): # 在 14-16 列附近找
                row_str = "".join(map(str, df_b.iloc[r, :]))
                if "員工人數" in row_str:
                    # 資料通常在 J (9) 或 L (11) 欄
                    info["emp"] = str(df_b.iloc[r, 11]).replace(".0", "").strip()

            # --- 3. 工作時數 (鎖定：基本資料 D16) ---
            info["hours"] = str(df_b.iloc[15, 3]).replace(".0", "").strip()

            # --- 4. 總面積 (鎖定：基本資料 L16) ---
            info["area"] = str(df_b.iloc[15, 11]).strip()

            # --- 5. 空調面積 (鎖定：基本資料 D17) ---
            info["air_area"] = str(df_b.iloc[16, 3]).strip()

            # --- 6. 診斷日期 (鎖定：基本資料 I3) ---
            info["date"] = str(df_b.iloc[2, 8]).replace("填表日期：", "").strip()

        except Exception as e:
            st.error(f"座標解析錯誤: {e}")
            
    # 最後檢查：如果抓到 nan，強制歸零
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
