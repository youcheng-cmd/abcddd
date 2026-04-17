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

# --- 2. 數據抓取邏輯 ---
def fetch_exact_data():
    # 初始化資料，日期預設為 115年1月1日
    info = {"comp": "未抓到名稱", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": "115年1月1日"}
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # 1. 抓名稱 (包含 "五之二" 的表)
            sheet_p = next((s for s in xl.sheet_names if "五之二" in s), None)
            if sheet_p:
                df_p = pd.read_excel(file, sheet_name=sheet_p, header=None)
                if len(df_p) > 5 and len(df_p.columns) > 4:
                    val = str(df_p.iloc[5, 4]).strip()
                    if val != "nan":
                        info["comp"] = val.split('(')[0]

            # 2. 抓數值 (包含 "三" 或 "基本資料" 的表)
            sheet_b = next((s for s in xl.sheet_names if "三" in s or "基本資料" in s), None)
            if sheet_b:
                df_b = pd.read_excel(file, sheet_name=sheet_b, header=None)
                
                def get_near_value(items, keyword, min_val=0):
                    import re
                    for i, item in enumerate(items):
                        if keyword in str(item):
                            for target in items[i+1 : i+5]:
                                if target is None or str(target).lower() == "nan": continue
                                clean = str(target).replace(",", "").replace(" ", "")
                                matches = re.findall(r"[-+]?\d*\.\d+|\d+", clean)
                                if matches:
                                    try:
                                        num = int(round(float(matches[0])))
                                        if num > min_val: return f"{num:,d}"
                                    except: continue
                    return None

                for r in range(len(df_b)):
                    row_list = list(df_b.iloc[r, :])
                    row_str = "".join([str(i) for i in row_list])
                    
                    if "員工人數" in row_str:
                        res = get_near_value(row_list, "員工人數")
                        if res: info["emp"] = res
                    if "全年工作時數" in row_str:
                        res = get_near_value(row_list, "全年工作時數")
                        if res: info["hours"] = res
                    if "總樓地板面積" in row_str:
                        res = get_near_value(row_list, "總樓地板面積", min_val=100)
                        if res: info["area"] = res
                    if "總空調使用面積" in row_str:
                        res = get_near_value(row_list, "總空調使用面積", min_val=100)
                        if res: info["air_area"] = res

        except Exception as e:
            st.error(f"解析發生錯誤: {e}")
            
    return info

# --- 3. 介面 ---
st.title("📋 用戶簡介自動化")
data_pack = fetch_exact_data()

c1, c2 = st.columns(2)
with c1:
    v_comp = st.text_input("用戶名稱 (紅字1)", data_pack["comp"])
    v_area = st.text_input("總面積 (紅字2)", data_pack["area"])
    v_air = st.text_input("空調面積 (紅字3)", data_pack["air_area"])
with c2:
    v_emp = st.text_input("員工人數 (紅字4)", data_pack["emp"])
    v_hours = st.text_input("工作時數 (紅字5)", data_pack["hours"])
    # 日期直接顯示，不帶標籤
    v_date = st.text_input("診斷日期 (紅字6)", data_pack["date"])

# --- 4. 生成 Word 並下載 ---
if st.button("📝 生成預覽並準備下載"):
    doc = Document()
    p_t1 = doc.add_paragraph(); set_font_kai(p_t1.add_run("二、能源用戶概述"), is_bold=True)
    p_t2 = doc.add_paragraph(); set_font_kai(p_t2.add_run("  2-1. 用戶簡介"), is_bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(28)
    
    set_font_kai(p.add_run(v_comp), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("總建物面積"))
    set_font_kai(p.add_run(v_area), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("平方公尺，空調使用面積"))
    set_font_kai(p.add_run(v_air), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("平方公尺，能源使用主要以"))
    set_font_kai(p.add_run("電力"), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("為主，員工約有"))
    set_font_kai(p.add_run(v_emp), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("人，全年使用時間約"))
    set_font_kai(p.add_run(v_hours), color=RGBColor(255, 0, 0))
    set_font_kai(p.add_run("小時，"))
    
    # 這裡只放 v_date，絕對乾淨
    set_font_kai(p.add_run(v_date), color=RGBColor(255, 0, 0)) 
    set_font_kai(p.add_run("經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

    buffer = io.BytesIO()
    doc.save(buffer)
    st.download_button(
        label="💾 下載用戶簡介 Word",
        data=buffer.getvalue(),
        file_name=f"能源用戶簡介_{v_comp}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
