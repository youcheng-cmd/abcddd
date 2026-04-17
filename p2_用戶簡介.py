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
    # 這裡直接給預設值
    info = {"comp": "", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": "115年1月1日"}
    
    # ... 中間抓取公司、面積、人數的代碼維持不變 ...

    # 【刪除】原本那段 if "填表日期" in row_str: ... 的代碼，不要從 Excel 撈日期了
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # 1. 抓名稱 (包含 "五之二" 的表)
            sheet_p = next((s for s in xl.sheet_names if "五之二" in s), None)
            if sheet_p:
                df_p = pd.read_excel(file, sheet_name=sheet_p, header=None)
                if len(df_p) > 5 and len(df_p.columns) > 4:
                    raw_comp = str(df_p.iloc[5, 4]).strip()
                    info["comp"] = raw_comp.split('(')[0] if raw_comp != "nan" else "未抓到名稱"

            # 2. 抓數值 (包含 "三" 或 "基本資料" 的表)
            sheet_b = next((s for s in xl.sheet_names if "三" in s or "基本資料" in s), None)
            if sheet_b:
                df_b = pd.read_excel(file, sheet_name=sheet_b, header=None)
                
                # --- 定義內部搜尋函數 (確保縮排在此處正確) ---
                def get_near_value(items, keyword, min_val=0):
                    import re
                    for i, item in enumerate(items):
                        if keyword in str(item):
                            # 找到關鍵字後，往後搜尋 4 格
                            for target in items[i+1 : i+5]:
                                if target is None or str(target).lower() == "nan":
                                    continue
                                clean = str(target).replace(",", "").replace(" ", "")
                                matches = re.findall(r"[-+]?\d*\.\d+|\d+", clean)
                                if matches:
                                    try:
                                        num = int(round(float(matches[0])))
                                        if num > min_val:
                                            return f"{num:,d}"
                                    except:
                                        continue
                    return None

                # 遍歷每一列進行匹配
                for r in range(len(df_b)):
                    row_list = list(df_b.iloc[r, :])
                    row_str = "".join([str(i) for i in row_list])
                    
                    if "員工人數" in row_str:
                        res = get_near_value(row_list, "員工人數")
                        if res: info["emp"] = res

                    if "全年工作時數" in row_str:
                        # 只抓 5780，避開右邊的面積
                        res = get_near_value(row_list, "全年工作時數")
                        if res: info["hours"] = res

                    if "總樓地板面積" in row_str:
                        res = get_near_value(row_list, "總樓地板面積", min_val=100)
                        if res: info["area"] = res

                    if "總空調使用面積" in row_str:
                        res = get_near_value(row_list, "總空調使用面積", min_val=100)
                        if res: info["air_area"] = res
                    
                    if "填表日期" in row_str:
                        for item in reversed(row_list):
                            if item and "年" in str(item):
                                # 這裡多加一個 .replace("填表日期：", "") 確保變數乾淨
                                info["date"] = str(item).replace("填表日期：", "").replace(" ", "").strip()
                                break

        except Exception as e:
            st.error(f"解析發生錯誤: {e}")
            
    # 清理數據
   # 最後清理資料
    for k in info:
        if "nan" in str(info[k]).lower() or not str(info[k]).strip():
            # 針對 date 設定特定預設值
            if k == "date":
                info[k] = "115年1月1日"
            else:
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
    # 這裡加一個 .replace 確保從 Excel 抓過來的日期不會帶有「填表日期：」
# 修改第 116 行
# 不再做任何 replace，直接讓使用者填寫
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
# 使用 .replace 強制刪除變數中可能帶有的「填表日期：」字樣
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
