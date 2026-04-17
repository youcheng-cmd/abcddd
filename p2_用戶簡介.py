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
    info = {
        "comp": "未抓到名稱", "area": "0", "air_area": "0", "emp": "0", "hours": "0", "date": "115年1月1日",
        # --- 電力系統新欄位 ---
        "elec_id": "0",          # 台電電號
        "contract_type": "高壓 3 段式", # 契約型式
        "contract_cap": "0",     # 契約容量
        "volt": "22.8",          # 供電電壓
        "trans_cap": "0",        # 變壓器容量
        "cap_cap": "0",          # 電容器容量
        "low_volt": "380/220",   # 低壓側電壓
        "total_kwh": "0",        # 年總用電度
        "total_fee": "0",        # 年總金額
        "avg_price": "0",        # 平均單價
        "avg_pf": "0",           # 平均功因
        "peak_max": "0",         # 尖峰最高需量
        "offpeak_max": "0"       # 離峰最高需量
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            sheet_p = next((s for s in xl.sheet_names if "五之二" in s), None)
            if sheet_p:
                df_p = pd.read_excel(file, sheet_name=sheet_p, header=None)
                if len(df_p) > 5 and len(df_p.columns) > 4:
                    val = str(df_p.iloc[5, 4]).strip()
                    if val != "nan":
                        info["comp"] = val.split('(')[0]

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
                        res = get_near_value(row_list, "員工人數"); 
                        if res: info["emp"] = res
                    if "全年工作時數" in row_str:
                        res = get_near_value(row_list, "全年工作時數"); 
                        if res: info["hours"] = res
                    if "總樓地板面積" in row_str:
                        res = get_near_value(row_list, "總樓地板面積", min_val=100); 
                        if res: info["area"] = res
                    if "總空調使用面積" in row_str:
                        res = get_near_value(row_list, "總空調使用面積", min_val=100); 
                        if res: info["air_area"] = res
        except Exception as e:
            st.error(f"解析發生錯誤: {e}")
    return info

# --- 3. 介面 ---
st.title("📋 用戶簡介自動化")
data_pack = fetch_exact_data()

# 把自動帶入的資料收進摺疊盒 (expander)
with st.expander("🔍 檢視/微調自動抓取資料 (通常不需修改)"):
    ec1, ec2 = st.columns(2)
    with ec1:
        v_comp = st.text_input("用戶名稱 (紅字1)", data_pack["comp"])
        v_area = st.text_input("總面積 (紅字2)", data_pack["area"])
        v_air = st.text_input("空調面積 (紅字3)", data_pack["air_area"])
    with ec2:
        v_emp = st.text_input("員工人數 (紅字4)", data_pack["emp"])
        v_hours = st.text_input("工作時數 (紅字5)", data_pack["hours"])

# 這裡單獨放「診斷日期」，因為你說這個最常改
v_date = st.text_input("📅 診斷日期 (紅字6)", data_pack["date"])

# --- 電力系統介面 (維持原樣) ---
st.markdown("### ⚡ 電力系統資料")
e_c1, e_c2, e_c3 = st.columns(3)
# ... 下面原本 e_c1, e_c2, e_c3 的內容維持不變 ...
with e_c1:
    v_elec_id = st.text_input("台電電號", data_pack["elec_id"])
    v_contract_type = st.text_input("契約型式", data_pack["contract_type"])
    v_total_kwh = st.text_input("年總用電度", data_pack["total_kwh"])
    v_avg_pf = st.text_input("平均功因", data_pack["avg_pf"])
with e_c2:
    v_contract_cap = st.text_input("契約容量 (kW)", data_pack["contract_cap"])
    v_trans_cap = st.text_input("主變壓器容量 (kVA)", data_pack["trans_cap"])
    v_total_fee = st.text_input("年總金額", data_pack["total_fee"])
    v_peak_max = st.text_input("尖峰最高需量", data_pack["peak_max"])
with e_c3:
    v_volt = st.text_input("供電電壓 (kV)", data_pack["volt"])
    v_cap_cap = st.text_input("電容器容量 (kVAR)", data_pack["cap_cap"])
    v_avg_price = st.text_input("平均單價", data_pack["avg_price"])
    v_offpeak_max = st.text_input("離峰最高需量", data_pack["offpeak_max"])
# --- 4. 封裝 Word 生成邏輯 ---
def generate_docx():
    doc = Document()
    # 標題
    p_t1 = doc.add_paragraph(); set_font_kai(p_t1.add_run("二、能源用戶概述"), is_bold=True)
    p_t2 = doc.add_paragraph(); set_font_kai(p_t2.add_run("  2-1. 用戶簡介"), is_bold=True)

    # 內文段落
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
    
    # 診斷日期 (純手動輸入內容)
    set_font_kai(p.add_run(v_date), color=RGBColor(255, 0, 0)) 
    set_font_kai(p.add_run("經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# --- 5. 一鍵下載按鈕 ---
st.markdown("---")
st.download_button(
    label="💾 生成並下載用戶簡介 Word",
    data=generate_docx(),  # 點擊時直接調用生成邏輯
    file_name=f"能源用戶簡介_{v_comp}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
