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
        "elec_id": "0", "contract_type": "高壓 3 段式", "contract_cap": "0", "volt": "22.8",
        "trans_cap": "0", "cap_cap": "0", "low_volt": "380/220",
        "total_kwh": "0", "total_fee": "0", "avg_price": "0", "avg_pf": "0", "peak_max": "0", "offpeak_max": "0"
    }
    
    if 'global_excel' in st.session_state and st.session_state['global_excel'] is not None:
        try:
            file = st.session_state['global_excel']
            xl = pd.ExcelFile(file)
            
            # --- 處理「表五之二」 (抓電號、度數、需量、功因) ---
            sheet_p = next((s for s in xl.sheet_names if "五之二" in s), None)
            if sheet_p:
                df_p = pd.read_excel(file, sheet_name=sheet_p, header=None)
                
                # 1. 抓電號 (固定在 A3 格附近)
                import re
                p3_val = str(df_p.iloc[2, 0])
                id_match = re.search(r'\d{11}', p3_val.replace("-", ""))
                if id_match: info["elec_id"] = id_match.group()

                # 2. 遍歷每一列，尋找「合計」與「平均」
                for r_idx in range(len(df_p)):
                    row_list = df_p.iloc[r_idx, :].tolist()
                    row_str = "".join([str(x) for x in row_list])
                    
                    if "合計" in row_str:
                        # 度數合計：在 M 欄 (索引 12)
                        val_kwh = row_list[12]
                        if pd.notnull(val_kwh):
                            info["total_kwh"] = f"{int(float(val_kwh)):,d}"
                        # 總電費：在 P 欄 (索引 15)
                        val_fee = row_list[15]
                        if pd.notnull(val_fee):
                            info["total_fee"] = f"{int(float(val_fee)):,d}"

                    if "平均" in row_str:
                        # 契約容量：在 D 欄 (索引 3)
                        if pd.notnull(row_list[3]):
                            info["contract_cap"] = str(int(float(row_list[3])))
                        # 平均功因：在 O 欄 (索引 14)
                        if pd.notnull(row_list[14]):
                            info["avg_pf"] = str(int(float(row_list[14])))
                        # 平均單價：在 P 欄 (索引 15)
                        if pd.notnull(row_list[15]):
                            info["avg_price"] = str(round(float(row_list[15]), 2))

                # 3. 抓最高需量 (掃描 E~H 欄的所有月份數字，找最大值)
                # 需量範圍約在第 10 列到 21 列 (索引 9~20)，欄位 E, F, G, H (索引 4, 5, 6, 7)
                demand_list = []
                for r_demand in range(9, 21):
                    for c_demand in [4, 5, 6, 7]:
                        v = df_p.iloc[r_demand, c_demand]
                        try:
                            if pd.notnull(v) and str(v).strip() != "-":
                                demand_list.append(float(v))
                        except: continue
                if demand_list:
                    info["peak_max"] = str(int(max(demand_list)))
            # --- 處理「表八」(變壓器、電容器) ---
            sheet_8 = next((s for s in xl.sheet_names if "八" in s), None)
            if sheet_8:
                df_8 = pd.read_excel(file, sheet_name=sheet_8, header=None)
                try:
                    # 變壓器容量加總 F8, G8, H8 (索引 7, 欄 5, 6, 7)
                    t_sum = 0
                    for c in [5, 6, 7]:
                        v = df_8.iloc[7, c]
                        if pd.notnull(v): t_sum += float(v)
                    info["trans_cap"] = f"{int(t_sum):,d}"
                    # 高壓電容器 O26 (索引 25, 欄 14)
                    info["cap_cap"] = str(int(float(df_8.iloc[25, 14])))
                except:
                    pass

            # --- 處理「基本資料」(人數、面積、工時) ---
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
st.title("📋 節能診斷自動化工具")
data_pack = fetch_exact_data()

with st.expander("🔍 檢視/微調自動抓取資料"):
    ec1, ec2 = st.columns(2)
    with ec1:
        v_comp = st.text_input("用戶名稱", data_pack["comp"])
        v_area = st.text_input("總面積", data_pack["area"])
        v_air = st.text_input("空調面積", data_pack["air_area"])
    with ec2:
        v_emp = st.text_input("員工人數", data_pack["emp"])
        v_hours = st.text_input("工作時數", data_pack["hours"])

v_date = st.text_input("📅 診斷日期", data_pack["date"])

st.markdown("### ⚡ 電力系統資料")
e_c1, e_c2, e_c3 = st.columns(3)
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
    set_font_kai(p.add_run(v_date), color=RGBColor(255, 0, 0)) 
    set_font_kai(p.add_run("經由實地查訪貴單位之公用系統使用情形及輔導診斷概述如下："))

    doc.add_paragraph()
    set_font_kai(doc.add_paragraph().add_run("1.電力系統："), is_bold=True)

    table = doc.add_table(rows=5, cols=3)
    table.style = 'Table Grid'
    
    cell_id = table.cell(0, 0); cell_id.merge(table.cell(0, 2))
    p_id = cell_id.paragraphs[0]
    set_font_kai(p_id.add_run("台電電號："), size=12)
    set_font_kai(p_id.add_run(v_elec_id), size=12, color=RGBColor(255, 0, 0))

    r1 = table.rows[1].cells
    set_font_kai(r1[0].paragraphs[0].add_run(f"契約型式：{v_contract_type}"), size=12)
    set_font_kai(r1[1].paragraphs[0].add_run(f"契約容量：{v_contract_cap} [kW]"), size=12)
    set_font_kai(r1[2].paragraphs[0].add_run(f"台電供電電壓：{v_volt} [kV]"), size=12)

    r2 = table.rows[2].cells
    set_font_kai(r2[0].paragraphs[0].add_run(f"主變壓器總裝置容量：{v_trans_cap} [kVA]"), size=12)
    set_font_kai(r2[1].paragraphs[0].add_run(f"電容器裝置容量：{v_cap_cap} [kVAR]"), size=12)
    set_font_kai(r2[2].paragraphs[0].add_run(f"低壓側電壓：380/220 [V]"), size=12)

    r3 = table.rows[3].cells
    set_font_kai(r3[0].paragraphs[0].add_run(f"年總用電度：{v_total_kwh} [kWh]"), size=12)
    set_font_kai(r3[1].paragraphs[0].add_run(f"年總金額：{v_total_fee} [元]"), size=12)
    set_font_kai(r3[2].paragraphs[0].add_run(f"平均單價：{v_avg_price} [元/kWh]"), size=12)

    r4 = table.rows[4].cells
    set_font_kai(r4[0].paragraphs[0].add_run(f"平均功因：{v_avg_pf} [%]"), size=12)
    set_font_kai(r4[1].paragraphs[0].add_run(f"尖峰最高需量：{v_peak_max} [kW]"), size=12)
    set_font_kai(r4[2].paragraphs[0].add_run(f"離峰最高需量：{v_offpeak_max} [kW]"), size=12)

    for row in table.rows:
        for cell in row.cells: cell.vertical_alignment = 1

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# --- 5. 下載按鈕 ---
st.markdown("---")
st.download_button(
    label="💾 生成並下載用戶簡介 Word",
    data=generate_docx(),
    file_name=f"能源用戶簡介_{v_comp}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
