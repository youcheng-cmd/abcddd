import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (數據精確抓取版)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", min_value=0, max_value=100, value=95)
pf_after = pf_after_input / 100

age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"]
)

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_pure_number(text):
    """強力數字提取器：處理逗號、單位、括號"""
    if pd.isna(text): return 0.0
    # 先過濾掉逗號，避免 1,500 被當成字串
    clean_text = str(text).replace(',', '').strip()
    # 使用正則表達式抓取數字（含小數點）
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", clean_text)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                if str(raw_df.iloc[r, c]).replace(' ', '') == "序號":
                    anchors.append((r, c))
        
        for r_start, c_start in anchors:
            for offset in range(1, 10):
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                if pd.isna(raw_df.iloc[r_start, target_col]) or str(raw_df.iloc[r_start, target_col]).strip() == "": continue
                
                # 初始化設備數據
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                     "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0, "現況功因": 0.8}
                specs = []
                
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    # 左右鄰居標籤探測 (處理 TR-7 等標籤偏左問題)
                    label_raw = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    if (label_raw == "nan" or label_raw == "") and c_start > 0:
                        label_raw = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                    
                    label_clean = label_raw.replace(' ', '')
                    if r_offset > 0 and "序號" in label_clean: break
                    if any(k in label_clean for k in filter_keywords): continue
                    
                    val_cell = raw_df.iloc[curr_r, target_col]
                    val_str = str(val_cell).strip()
                    if label_raw == "nan" or not label_raw: continue

                    # --- 數據提取邏輯 (使用強力提取器) ---
                    if "位置" in label_raw or "建築" in label_raw: d["建築物"] = val_str
                    if "編號" in label_raw: d["編號"] = val_str
                    if "廠牌" in label_raw: d["廠牌"] = val_str
                    if "型式" in label_raw: d["型式"] = val_str
                    
                    if any(k in label_raw for k in ["年份", "出廠"]):
                        y_num = extract_pure_number(val_str)
                        d["年份"] = int(y_num) + 1911 if 0 < y_num < 200 else int(y_num)
                        
                    if "容量" in label_raw:
                        d["容量"] = int(extract_pure_number(val_str))
                        
                    if any(k in label_raw for k in ["利用率", "負載率"]):
                        u_num = extract_pure_number(val_str)
                        # 智慧判斷：如果 Excel 給 0.32 轉為 32%
                        d["負載率"] = u_num * 100 if 0 < u_num < 1 else u_num
                        
                    if any(k in label_raw for k in ["功因", "PF"]):
                        pf_num = extract_pure_number(val_str)
                        d["現況功因"] = pf_num / 100 if pf_num > 1 else pf_num
                        
                    if any(k in label_raw for k in ["鐵損", "無載損"]):
                        d["鐵損"] = extract_pure_number(val_str)
                        
                    if any(k in label_raw for k in ["銅損", "負載損", "全載損"]):
                        d["滿載銅損"] = extract_pure_number(val_str)

                    if val_str != "nan":
                        specs.append((label_raw, val_str))
                
                # 只有抓到有效容量且有數據的才存入
                if d["容量"] > 0:
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # 計算數據
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    all_transformer_data.append({"specs": specs, "analysis": d, "capacity": d["容量"], "usage_rate": d["負載率"]})

    if all_transformer_data:
        # 統計與 Word 產出 (維持原結構)
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        avg_usage = sum([t["usage_rate"] for t in all_transformer_data]) / len(all_transformer_data)

        st.success(f"✅ 解析完成！符合條件共 {len(all_transformer_data)} 台。")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("1. 總裝置容量", f"{total_capacity} kVA")
            st.write("**2. 設備規格分布：**")
            for k, v in sorted(cap_counts.items(), reverse=True):
                if k > 0: st.write(f"　🔹 {k} kVA × {v} 台")
        with col2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # --- 報告生成邏輯 ---
        doc = Document()
        # 壹、統計總表
        doc.add_heading('壹、 設備統計總表', 1)
        stb = doc.add_table(rows=0, cols=2); stb.style = 'Table Grid'
        def add_r(l, v):
            row = stb.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(v), 12)
        add_r("總裝置容量", f"{total_capacity} kVA")
        add_r("設備規格分布", "、".join([f"{k}kVAx{v}台" for k, v in sorted(cap_counts.items(), reverse=True)]))
        add_r("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、改善前數據分析大表
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_t = doc.add_table(rows=1, cols=11); ana_t.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_t.rows[0].cells[i].paragraphs[0].add_run(h), 9, True)
        for item in all_transformer_data:
            d = item["analysis"]
            row = ana_t.add_row().cells
            row_data = [d["建築物"], d["編號"], d["年份"], d["廠牌"], d["容量"], d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能'])}"]
            for i, val in enumerate(row_data): set_font_kai(row[i].paragraphs[0].add_run(str(val)), 8)
        doc.add_page_break()

        # 參、詳細資料
        doc.add_heading('參、 詳細設備數據', 1)
        for item in all_transformer_data:
            specs = item["specs"]
            doc.add_paragraph().add_run(f"設備詳細資料 (序號：{specs[0][1]})").bold = True
            t_dt = doc.add_table(rows=0, cols=2); t_dt.style = 'Table Grid'
            for l, v in specs:
                cells = t_dt.add_row().cells
                set_font_kai(cells[0].paragraphs[0].add_run(l), 10)
                set_font_kai(cells[1].paragraphs[0].add_run(v), 10)
            doc.add_page_break()

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button("📥 下載完整報告", output, f"Transformer_Analysis_{base_year}.docx")
