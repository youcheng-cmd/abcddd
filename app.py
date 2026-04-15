import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化報告 (自動抓取現況功因)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)

# 這裡改為「改善後」的預期功因設定
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

excel_file = st.file_uploader("請上傳您的 Excel", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機", "電能系統"]

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
                if str(raw_df.iloc[r_start, target_col]) in ["nan", ""]: continue
                
                # 初始化
                d = {
                    "建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0, 
                    "型式": "-", "負載率": 0.0, "鐵損": 0.0, "滿載銅損": 0.0,
                    "現況功因": 0.8  # 預設值，若有抓到會覆蓋
                }
                specs = []
                
                for r_offset in range(0, 50):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    if (label == "nan" or label == "") and c_start > 0:
                        label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                    
                    label_clean = label.replace(' ', '')
                    if r_offset > 0 and label_clean == "序號": break
                    if any(k in label_clean for k in filter_keywords): continue
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label == "nan" or not label: continue

                    # --- 自動抓取邏輯 ---
                    if "位置" in label or "建築" in label: d["建築物"] = val
                    if "編號" in label: d["編號"] = val
                    if "廠牌" in label: d["廠牌"] = val
                    if "型式" in label: d["型式"] = val
                    if any(k in label for k in ["年份", "出廠"]):
                        digits = ''.join(filter(str.isdigit, val))
                        if digits:
                            y = int(digits)
                            d["年份"] = y + 1911 if y < 200 else y
                    if "容量" in label:
                        cap_digits = ''.join(filter(str.isdigit, val))
                        if cap_digits: d["容量"] = int(cap_digits)
                    if any(k in label for k in ["利用率", "負載率"]):
                        u_raw = val.replace('%', '').strip()
                        try:
                            u_val = float(u_raw)
                            d["負載率"] = u_val * 100 if 0 < u_val < 1 else u_val
                        except: pass
                    
                    # 重要：自動抓取現況功率因數
                    if any(k in label for k in ["功率因數", "功因", "PF"]):
                        pf_raw = val.replace('%', '').strip()
                        try:
                            pf_val = float(pf_raw)
                            # 如果抓到的是 95，轉成 0.95；如果抓到 0.95 則保持
                            d["現況功因"] = pf_val / 100 if pf_val > 1 else pf_val
                        except: pass

                    if "無載損" in label or "鐵損" in label:
                        i_num = ''.join(filter(str.isdigit, val))
                        d["鐵損"] = int(i_num) if i_num else 0
                    if "負載損" in label or "銅損" in label:
                        cu_num = ''.join(filter(str.isdigit, val))
                        d["滿載銅損"] = int(cu_num) if cu_num else 0

                    specs.append((label, val))
                
                if specs:
                    age = base_year - d["年份"] if d["年份"] else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    # 計算改善前數據 (使用抓到的現況功因)
                    d["實際銅損"] = d["滿載銅損"] * ((d["負載率"]/100)**2)
                    d["現況輸出功率"] = d["容量"] * d["現況功因"] * (d["負載率"]/100)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    all_transformer_data.append({"specs": specs, "analysis": d, "capacity": d["容量"], "usage_rate": d["負載率"]})

    if all_transformer_data:
        # 摘要數據
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        valid_usages = [t["usage_rate"] for t in all_transformer_data if t["usage_rate"] > 0]
        avg_usage = sum(valid_usages) / len(valid_usages) if valid_usages else 0

        st.success(f"✅ 解析完成！已自動抓取每台設備之現況功因。")
        
        # 顯示預覽表
        df_preview = pd.DataFrame([t["analysis"] for t in all_transformer_data])
        st.write("### 📋 改善前分析預覽 (包含各台抓取之功因)")
        st.dataframe(df_preview[["建築物", "編號", "現況功因", "現況輸出功率", "改善前耗能"]])

        # Word 生成部分 (略，結構同前)
        # ... (後續生成 Word 邏輯，並在表格內顯示現況功因)
