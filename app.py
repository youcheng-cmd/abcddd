import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import io
import re
from collections import Counter

st.set_page_config(page_title="變壓器節能分析系統", layout="wide")
st.title("📑 變壓器自動化分析報告 (鋼鐵基準版)")

# --- 1. 側邊欄參數設定 ---
st.sidebar.header("⚙️ 參數設定")
base_year = st.sidebar.number_input("請輸入基準年份：", value=2026)
pf_after_input = st.sidebar.number_input("設定【改善後】目標功率因數 (%)：", value=95)
age_filter = st.sidebar.selectbox("選擇變壓器齡篩選：", ["顯示全部", "超過 10 年", "超過 15 年", "超過 20 年"])

# --- 通用工具函數 ---
def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def extract_number(text):
    if pd.isna(text): return 0.0
    s = str(text).replace(',', '').replace(' ', '').strip()
    nums = re.findall(r"[-+]?\d*\.\d+|\d+", s)
    return float(nums[0]) if nums else 0.0

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取 Excel 內容
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    # 【關鍵】防止台數重複的紀錄器
    seen_device_keys = set() 

    for sheet_name, raw_df in all_sheets.items():
        # 尋找「序號」錨點
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_val = str(raw_df.iloc[r, c]).replace(' ', '')
                if cell_val == "序號":
                    # 確認下方是否有相關標籤，避免誤抓非規格表的「序號」
                    if r + 1 < len(raw_df):
                        check_label = str(raw_df.iloc[r+1, c]).replace(' ', '')
                        if any(k in check_label for k in ["建築", "位置", "編號"]):
                            anchors.append((r, c))
        
        for r_start, c_start in anchors:
            # 橫向掃描數據欄位 (TR-1 ~ TR-7)
            for offset in range(1, 15): # 稍微放寬掃描範圍確保 TR-7 在內
                target_col = c_start + offset
                if target_col >= len(raw_df.columns): break
                
                # 檢查該欄是否有內容（序號數字）
                sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                if sn_val in ["nan", ""] or not sn_val.replace('.0','').isdigit(): continue
                
                # 初始化單台設備字典
                d = {"建築物": "-", "編號": "-", "年份": 0, "廠牌": "-", "容量": 0.0, 
                     "型式": "-", "負載率": 0.0, "現況功因": 0.0, "鐵損": 0.0, "滿載銅損": 0.0}
                specs = []
                
                # 垂直掃描該設備的所有屬性
                for r_offset in range(0, 55):
                    curr_r = r_start + r_offset
                    if curr_r >= len(raw_df): break
                    
                    l1 = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                    l2 = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip() if c_start > 0 else ""
                    label = l1 if (l1 != "nan" and l1 != "") else l2
                    label_p = label.replace(' ', '')
                    
                    # 碰到下一個序號標籤就停止掃描此台設備
                    if r_offset > 0 and label_p == "序號": break
                    
                    val = str(raw_df.iloc[curr_r, target_col]).strip()
                    if label_p == "nan" or not label_p: continue

                    # 數據精準抓取
                    if any(k in label_p for k in ["建築", "位置"]): d["建築物"] = val
                    if "編號" in label_p: d["編號"] = val
                    if "廠牌" in label_p: d["廠牌"] = val
                    if "型式" in label_p: d["型式"] = val
                    if any(k in label_p for k in ["年份", "出廠"]):
                        n = extract_number(val)
                        d["年份"] = int(n) + 1911 if 0 < n < 200 else int(n)
                    if "容量" in label_p: d["容量"] = extract_number(val)
                    if any(k in label_p for k in ["利用率", "負載率"]):
                        n = extract_number(val)
                        d["負載率"] = n * 100 if 0 < n < 1 else n
                    
                    # --- 1. 強化功因抓取 (精確對應 0.98) ---
                    if label_p == "功因" or any(k in label_p for k in ["功率因數", "PF", "P.F"]):
                        n = extract_number(val)
                        if n > 0: d["現況功因"] = n / 100 if n > 1 else n

                    specs.append((label, val))
                
                # 只有容量大於 0 的才算有效設備
                if d["容量"] > 0:
                    # 【核心】建立唯一 ID 防止台數重複
                    device_key = f"{sheet_name}_{d['編號']}_{d['容量']}"
                    if device_key in seen_device_keys: continue
                    
                    # 2. 功因防錯：沒抓到才補 0.8
                    if d["現況功因"] <= 0: d["現況功因"] = 0.8

                    # 3. 損耗公式計算
                    base_iron_loss = d["容量"] * 2.5    # 鐵損估算 (W)
                    base_copper_loss = d["容量"] * 13.0 # 滿載銅損估算 (W)
                    d["鐵損"] = base_iron_loss
                    d["實際銅損"] = base_copper_loss * ((d["負載率"] / 100) ** 2)
                    d["改善前耗能"] = (d["鐵損"] + d["實際銅損"]) * 8760 / 1000
                    
                    # 機齡過濾
                    age = base_year - d["年份"] if d["年份"] > 0 else 0
                    if (age_filter == "超過 10 年" and age < 10) or \
                       (age_filter == "超過 15 年" and age < 15) or \
                       (age_filter == "超過 20 年" and age < 20): continue
                    
                    all_transformer_data.append({"specs": specs, "analysis": d})
                    seen_device_keys.add(device_key)

    if all_transformer_data:
        # --- 網頁摘要 ---
        caps = [t["analysis"]["容量"] for t in all_transformer_data]
        total_cap = sum(caps)
        cap_dist = Counter(caps)
        avg_usage = sum(t["analysis"]["負載率"] for t in all_transformer_data) / len(all_transformer_data)

        st.success(f"✅ 解析完成！符合條件共 {len(all_transformer_data)} 台")
        c1, c2 = st.columns(2)
        with c1:
            st.metric("1. 總裝置容量", f"{total_cap:,.0f} kVA")
            st.write("**2. 規格台數分布：**")
            for k, v in sorted(cap_dist.items(), reverse=True):
                st.write(f"　🔹 {k:,.0f} kVA × {v} 台")
        with c2:
            st.metric("3. 平均負載利用率", f"{avg_usage:.2f} %")

        # --- Word 生成 (壹、貳、參) ---
        doc = Document()
        # 壹、總表
        doc.add_heading('壹、 設備統計總表', 1)
        st_table = doc.add_table(rows=0, cols=2); st_table.style = 'Table Grid'
        def add_sum(l, v):
            r = st_table.add_row().cells
            set_font_kai(r[0].paragraphs[0].add_run(l), 12, True)
            set_font_kai(r[1].paragraphs[0].add_run(v), 12)
        add_sum("總裝置容量", f"{total_cap:,.0f} kVA")
        add_sum("設備規格分布", "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_dist.items(), reverse=True)]))
        add_sum("平均負載利用率", f"{avg_usage:.2f} %")
        doc.add_page_break()

        # 貳、分析大表 (11欄)
        doc.add_heading('貳、 變壓器設備改善前數據分析表', 1)
        ana_table = doc.add_table(rows=1, cols=11); ana_table.style = 'Table Grid'
        headers = ["建築物", "編號", "年份", "廠牌", "容量", "型式", "負載率", "現況功因", "銅損(W)", "鐵損(W)", "改善前耗能"]
        for i, h in enumerate(headers): set_font_kai(ana_table.rows[0].cells[i].paragraphs[0].add_run(h), 8, True)
        for t in all_transformer_data:
            d = t["analysis"]
            row = ana_table.add_row().cells
            row_vals = [d["建築物"], d["編號"], d["年份"], d["廠牌"], f"{d['容量']:.0f}", d["型式"], f"{d['負載率']:.1f}%", f"{d['現況功因']:.2f}", f"{d['實際銅損']:.1f}", f"{d['鐵損']:.1f}", f"{int(d['改善前耗能']):,}"]
            for i, v in enumerate(row_vals): set_
