import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from datetime import datetime

st.set_page_config(page_title="變壓器專業報告-年份篩選版", layout="wide")
st.title("📑 變壓器自動化報告 (含年份篩選功能)")

# 獲取今年年份
current_year = datetime.now().year

# 側邊欄：功能選單
st.sidebar.header("🔍 篩選與設定")
age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年 (汰換參考)", "超過 15 年 (優先汰換)"]
)

def set_font_kai(run, size=11, is_bold=False):
    run.font.name = '標楷體'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

excel_file = st.file_uploader("請上傳您的 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    all_transformer_data = []
    filter_keywords = ["註：", "註:", "1.", "2.", "3.", "4.", "變壓器型式請", "各迴路", "總盤抄表", "緊急發電機", "電能系統"]

    for sheet_name, raw_df in all_sheets.items():
        anchors = []
        for r in range(len(raw_df)):
            for c in range(len(raw_df.columns)):
                cell_text = str(raw_df.iloc[r, c]).replace(' ', '').replace('\n', '')
                if cell_text == "序號":
                    anchors.append((r, c))
        
        if anchors:
            for r_start, c_start in anchors:
                for offset in range(1, 10):
                    target_col = c_start + offset
                    if target_col >= len(raw_df.columns): break
                    sn_val = str(raw_df.iloc[r_start, target_col]).strip()
                    if sn_val == "nan" or sn_val == "": continue
                    
                    specs = []
                    mfg_year = None # 用來記錄製造年份
                    
                    for r_offset in range(0, 50):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        label_raw = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                        
                        if r_offset > 0 and label_raw == "序號": break
                        if any(k in label_raw for k in filter_keywords): continue

                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label or any(k in label for k in filter_keywords): continue
                        
                        # --- 年份識別邏輯 ---
                        if "出廠" in label or "製造日期" in label or "年份" in label:
                            try:
                                # 嘗試抓取數字（例如從 '2010年' 或 '民國99年' 抓取數字）
                                year_match = pd.to_numeric(''.join(filter(str.isdigit, value)))
                                # 處理民國與西元切換
                                if year_match < 200: # 假設是民國
                                    mfg_year = year_match + 1911
                                else:
                                    mfg_year = year_match
                            except:
                                pass
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    # --- 執行篩選過濾 ---
                    if specs:
                        age = current_year - mfg_year if mfg_year else 0
                        
                        if age_filter == "超過 10 年 (汰換參考)":
                            if age >= 10: all_transformer_data.append(specs)
                        elif age_filter == "超過 15 年 (優先汰換)":
                            if age >= 15: all_transformer_data.append(specs)
                        else:
                            all_transformer_data.append(specs)

    # 顯示結果
    if all_transformer_data:
        st.success(f"📊 在「{age_filter}」條件下，共找到 {len(all_transformer_data)} 台設備。")
        
        if st.button("🚀 生成篩選後的 Word 報告"):
            doc = Document()
            for specs in all_transformer_data:
                p = doc.add_paragraph()
                run_t = p.add_run(f"變壓器設備資料 (序號：{specs[0][1]})")
                set_font_kai(run_t, 14, is_bold=True)
                
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                for label, value in specs:
                    cells = table.add_row().cells
                    set_font_kai(cells[0].paragraphs[0].add_run(label), 10)
                    set_font_kai(cells[1].paragraphs[0].add_run(value), 10)
                doc.add_page_break()

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載篩選報告", output, f"Transformer_{age_filter}.docx")
    else:
        st.warning(f"⚠️ 在「{age_filter}」條件下，沒有找到符合的變壓器。")
