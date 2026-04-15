import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import io
from collections import Counter

st.set_page_config(page_title="變壓器專業報告-彙總統計版", layout="wide")
st.title("📑 變壓器自動化報告 (含設備統計總表)")

# --- 側邊欄設定區 ---
st.sidebar.header("⚙️ 設定基準與篩選")
base_year = st.sidebar.number_input("請輸入基準年份 (計算機齡用)：", min_value=1900, max_value=2100, value=2026)
age_filter = st.sidebar.selectbox(
    "選擇變壓器齡篩選：",
    ["顯示全部", "超過 10 年 (汰換參考)", "超過 15 年 (優先汰換)", "超過 20 年 (屆齡汰換)"]
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
                    mfg_year = None 
                    capacity = 0
                    usage_rate = None
                    
                    for r_offset in range(0, 50):
                        curr_r = r_start + r_offset
                        if curr_r >= len(raw_df): break
                        label_raw = str(raw_df.iloc[curr_r, c_start]).replace(' ', '').replace('\n', '')
                        
                        if r_offset > 0 and label_raw == "序號": break
                        if any(k in label_raw for k in filter_keywords): continue

                        label = str(raw_df.iloc[curr_r, c_start]).replace('\n', '').strip()
                        if (label == "nan" or label == "") and c_start > 0:
                            label = str(raw_df.iloc[curr_r, curr_r-1] if 'curr_r' in locals() else c_start-1).strip() # 修正標籤抓取
                            label = str(raw_df.iloc[curr_r, c_start-1]).replace('\n', '').strip()
                        
                        value = str(raw_df.iloc[curr_r, target_col]).strip()
                        if label == "nan" or not label or any(k in label for k in filter_keywords): continue
                        
                        # --- 數據抓取：年份、容量、利用率 ---
                        if any(k in label for k in ["製造年份", "製造日期", "出廠", "年份"]):
                            year_digits = ''.join(filter(str.isdigit, value))
                            if year_digits:
                                y = int(year_digits)
                                mfg_year = y + 1911 if y < 200 else y
                        
                        if "容量" in label and "kVA" in value.upper() or "容量" in label:
                            cap_digits = ''.join(filter(str.isdigit, value))
                            if cap_digits: capacity = int(cap_digits)
                        
                        if any(k in label for k in ["利用率", "負載率"]):
                            usage_digits = value.replace('%', '').strip()
                            try: usage_rate = float(usage_digits)
                            except: pass
                        
                        specs.append((label, value if value != "nan" else "-"))
                    
                    if specs:
                        age = base_year - mfg_year if mfg_year else 0
                        should_add = False
                        if age_filter == "顯示全部": should_add = True
                        elif "10" in age_filter and age >= 10: should_add = True
                        elif "15" in age_filter and age >= 15: should_add = True
                        elif "20" in age_filter and age >= 20: should_add = True
                        
                        if should_add:
                            # 儲存額外統計資訊
                            all_transformer_data.append({
                                "specs": specs,
                                "capacity": capacity,
                                "usage_rate": usage_rate,
                                "age": age
                            })

    if all_transformer_data:
        # --- 進行數據匯總統計 ---
        total_capacity = sum(t["capacity"] for t in all_transformer_data)
        cap_counts = Counter(t["capacity"] for t in all_transformer_data)
        valid_usages = [t["usage_rate"] for t in all_transformer_data if t["usage_rate"] is not None]
        avg_usage = sum(valid_usages) / len(valid_usages) if valid_usages else 0

        st.success(f"📊 統計摘要：總台數 {len(all_transformer_data)} | 總容量 {total_capacity} kVA | 平均利用率 {avg_usage:.2f}%")

        if st.button("🚀 生成含統計總表之報告"):
            doc = Document()
            
            # 1. 產出設備統計總表
            p_title = doc.add_paragraph()
            run_title = p_title.add_run("壹、 設備統計總表")
            set_font_kai(run_title, 16, is_bold=True)
            
            summary_table = doc.add_table(rows=0, cols=2)
            summary_table.style = 'Table Grid'
            
            # 總裝置容量
            row = summary_table.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run("總裝置容量"), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(f"{total_capacity} kVA"), 12)
            
            # 規格統計 (多少kVA x 幾台)
            row = summary_table.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run("設備規格分布"), 12, True)
            cap_detail = "、".join([f"{k}kVA x {v}台" for k, v in sorted(cap_counts.items(), reverse=True)])
            set_font_kai(row[1].paragraphs[0].add_run(cap_detail), 11)
            
            # 平均利用率
            row = summary_table.add_row().cells
            set_font_kai(row[0].paragraphs[0].add_run("平均利用率 (改善對象)"), 12, True)
            set_font_kai(row[1].paragraphs[0].add_run(f"{avg_usage:.2f} %"), 12)
            
            doc.add_page_break()

            # 2. 產出個別設備詳表
            p_detail = doc.add_paragraph()
            run_detail = p_detail.add_run("貳、 設備詳細數據")
            set_font_kai(run_detail, 16, is_bold=True)

            for t_item in all_transformer_data:
                specs = t_item["specs"]
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
            st.download_button("📥 下載完整統計報告", output, f"Transformer_Summary_{base_year}.docx")
