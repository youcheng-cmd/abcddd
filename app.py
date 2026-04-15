import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="變壓器數據抓取工具", layout="wide")
st.title("🔍 變壓器 Excel 數據精確抓取")

excel_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取整個 Excel 頁面
    all_data = pd.read_excel(excel_file, header=None)
    
    # 步驟 1: 尋找標題列（掃描包含「容量」二字的儲存格）
    target_row = None
    for idx, row in all_data.iterrows():
        if row.astype(str).str.contains("容量").any():
            target_row = idx
            break
    
    if target_row is not None:
        # 步驟 2: 重新定義 Dataframe，將該列設為標題
        df = pd.read_excel(excel_file, header=target_row + 1)
        
        # 清洗欄位名稱（去除換行符、空格）
        df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
        
        # 步驟 3: 定義我們要找的欄位對應（模糊搜尋）
        def find_col(keywords):
            for col in df.columns:
                if any(k in col for k in keywords):
                    return col
            return None

        col_cap = find_col(["容量", "kVA"])
        col_load = find_col(["負載率", "%"])
        col_eff = find_col(["效率", "%"])
        col_no = find_col(["序號", "編號"])

        # 步驟 4: 數據清洗 - 只保留「容量」那一欄有數字的行
        # 將容量轉為數字，無法轉換的會變成 NaN，然後刪除 NaN
        df[col_cap] = pd.to_numeric(df[col_cap], errors='coerce')
        final_df = df.dropna(subset=[col_cap]).copy()

        st.success(f"✅ 已成功定位！找到 {len(final_df)} 台變壓器數據。")
        
        # 網頁預覽
        display_cols = [c for c in [col_no, col_cap, col_load, col_eff] if c]
        st.write("### 🤖 程式抓取到的數據預覽：")
        st.dataframe(final_df[display_cols])

        if st.button("🚀 確認無誤，產生報告"):
            try:
                doc = Document()
                doc.add_heading('變壓器高效率改善建議表', 0)
                
                # 建立 10 欄表格（符合你截圖的需求）
                table = doc.add_table(rows=1, cols=10)
                table.style = 'Table Grid'
                headers = ['序號', '容量', '負載率', '前鐵損', '前銅損', '前總損', '後鐵損', '後銅損', '後總損', '節省kW']
                for i, h in enumerate(headers):
                    table.rows[0].cells[i].text = h

                for _, row in final_df.iterrows():
                    # 抓取數值並處理
                    cap = float(row[col_cap])
                    load_p = float(str(row[col_eff]).replace('%','')) / 100 if col_load else 0.5
                    eff_b = float(str(row[col_eff]).replace('%','')) / 100
                    
                    # 計算邏輯 (公式)
                    loss_b = cap * (1 - eff_b)
                    iron_b = loss_b * 0.25
                    copper_b = (loss_b - iron_b) * (load_p**2)
                    total_b = iron_b + copper_b
                    
                    # 改善後 (假設 99.2% 效率)
                    eff_a = 0.992
                    loss_a = cap * (1 - eff_a)
                    iron_a = loss_a * 0.20
                    copper_a = (loss_a - iron_a) * (load_p**2)
                    total_a = iron_a + copper_a

                    cells = table.add_row().cells
                    cells[0].text = str(row[col_no]) if col_no else "-"
                    cells[1].text = str(cap)
                    cells[2].text = f"{int(load_p*100)}%"
                    cells[3].text = str(round(iron_b, 3))
                    cells[4].text = str(round(copper_b, 3))
                    cells[5].text = str(round(total_b, 3))
                    cells[6].text = str(round(iron_a, 3))
                    cells[7].text = str(round(copper_a, 3))
                    cells[8].text = str(round(total_a, 3))
                    cells[9].text = str(round(total_b - total_a, 3))

                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                st.download_button("📥 下載產出的報告", output, "Transformer_Report.docx")
                
            except Exception as e:
                st.error(f"計算或生成 Word 時出錯：{e}")
    else:
        st.error("❌ 找不到含有『容量』的標題列。請確認 Excel 內容是否有這兩個字。")
