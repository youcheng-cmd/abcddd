import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="變壓器縱向數據抓取", layout="wide")
st.title("📑 變壓器數據自動化 (按序號分組抓取)")

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 讀取 Excel 原始資料
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 1. 定位「序號」所在的座標 (row, col)
    sn_row, sn_col = None, None
    for r in range(len(raw_df)):
        for c in range(len(raw_df.columns)):
            if "序號" in str(raw_df.iloc[r, c]):
                sn_row, sn_col = r, c
                break
        if sn_row is not None: break

    if sn_row is not None:
        # 2. 確定有多少個序號（從序號格子往右看，有數字的就是一組變壓器）
        transformer_columns = []
        for c in range(sn_col + 1, len(raw_df.columns)):
            val = raw_df.iloc[sn_row, c]
            if pd.notna(val) and str(val).strip().isdigit():
                transformer_columns.append(c)
        
        # 3. 抓取每一組變壓器的資訊
        # 掃描從序號那一列開始往下的所有資料
        data_results = []
        
        for col_idx in transformer_columns:
            transformer_data = {}
            # 遍歷每一列，左邊是「標題」，右邊是「該序號的數值」
            for r in range(sn_row, len(raw_df)):
                # 取得左側的標題（可能在 sn_col 或往左幾格）
                # 這裡尋找同一行中，序號左邊最靠近的文字描述
                label = ""
                for search_c in range(sn_col, -1, -1):
                    if pd.notna(raw_df.iloc[r, search_c]):
                        label = str(raw_df.iloc[r, search_c]).replace('\n', '').strip()
                        break
                
                if label:
                    val = raw_df.iloc[r, col_idx]
                    transformer_data[label] = val
            
            data_results.append(transformer_data)

        # 4. 轉換為表格預覽
        final_df = pd.DataFrame(data_results)
        
        # 將標題整理乾淨（如果標題重複，只保留有意義的部分）
        st.success(f"✅ 成功辨識！找到 {len(transformer_columns)} 組變壓器數據。")
        
        st.write("### 🤖 程式抓取到的各組變壓器數據：")
        st.dataframe(final_df)

        if st.button("🚀 產出 Word 資料表"):
            doc = Document()
            doc.add_heading('變壓器電能系統資料清單', 0)
            
            # 因為欄位很多，我們轉置表格讓它好讀（或是產出你指定的格式）
            table = doc.add_table(rows=1, cols=len(final_df.columns))
            table.style = 'Table Grid'
            
            # 填入標題
            for i, col_name in enumerate(final_df.columns):
                table.rows[0].cells[i].text = str(col_name)
            
            # 填入數據
            for _, row in final_df.iterrows():
                row_cells = table.add_row().cells
                for i, val in enumerate(row):
                    row_cells[i].text = str(val) if pd.notna(val) else ""

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載 Word 報告", output, "Transformer_Groups.docx")
            
    else:
        st.error("❌ 無法定位到『序號』，請確認 Excel 中有這兩個字。")
