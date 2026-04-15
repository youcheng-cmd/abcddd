import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="數據抓取測試", layout="wide")
st.title("📊 變壓器數據抓取測試")

excel_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if excel_file:
    # 1. 讀取整個工作表（不設標題列）
    raw_df = pd.read_excel(excel_file, header=None)
    
    # 2. 尋找數據起點：找尋包含「容量」字樣的列
    header_idx = None
    for idx, row in raw_df.iterrows():
        if row.astype(str).str.contains("容量").any():
            header_idx = idx
            break
            
    if header_idx is not None:
        # 3. 重新以找到的那一列作為標題
        df = pd.read_excel(excel_file, header=header_idx + 1)
        
        # 4. 清理欄位名稱（去除空格與換行）
        df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
        
        # 5. 定義關鍵欄位名稱 (根據您的截圖)
        # 我們找：序號、容量(kVA)、負載率(%)、效率(%)
        target_cols = {
            "no": next((c for c in df.columns if "序號" in c), None),
            "cap": next((c for c in df.columns if "容量" in c), None),
            "load": next((c for c in df.columns if "負載率" in c), None),
            "eff": next((c for c in df.columns if "效率" in c), None)
        }

        # 6. 過濾數據：只保留「容量」是數字的列（排除掉後面的備註或空行）
        df[target_cols["cap"]] = pd.to_numeric(df[target_cols["cap"]], errors='coerce')
        final_df = df.dropna(subset=[target_cols["cap"]]).copy()

        st.success(f"✅ 成功定位標題列！偵測到 {len(final_df)} 台設備。")
        
        # 顯示抓取到的表格預覽
        st.write("### ⬇️ 這是程式抓取到的數據表格：")
        show_df = final_df[[target_cols["no"], target_cols["cap"], target_cols["load"], target_cols["eff"]]]
        st.table(show_df)

        if st.button("📝 將此表格輸出為 Word"):
            doc = Document()
            doc.add_heading('變壓器數據清單', 0)
            
            # 建立表格
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '序號'
            hdr_cells[1].text = '容量(kVA)'
            hdr_cells[2].text = '負載率(%)'
            hdr_cells[3].text = '效率(%)'

            for _, row in show_df.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row[target_cols["no"]])
                row_cells[1].text = str(row[target_cols["cap"]])
                row_cells[2].text = str(row[target_cols["load"]])
                row_cells[3].text = str(row[target_cols["eff"]])

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            st.download_button("📥 下載抓取的數據表格", output, "Data_List.docx")
    else:
        st.error("❌ 找不到含有『容量』的欄位，請檢查 Excel 內容。")
