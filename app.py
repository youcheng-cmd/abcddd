import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

st.set_page_config(page_title="變壓器數據抓取工具", layout="wide")
st.title("🔍 變壓器 Excel 數據精確抓取")

excel_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if excel_file:
    # 1. 讀取原始 Excel (不設定 header，手動掃描)
    raw_df = pd.read_excel(excel_file)
    
    # 2. 自動尋找包含「容量」二字的列作為標題列
    header_row_index = None
    for i, row in raw_df.iterrows():
        if "容量" in str(row.values):
            header_row_index = i
            break
            
    if header_row_index is not None:
        # 重新讀取，以找到的那一列為標題
        df = pd.read_excel(excel_file, header=header_row_index + 1)
        
        # 清洗欄位名稱：去除換行符號與空白，確保程式抓得到
        df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
        
        # 定義我們要抓取的關鍵欄位 (模糊匹配)
        col_map = {
            'cap': next((c for c in df.columns if "容量" in c), None),
            'load': next((c for c in df.columns if "負載率" in c), None),
            'eff': next((c for c in df.columns if "效率" in c), None),
            'no': next((c for c in df.columns if "序號" in c or "編號" in c), None)
        }

        st.success(f"✅ 已定位標題列（第 {header_row_index + 2} 列）")
        
        # 3. 過濾掉空值，只顯示有數據的列
        final_df = df[df[col_map['cap']].notna()].copy()
        
        st.write("### 🤖 程式抓取到的關鍵數據確認：")
        st.dataframe(final_df[[col_map['no'], col_map['cap'], col_map['load'], col_map['eff']]])

        if st.button("🚀 確認數據無誤，產生報告"):
            try:
                doc = Document()
                doc.add_heading('變壓器節能改善數據報告', 0)
                
                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = '序號'
                hdr_cells[1].text = '容量(kVA)'
                hdr_cells[2].text = '負載率(%)'
                hdr_cells[3].text = '改善前效率'
                hdr_cells[4].text = '改善前損耗(kW)'
                hdr_cells[5].text = '節省電力(kW)'

                for _, row in final_df.iterrows():
                    # 轉換數值
                    c = float(row[col_map['cap']])
                    l = float(row[col_map['load']]) / 100
                    e = float(row[col_map['eff']]) / 100
                    
                    # 簡單計算改善前損耗 (kW)
                    loss_b = c * l * (1 - e)
                    # 假設改善後為一級能效 (約節省 15% 損耗)
                    saving = loss_b * 0.15 

                    row_cells = table.add_row().cells
                    row_cells[0].text = str(row[col_map['no']])
                    row_cells[1].text = str(c)
                    row_cells[2].text = f"{row[col_map['load']]}%"
                    row_cells[3].text = f"{row[col_map['eff']]}%"
                    row_cells[4].text = str(round(loss_b, 3))
                    row_cells[5].text = str(round(saving, 3))

                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                st.download_button("📥 下載產出的 Word 報告", output, "Report.docx")
                
            except Exception as e:
                st.error(f"產生報告時出錯：{e}")
                st.write("目前抓到的欄位名稱：", list(df.columns))
    else:
        st.error("❌ 在 Excel 中找不到包含『容量』的標題列，請確認 Excel 格式。")
