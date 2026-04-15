import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="變壓器節能報告生成器", layout="wide")
st.title("⚡ 變壓器高效率改善自動化報告")

# 只需要上傳 Excel
excel_file = st.file_uploader("請上傳 Excel (電能系統資料)", type=["xlsx"])

if excel_file:
    # 讀取 Excel，header=2 是因為標題通常在第 3 行
    df = pd.read_excel(excel_file, header=2)
    st.success("✅ Excel 讀取成功！")
    st.write("資料預覽：")
    st.dataframe(df.head())

    if st.button("🚀 執行計算並產出 Word 報告"):
        try:
            # 建立新的 Word 文件
            doc = Document()
            
            # 設定標題
            title = doc.add_heading('變壓器高效率改善建議表 (附件 A-3)', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 建立表格
            table = doc.add_table(rows=1, cols=10)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            headers = ['序號', '容量\n(kVA)', '負載率\n(%)', '改善前\n鐵損(kW)', '改善前\n銅損(kW)', '改善前\n總損(kW)', '改善後\n鐵損(kW)', '改善後\n銅損(kW)', '改善後\n總損(kW)', '節省電力\n(kW)']
            for i, h in enumerate(headers):
                hdr_cells[i].text = h

            total_saving_kw = 0
            
            # 遍歷 Excel 進行計算
            for idx, row in df.iterrows():
                if pd.isna(row.get('容量(kVA)')): continue
                
                # --- 抓取數據 ---
                cap = float(row.get('容量(kVA)', 0))
                load_p = float(row.get('負載率(%)', 0)) / 100
                eff_b_p = float(row.get('效率(%)', 0)) / 100
                
                # --- 改善前計算 ---
                # 總損 P = 容量 * (1-效率)
                loss_b = cap * (1 - eff_b_p)
                iron_b = loss_b * 0.25 # 預設鐵損佔 25%
                copper_b = (loss_b - iron_b) * (load_p**2)
                total_loss_b = iron_b + copper_b
                
                # --- 改善後計算 (高效率變壓器基準) ---
                eff_a_p = 0.992 # 預設改善後效率 99.2%
                loss_a = cap * (1 - eff_a_p)
                iron_a = loss_a * 0.20 # 高效率鐵損較低
                copper_a = (loss_a - iron_a) * (load_p**2)
                total_loss_a = iron_a + copper_a
                
                saving = total_loss_b - total_loss_a
                total_saving_kw += saving

                # 填入表格
                row_cells = table.add_row().cells
                row_cells[0].text = str(int(row.get('序號', idx + 1)))
                row_cells[1].text = str(cap)
                row_cells[2].text = f"{row.get('負載率(%)', 0)}%"
                row_cells[3].text = str(round(iron_b, 3))
                row_cells[4].text = str(round(copper_b, 3))
                row_cells[5].text = str(round(total_loss_b, 3))
                row_cells[6].text = str(round(iron_a, 3))
                row_cells[7].text = str(round(copper_a, 3))
                row_cells[8].text = str(round(total_loss_a, 3))
                row_cells[9].text = str(round(saving, 3))

            # 新增總計資訊
            doc.add_paragraph(f"\n總節省電力：{round(total_saving_kw, 3)} kW")
            doc.add_paragraph(f"估計年節電量 (8760小時)：{round(total_saving_kw * 8760, 0)} kWh/年")

            # 轉存為下載流
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            st.success("🎉 報告已自動生成！")
            st.download_button(
                label="📥 下載產出的 Word 報告",
                data=output,
                file_name="變壓器改善數據報告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{e}")
