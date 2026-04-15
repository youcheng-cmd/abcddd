import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="節能報告工具", layout="wide")
st.title("📑 變壓器報告自動化生成")

excel_file = st.file_uploader("1. 上傳 Excel", type=["xlsx"])
template_file = st.file_uploader("2. 上傳 Word 模板", type=["docx"])

if excel_file and template_file:
    df = pd.read_excel(excel_file, header=2) # 假設標題在第3行
    st.write("資料預覽：", df.head())

    if st.button("🚀 產出報告"):
        try:
            doc = DocxTemplate(template_file)
            items = []
            for idx, row in df.iterrows():
                if pd.isna(row.get('容量(kVA)')):
                    continue
                items.append({
                    'no': idx + 1,
                    'cap': row.get('容量(kVA)', 0),
                    'load': row.get('負載率(%)', 0),
                    'eff': row.get('效率(%)', 0)
                })
            
            doc.render({'items': items})
            output = io.BytesIO()
            doc.save(output)
            st.download_button("📥 下載報告", output.getvalue(), "report.docx")
            st.success("完成！")
        except Exception as e:
            st.error(f"錯誤：{e}")
