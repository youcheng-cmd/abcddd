import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="Excel轉Word工具", layout="centered")

st.title("📄 Excel 轉 Word 自動報告生成器")
st.markdown("上傳 Excel 與 Word 模板，快速生成報告！")

# 1. 上傳檔案區
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("上傳 Excel 資料表", type=["xlsx"])
with col2:
    template_file = st.file_uploader("上傳 Word 模板", type=["docx"])

if excel_file and template_file:
    # 讀取 Excel
    df = pd.read_excel(excel_file)
    st.success("Excel 讀取成功！")
    st.dataframe(df.head(5)) # 顯示前五筆資料
    
    # 選擇要產生的資料行
    row_to_gen = st.selectbox("請選擇要產生的資料行 (根據索引)", df.index)
    
    if st.button("🚀 開始生成報告"):
        try:
            # 讀取模板
            doc = DocxTemplate(template_file)
            
            # 將該列資料轉為字典
            context = df.iloc[row_to_gen].to_dict()
            
            # 渲染 Word (把 {{變數}} 換掉)
            doc.render(context)
            
            # 轉換為下載流
            target_stream = io.BytesIO()
            doc.save(target_stream)
            target_stream.seek(0)
            
            st.download_button(
                label="✅ 下載產出的報告",
                data=target_stream,
                file_name=f"report_{row_to_gen}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"發生錯誤：{e}")
