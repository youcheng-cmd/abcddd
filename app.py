import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="節能報告生成系統", layout="wide")
st.title("📑 變壓器節能改善報告自動化")

excel_file = st.file_uploader("1. 上傳 Excel (電能系統資料)", type=["xlsx"])
template_file = st.file_uploader("2. 上傳 Word 模板", type=["docx"])

if excel_file and template_file:
    # 讀取資料，設定標題列在第 2 行 (Python 從 0 開始算，所以 header=2)
    df = pd.read_excel(excel_file, header=2)
    st.write("### 資料預覽 (前5筆)")
    st.dataframe(df.head())

    if st.button("🚀 開始計算並生成報告"):
        try:
            doc = DocxTemplate(template_file)
            
            transformers = []
            sum_saving_kw = 0
            
            # 遍歷 Excel 的每一列
            for index, row in df.iterrows():
                # 排除空行 (如果容量是空的就跳過)
                if pd.isna(row['容量(kVA)']):
改善前損耗                    continue
                
                # --- 抓取數據與計算 ---
                cap = float(row['容量(kVA)'])
                load_factor = float(row['負載率(%)']) / 100
                eff_before = float(row['效率(%)']) / 100
                
                # 假設改善後的效率是 99% (或從 Excel 其他欄位抓取)
                eff_after = 0.99 
                
                # 簡單損耗計算公式示例：損耗 = 容量 * 負載率 * (1 - 效率)
                loss_before = cap * load_factor * (1 - eff_before)
                loss_after = cap * load_factor * (1 - eff_after)
                saving_kw = loss_before - loss_after
                
                transformers.append({
                    'no': row['變壓器編號'] if '變壓器編號' in df.columns else index + 1,
                    'cap': cap,
                    'load': f"{row['負載率(%)']}%",
                    'eff_b': f"{row['效率(%)']}%",
                    'kw_b': round(loss_before, 3),
                    'kw_a': round(loss_after, 3),
                    'saving': round(saving_kw, 3)
                })
                sum_saving_kw += saving_kw

            # --- 整理要填入 Word 的資料 ---
            context = {
                'items': transformers,
                'total_saving': round(sum_saving_kw, 2),
                'annual_saving': round(sum_saving_kw * 8760, 0) # 假設全年運轉 8760 小時
            }
            
            doc.render(context)
            
            # --- 檔案下載 ---
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.success("🎉 計算完成！")
            st.download_button(
                label="📥 下載 Word 報告",
                data=output,
                file_name="變壓器改善報告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"程式執行出錯：{e}")
            st.info("請檢查 Excel 欄位名稱是否與程式碼中的文字完全一致。")
