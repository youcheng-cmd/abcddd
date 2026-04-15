import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io

st.set_page_config(page_title="變壓器節能計算工具", layout="wide")
st.title("⚡ 變壓器高效率改善報告生成器")

excel_file = st.file_uploader("1. 上傳 Excel (電能系統資料)", type=["xlsx"])
template_file = st.file_uploader("2. 上傳 Word 模板", type=["docx"])

if excel_file and template_file:
    # 讀取 Excel，header=2 對應 Excel 第 3 列標題
    df = pd.read_excel(excel_file, header=2)
    st.success("✅ Excel 讀取成功！")
    st.dataframe(df.head())

    if st.button("🚀 執行計算並產出報告"):
        try:
            doc = DocxTemplate(template_file)
            items = []
            
            for idx, row in df.iterrows():
                if pd.isna(row.get('容量(kVA)')): continue
                
                # --- 抓取原始數據 ---
                cap = float(row.get('容量(kVA)', 0))
                load_p = float(row.get('負載率(%)', 0)) / 100  # 負載率
                eff_b_p = float(row.get('效率(%)', 0)) / 100    # 改善前效率
                
                # --- 改善前計算 (Before) ---
                # 總損耗 P = 容量 * (1 - 效率) / 效率 (因效率定義為 輸出/輸入)
                # 為了簡化與對應截圖，假設 總損 = 容量 * (1-效率)
                total_loss_b = cap * (1 - eff_b_p) 
                
                # 分配鐵損與銅損 (通常舊型變壓器鐵損約佔總損 20%-30%)
                iron_b = total_loss_b * 0.25 
                copper_b = (total_loss_b - iron_b) * (load_p**2)
                full_total_b = iron_b + copper_b

                # --- 改善後計算 (After - 假設採用一級能效等級) ---
                # 這裡使用一般高效率變壓器基準值，亦可改為從 Excel 抓
                eff_a_p = 0.99 
                total_loss_a = cap * (1 - eff_a_p)
                iron_a = total_loss_a * 0.20  # 高效率變壓器鐵損較低
                copper_a = (total_loss_a - iron_a) * (load_p**2)
                full_total_a = iron_a + copper_a
                
                # 節省電力
                saving_kw = full_total_b - full_total_a
                
                items.append({
                    'no': row.get('序號', idx + 1),
                    'cap': cap,
                    'load': f"{row.get('負載率(%)', 0)}%",
                    'eff_b': f"{row.get('效率(%)', 0)}%",
                    'iron_b': round(iron_b, 3),
                    'cop_b': round(copper_b, 3),
                    'tot_b': round(full_total_b, 3),
                    'iron_a': round(iron_a, 3),
                    'cop_a': round(copper_a, 3),
                    'tot_a': round(full_total_a, 3),
                    'saving': round(saving_kw, 3)
                })

            # --- 填入 Word 變數 ---
            context = {
                'items': items,
                'total_saving': round(sum(i['saving'] for i in items), 3),
                'year_saving': round(sum(i['saving'] for i in items) * 8760, 0)
            }
            
            doc.render(context)
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.download_button("📥 下載專業報告", output, "變壓器改善報告.docx")
            st.balloons()
            
        except Exception as e:
            st.error(f"計算錯誤：{e}")
