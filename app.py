import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
from utils import preprocess_data, generate_recommendations

# 主應用程式
def main():
    st.title("優化後的調貨建議生成器")

    # 添加一個選項來使用預設文件路徑
    use_default_file = st.checkbox("使用預設文件路徑進行調試")

    uploaded_file = None
    if use_default_file:
        default_file_path = r"C:\Users\BestO\Dropbox\SASA\ELE_08Sep2025.XLSX"
        if os.path.exists(default_file_path):
            uploaded_file = default_file_path
            st.info(f"正在使用預設文件: {default_file_path}")
        else:
            st.error(f"預設文件不存在: {default_file_path}")
            return
    else:
        uploaded_file = st.file_uploader("上傳 Excel 文件 (.xlsx)", type="xlsx")

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("文件上傳成功！")

            # 進行數據預處理
            processed_df, logs = preprocess_data(df.copy())

            if processed_df is not None:
                # 顯示數據預覽
                st.subheader("數據預覽 (前 5 行)")
                st.dataframe(processed_df.head())

                # 顯示統計信息
                st.subheader("數據統計")
                st.write(f"總行數: {len(processed_df)}")
                st.write(f"缺失值數量: {processed_df.isnull().sum().sum()}")

                # 顯示預處理日誌
                if logs:
                    st.subheader("數據處理日誌")
                    for log in logs:
                        st.warning(log)

                if st.button("生成調貨建議"):
                    with st.spinner("正在生成建議，請稍候..."):
                        recommendations, summary_kpis, summary_details = generate_recommendations(processed_df)

                    st.subheader("調貨建議")
                    st.dataframe(recommendations)

                    # 準備下載文件
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        recommendations.to_excel(writer, index=False, sheet_name='調貨建議')
                        # 將 KPI 和詳細統計信息寫入同一工作表
                        summary_df = pd.concat([pd.DataFrame([summary_kpis]), pd.DataFrame(), summary_details['by_article'], pd.DataFrame(), summary_details['by_om'], pd.DataFrame(), summary_details['by_transfer_type'], pd.DataFrame(), summary_details['by_receive_priority']], ignore_index=True)
                        summary_df.to_excel(writer, index=False, sheet_name='統計摘要')

                    file_name = f"調貨建議_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    st.download_button(
                        label="下載調貨建議 Excel 文件",
                        data=output.getvalue(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("統計摘要")
                    st.write("**關鍵績效指標 (KPI)**")
                    st.json(summary_kpis)
                    st.write("**詳細統計**")
                    for key, df_summary in summary_details.items():
                        st.write(f"**{key}**")
                        st.dataframe(df_summary)

        except Exception as e:
            st.error(f"處理文件時發生錯誤: {e}")
    else:
        st.info("請上傳一個 .xlsx 文件。")

if __name__ == "__main__":
    main()