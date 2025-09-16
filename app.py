import streamlit as st
import pandas as pd
from io import BytesIO
from utils import preprocess_data, generate_recommendations

st.set_page_config(layout="wide")

st.title("貨品調貨建議系統")

st.info("請上傳 Excel 文件以開始。")

# 文件上傳
uploaded_file = st.file_uploader("上傳 Excel 文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 修正：將檔名轉為小寫以正確判斷引擎
        engine = 'openpyxl' if uploaded_file.name.lower().endswith('xlsx') else 'xlrd'
        df = pd.read_excel(uploaded_file, engine=engine)
        st.success("文件上傳成功！")

        # 顯示原始數據預覽
        with st.expander("顯示原始數據預覽"):
            st.dataframe(df.head(100))

        # 數據預處理
        processed_df, logs = preprocess_data(df.copy())

        # 顯示預處理日誌
        if logs:
            with st.expander("查看數據預處理日誌"):
                for log in logs:
                    if "錯誤" in log:
                        st.error(log)
                    elif "警告" in log:
                        st.warning(log)
                    else:
                        st.info(log)

        if processed_df is not None:
            # 生成調貨建議
            if st.button("生成調貨建議"):
                with st.spinner("正在分析數據並生成建議..."):
                    recommendations_df, summary_kpis, summary_details = generate_recommendations(processed_df.copy())

                if not recommendations_df.empty:
                    st.subheader("調貨建議")
                    st.dataframe(recommendations_df)

                    st.subheader("統計摘要")
                    # KPI 指標
                    cols = st.columns(len(summary_kpis))
                    for i, (k, v) in enumerate(summary_kpis.items()):
                        cols[i].metric(k, v)

                    # 詳細統計
                    with st.expander("查看詳細統計數據"):
                        for name, detail_df in summary_details.items():
                            st.write(f"**按 {name.replace('_', ' ')} 統計**")
                            st.dataframe(detail_df)

                    # 下載功能
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        recommendations_df.to_excel(writer, sheet_name='調貨建議', index=False)
                        # 將統計摘要寫入同一個 Excel 文件的不同工作表
                        summary_df = pd.DataFrame([summary_kpis])
                        summary_df.to_excel(writer, sheet_name='統計摘要', index=False)
                        for name, detail_df in summary_details.items():
                            detail_df.to_excel(writer, sheet_name=f'按{name}統計', index=False)

                    st.download_button(
                        label="下載調貨建議 (Excel)",
                        data=output.getvalue(),
                        file_name="reallocation_recommendations.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("根據當前規則，沒有生成任何調貨建議。")

    except Exception as e:
        st.error(f"處理文件時發生錯誤: {e}")