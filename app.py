import streamlit as st
import pandas as pd
import time
from io import BytesIO
from utils import (
    preprocess_data, 
    estimate_transfer_potential,
    generate_recommendations, 
    create_om_transfer_chart, 
    generate_excel_export
)

# 1. 頁面配置
st.set_page_config(
    page_title="調貨建議生成系統",
    page_icon="📦",
    layout="wide"
)

def apply_ui_style():
    st.markdown(
        "<link href='https://fonts.googleapis.com/icon?family=Material+Icons' rel='stylesheet'>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <style>
        :root {
          --color-bg: #F7F7F7;
          --color-surface: #FFFFFF;
          --color-border: #E5E5E5;
          --color-text: #333333;
          --color-text-muted: #666666;
          --color-text-disabled: #999999;
          --color-brand: #2B6CB0;
          --radius-4: 4px;
          --radius-8: 8px;
          --elev-2: 0 2px 8px rgba(0,0,0,0.08);
          --elev-3: 0 4px 12px rgba(0,0,0,0.10);
          --elev-4: 0 8px 24px rgba(0,0,0,0.12);
        }
        [data-testid="stAppViewContainer"] { background: var(--color-bg); }
        [data-testid="stSidebar"] { background: var(--color-surface); border-right: 1px solid var(--color-border); }
        h1, h2, h3, h4, h5 { color: var(--color-text); }
        h3, h4 { margin-top: 16px; margin-bottom: 12px; }
        hr { border: none; border-top: 1px solid var(--color-border); margin: 24px 0; }
        div.stButton > button {
          background: var(--color-brand);
          color: #fff;
          border-radius: var(--radius-8);
          box-shadow: var(--elev-3);
          border: none;
        }
        div.stButton > button:hover { filter: brightness(0.92); }
        div.stDownloadButton > button {
          background: var(--color-brand);
          color: #fff;
          border-radius: var(--radius-8);
          box-shadow: var(--elev-2);
          border: none;
        }
        .stTextInput > div > div > input, .stSelectbox > div > div {
          border-radius: var(--radius-4);
          background: var(--color-surface);
          border: 1px solid var(--color-border);
        }
        /* Radio group去除不合背景的盒框，改用品牌選取色 */
        .stRadio > div {
          background: transparent !important;
          border: none !important;
          box-shadow: none !important;
          padding: 0 !important;
        }
        .stRadio label { color: var(--color-text); margin: 4px 0; }
        input[type="radio"] { accent-color: var(--color-brand); }
        div[data-testid="stMetric"] {
          background: var(--color-surface);
          border-radius: var(--radius-8);
          box-shadow: var(--elev-2);
          border: 1px solid var(--color-border);
          padding: 16px;
          margin-bottom: 24px;
        }
        .stDataFrame {
          background: var(--color-surface);
          border-radius: var(--radius-8);
          box-shadow: var(--elev-2);
          border: 1px solid var(--color-border);
          padding: 8px;
          margin-bottom: 24px;
        }
        div[data-testid="stHorizontalBlock"] { column-gap: 24px; row-gap: 24px; }
        div[data-testid="stVerticalBlock"] { row-gap: 24px; }
        [data-testid="stAppViewContainer"] .block-container { max-width: 1200px; }
        .block-container { padding-top: 24px; padding-bottom: 24px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def render_summary_panel(kpi_metrics):
    items = []
    mapping = {
        "總調貨建議數量": "總調貨建議行數",
        "總調貨件數": "總調貨件數",
        "涉及產品數量": "涉及產品數量",
        "涉及OM數量": "涉及OM數量",
    }
    for key, label in mapping.items():
        if key in kpi_metrics:
            items.append((label, kpi_metrics[key]))
    if not items:
        return
    rows = []
    for i, (label, value) in enumerate(items):
        rows.append(f"""
        <div class='summary-item'>
          <div class='label'>{label}</div>
          <div class='value'>{value}</div>
        </div>
        """)
    html = """
    <style>
    .summary-item { display:flex; align-items:center; border:1px solid var(--color-border); border-radius:8px; overflow:hidden; box-shadow: var(--elev-2); margin-bottom:8px; }
    .summary-item .label { flex:1; background:#E9F2FF; color:var(--color-text); padding:8px 12px; }
    .summary-item .value { width:120px; background:#DFF1E0; color:var(--color-text); font-weight:600; text-align:center; padding:8px 12px; }
    </style>
    <div class='summary-panel'>%s</div>
    """ % "".join(rows)
    st.markdown(html, unsafe_allow_html=True)

apply_ui_style()

# 2. 側邊欄設計
with st.sidebar:
    st.header("系統資訊")
    st.info(""" 
    **版本：v1.7** 
    **開發者:Ricky** 
    
    **核心功能：**  
    - ✅ ND/RF類型智慧識別 
    - ✅ 優先順序調貨匹配 
    - ✅ RF過剩轉出限制 
    - ✅ 統計分析和圖表 
    - ✅ Excel格式匯出 
    """)
    st.sidebar.header("操作指引")
    st.sidebar.markdown("""
    1.  **上傳 Excel 文件**：點擊瀏覽文件或拖放文件到上傳區域。
    2.  **啟動分析**：點擊「啟動分析」按鈕開始處理。
    3.  **查看結果**：在主頁面查看KPI、建議和圖表。
    4.  **下載報告**：點擊下載按鈕獲取 Excel 報告。
    """)

# 新增：自動化調撥頁面
with st.sidebar.expander("自動化調撥", expanded=False):
    st.info("此功能將很快推出！")

# 3. 頁面頭部
st.title("📦 調貨建議生成系統")
st.markdown("---")

# 4. 主要區塊
# 4.1. 資料上傳區塊
st.header("1. 資料上傳")
uploaded_file = st.file_uploader(
    "請上傳包含庫存和銷量數據的 Excel 文件",
    type=["xlsx", "xls"],
    help="必需欄位：Article, Article Description, RP Type, Site, OM, SaSa Net Stock, Pending Received, Safety Stock, Last Month Sold Qty, MTD Sold Qty"
)

if uploaded_file is not None:
    progress_bar = st.progress(0, text="準備開始處理文件...")
    try:
        # 文件上傳驗證
        progress_bar.progress(10, text="正在驗證文件格式...")
        engine = 'openpyxl' if uploaded_file.name.lower().endswith('xlsx') else 'xlrd'
        df = pd.read_excel(uploaded_file, engine=engine)
        progress_bar.progress(25, text="文件讀取成功！正在驗證內容...")

        if df.empty:
            st.error("錯誤：上傳的文件為空，請檢查文件內容。")
            st.stop()

        st.success("文件上傳與初步驗證成功！")

        # 4.2. 資料預覽區塊
        with st.expander("基本統計和資料樣本展示", expanded=False):
            st.subheader("資料基本統計")
            st.dataframe(df.describe())
            st.subheader("資料樣本（前100行）")
            st.dataframe(df.head(100))

        # 數據預處理
        progress_bar.progress(40, text="正在進行數據預處理與驗證...")
        processed_df, logs = preprocess_data(df.copy())
        progress_bar.progress(60, text="數據預處理完成！")

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
            st.session_state.cleaned_df = processed_df

            # 4.3. 分析按鈕區塊
            st.header("2. 分析與建議")

            # 預先計算潛在調貨量
            with st.spinner("正在預先計算潛在調貨量..."):
                potential = estimate_transfer_potential(st.session_state.cleaned_df.copy())
            
            st.subheader("潛在調貨量預估")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("A/B模式總需求量", f"{potential['total_needed_A']} 件")
            col2.metric("C模式總需求量", f"{potential['total_needed_C']} 件")
            col3.metric("A模式潛在可轉出", f"{potential['potential_transfer_A']} 件")
            col4.metric("B模式潛在可轉出", f"{potential['potential_transfer_B']} 件")

            transfer_mode = st.radio(
                "請根據預估選擇轉貨力度：",
                ('A: 保守轉貨', 'B: 加強轉貨', 'C: 重點補0'),
                key='transfer_mode',
                help="A模式優先保障安全庫存，B模式則更積極地處理滯銷品，C模式專注於補貨庫存極低的店鋪。"
            )
            
            st.info(f"當前選擇的模式為： **{transfer_mode}**")

            if st.button("🚀 啟動分析生成調貨建議", type="primary"):
                progress_bar.progress(70, text="正在分析數據並生成建議...")
                with st.spinner("演算法運行中，請稍候..."):
                    (
                        recommendations_df, 
                        kpi_metrics, 
                        stats_by_article, 
                        stats_by_om, 
                        transfer_type_dist, 
                        receive_type_dist
                    ) = generate_recommendations(st.session_state.cleaned_df.copy(), transfer_mode)
                    time.sleep(1) # 模擬耗時操作
                progress_bar.progress(90, text="分析完成！正在準備結果展示...")

                if not recommendations_df.empty:
                    st.success("分析完成！")
                    
                    # 4.4. 結果展示區塊
                    st.header("3. 分析結果")
                    
                    st.subheader("統計摘要")
                    kpi_left, _kpi_right = st.columns([4, 8])
                    with kpi_left:
                        render_summary_panel(kpi_metrics)
                    
                    st.markdown("---")

                    # 調貨建議表格
                    st.subheader("調貨建議清單")
                    st.dataframe(recommendations_df)

                    st.markdown("---")

                    # 統計摘要（對齊 Excel 摘要布局）
                    st.subheader("詳細統計摘要")

                    row1_left, row1_right = st.columns([6, 6])
                    with row1_left:
                        st.write("#### 按Article統計")
                        st.dataframe(stats_by_article)
                    with row1_right:
                        st.write("#### 按OM統計")
                        st.dataframe(stats_by_om)

                    row2_left, row2_right = st.columns([6, 6])
                    with row2_left:
                        st.write("#### 轉出類型分析")
                        st.dataframe(transfer_type_dist)
                    with row2_right:
                        st.write("#### 接收類型分析")
                        st.dataframe(receive_type_dist)
                    
                    st.markdown("---")

                    # Display the OM Transfer vs Receive Analysis Chart
                    st.subheader("OM 調貨分析圖表 (OM Transfer vs Receive Analysis Chart)")
                    om_chart_fig = create_om_transfer_chart(recommendations_df, transfer_mode)
                    st.pyplot(om_chart_fig)

                    st.success("Analysis complete! You can now download the recommendations.")

                    excel_data = generate_excel_export(
                        recommendations_df,
                        kpi_metrics,
                        stats_by_article,
                        stats_by_om,
                        transfer_type_dist,
                        receive_type_dist,
                        transfer_mode 
                    )

                    st.download_button(
                        label="📥 下載調貨建議 (Excel)",
                        data=excel_data,
                        file_name=f"調貨建議_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    progress_bar.progress(100, text="處理完畢！")
                else:
                    st.info("根據當前規則，沒有生成任何調貨建議。")
                    progress_bar.progress(100, text="處理完畢！")

    except Exception as e:
        st.error(f"處理文件時發生嚴重錯誤: {e}")
        st.exception(e) # 顯示詳細的錯誤追蹤信息
        if 'progress_bar' in locals():
            progress_bar.progress(100, text="處理失敗！")

# 新增：自動化調撥頁面內容
def show_automated_transfer_page():
    st.header("🤖 自動化調撥中心")
    st.markdown("""
    歡迎來到自動化調撥中心！在這裡，您可以監控系統自動執行的調撥任務、查看詳細的日誌記錄，並追蹤每一次調撥的狀態。
    """)

    # 調撥狀態監控
    st.subheader("即時調撥狀態")
    # 此處可以加入一個表格或儀表板，顯示正在進行的調撥任務
    st.info("目前沒有正在進行的自動化調撥任務。")

    # 調撥記錄與審計
    st.subheader("調撥記錄與審計追蹤")
    # 此處可以加入一個可供篩選和搜尋的歷史記錄表格
    st.text("歷史記錄功能正在開發中...")

# 根據選擇顯示不同頁面
if 'page' not in st.session_state:
    st.session_state.page = "調貨建議"

if st.sidebar.button("調貨建議", key="manual_transfer"):
    st.session_state.page = "調貨建議"
if st.sidebar.button("自動化調撥", key="auto_transfer"):
    st.session_state.page = "自動化調撥"

if st.session_state.page == "調貨建議":
    # 原有的程式碼
    pass
else:
    show_automated_transfer_page()