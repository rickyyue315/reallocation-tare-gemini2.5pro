import pandas as pd
import numpy as np
import os

def generate_test_data(filename="test_data_v1.6.xlsx"):
    """
    生成用於測試的模擬Excel數據，涵蓋各種業務場景。
    """
    data = {
        'Article': ['A001'] * 5 + ['A002'] * 5 + ['A003'] * 4,
        'Article Description': ['Product A'] * 5 + ['Product B'] * 5 + ['Product C'] * 4,
        'RP Type': ['ND', 'RF', 'RF', 'RF', 'RF'] + ['RF', 'RF', 'RF', 'ND', 'RF'] + ['RF', 'RF', 'RF', 'RF'],
        'Site': ['S01', 'S02', 'S03', 'S04', 'S05'] + ['S01', 'S02', 'S03', 'S04', 'S05'] + ['S06', 'S07', 'S08', 'S09'],
        'OM': ['OM1'] * 5 + ['OM1'] * 5 + ['OM2'] * 4,
        'SaSa Net Stock': [10, 0, 5, 20, 100] + [50, 80, 0, 5, 10] + [0, 30, 15, 5],
        'Pending Received': [0, 5, 2, 0, 10] + [10, 0, 5, 0, 0] + [10, 5, 0, 0],
        'Safety Stock': [5, 10, 8, 15, 20] + [20, 25, 15, 10, 12] + [5, 10, 8, 6],
        'Last Month Sold Qty': [0, 20, 5, 10, 100] + [30, 0, 10, 0, 5] + [8, 12, 0, 3],
        'MTD Sold Qty': [2, 15, 3, 8, 80] + [25, 5, 8, 1, 2] + [6, 10, 5, 2]
    }
    df = pd.DataFrame(data)

    # 引入一些異常數據
    df.loc[1, 'SaSa Net Stock'] = -5 # 負庫存
    df.loc[3, 'Last Month Sold Qty'] = 120000 # 銷量異常值
    df.loc[7, 'RP Type'] = 'INVALID' # 無效RP Type
    df.loc[9, 'OM'] = np.nan # OM為空值

    # 保存到Excel
    df.to_excel(filename, index=False)
    print(f"測試數據已生成: {filename}")

def test_chart_generation():
    """
    測試圖表生成功能。
    需要從主應用導入相關函數，這裡僅為結構示例。
    實際測試中，我們會加載數據，調用 `generate_recommendations` 和 `create_om_transfer_chart`。
    """
    print("\n--- 開始測試圖表生成功能 ---")
    # 假設我們已經有了 recommendations_df
    # from utils import preprocess_data, generate_recommendations, create_om_transfer_chart
    
    # 1. 載入並預處理數據
    # if os.path.exists("test_data_v1.6.xlsx"):
    #     df = pd.read_excel("test_data_v1.6.xlsx")
    #     processed_df, _ = preprocess_data(df)
    #     if processed_df is not None:
    #         rec_df, _, _, _, _, _ = generate_recommendations(processed_df)
            
    #         # 2. 生成圖表
    #         if not rec_df.empty:
    #             fig = create_om_transfer_chart(rec_df)
    #             if fig:
    #                 print("圖表生成成功！")
    #                 # 可以選擇保存圖表以供驗證
    #                 # fig.savefig("test_chart.png")
    #                 # print("圖表已保存為 test_chart.png")
    #             else:
    #                 print("圖表生成失敗！")
    #         else:
    #             print("沒有生成調貨建議，無法測試圖表。")
    # else:
    #     print("找不到測試數據文件。")
    print("（此處為結構示例，完整測試需在Streamlit環境下運行）")
    print("--- 測試結束 ---")

if __name__ == "__main__":
    generate_test_data()
    test_chart_generation()