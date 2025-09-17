import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

# 假設 utils.py 在同一目錄下
from utils import preprocess_data, generate_recommendations, create_om_transfer_chart

def generate_test_data(num_rows=500):
    """
    生成用於調貨建議系統的模擬Excel測試數據。
    - 包含各種RP Type組合
    - 涵蓋邊界情況（零庫存、高銷量、無銷量等）
    """
    data = {
        'Article': [f'ART{1000 + i}' for i in range(num_rows)],
        'Article Description': [f'Product Description {i}' for i in range(num_rows)],
        'RP Type': np.random.choice(['ND', 'RF'], size=num_rows, p=[0.2, 0.8]),
        'Site': [f'Site{100 + (i % 20)}' for i in range(num_rows)],
        'OM': [f'OM{1 + (i % 5)}' for i in range(num_rows)],
        'MOQ': np.random.choice([1, 2, 6, 12], size=num_rows),
        'SaSa Net Stock': np.random.randint(0, 100, size=num_rows),
        'Pending Received': np.random.randint(0, 50, size=num_rows),
        'Safety Stock': np.random.randint(5, 20, size=num_rows),
        'Last Month Sold Qty': np.random.randint(0, 40, size=num_rows),
        'MTD Sold Qty': np.random.randint(0, 30, size=num_rows),
    }
    df = pd.DataFrame(data)

    # --- 創建邊界情況 ---
    # 1. 緊急缺貨 (庫存=0, 有銷量)
    urgent_shortage_indices = df.sample(frac=0.1).index
    df.loc[urgent_shortage_indices, 'SaSa Net Stock'] = 0
    df.loc[urgent_shortage_indices, 'RP Type'] = 'RF'
    df.loc[urgent_shortage_indices, 'Last Month Sold Qty'] = np.random.randint(5, 15, size=len(urgent_shortage_indices))

    # 2. 潛在缺貨 (庫存+在途 < 安全庫存, 銷量為組內最高)
    potential_shortage_indices = df.sample(frac=0.1).index
    df.loc[potential_shortage_indices, 'SaSa Net Stock'] = 1
    df.loc[potential_shortage_indices, 'Pending Received'] = 1
    df.loc[potential_shortage_indices, 'Safety Stock'] = 10
    df.loc[potential_shortage_indices, 'RP Type'] = 'RF'

    # 3. ND 轉出 (有庫存)
    nd_transfer_indices = df[df['RP Type'] == 'ND'].sample(frac=0.5).index
    df.loc[nd_transfer_indices, 'SaSa Net Stock'] = np.random.randint(10, 30, size=len(nd_transfer_indices))

    # 4. RF 過剩/加強轉出 (庫存+在途 > 安全庫存)
    rf_surplus_indices = df[df['RP Type'] == 'RF'].sample(frac=0.3).index
    df.loc[rf_surplus_indices, 'SaSa Net Stock'] = np.random.randint(50, 80, size=len(rf_surplus_indices))
    df.loc[rf_surplus_indices, 'Pending Received'] = np.random.randint(20, 40, size=len(rf_surplus_indices))
    df.loc[rf_surplus_indices, 'Safety Stock'] = 5

    # 5. 銷量為0的產品
    zero_sales_indices = df.sample(frac=0.1).index
    df.loc[zero_sales_indices, 'Last Month Sold Qty'] = 0
    df.loc[zero_sales_indices, 'MTD Sold Qty'] = 0
    
    # 6. 確保潛在缺貨候選的銷量是最高的
    for article, group in df.groupby('Article'):
        max_sales = group['Last Month Sold Qty'].max()
        potential_indices = group[group.index.isin(potential_shortage_indices)].index
        if not potential_indices.empty:
            df.loc[potential_indices, 'Last Month Sold Qty'] = max_sales + 5

    return df

def run_tests(df):
    """
    運行所有測試用例：數據處理、兩種模式下的建議生成和圖表創建。
    """
    print("--- 開始測試數據預處理 ---")
    processed_df, logs = preprocess_data(df.copy())
    assert processed_df is not None, "數據預處理失敗"
    print("✅ 數據預處理成功。")

    # 測試兩種模式
    for mode in ["A: 保守轉貨", "B: 加強轉貨"]:
        print(f"\n--- 開始測試模式: {mode} ---")
        
        # 1. 測試建議生成
        recs, kpis, _, _, _, _ = generate_recommendations(processed_df.copy(), mode)
        print(f"✅ 建議生成成功。生成了 {len(recs)} 條建議。")
        if not recs.empty:
            print(f"    - 總調貨件數: {kpis.get('總調貨件數', 0)}")
            print(f"    - 涉及產品數: {kpis.get('涉及產品數量', 0)}")

        # 2. 測試圖表生成
        fig = create_om_transfer_chart(recs, mode)
        assert fig is not None, f"圖表生成失敗於模式 {mode}"
        
        # 保存圖表到文件
        chart_filename = f'test_chart_{mode.split(": ")[0]}.png'
        fig.savefig(chart_filename)
        print(f"✅ 圖表生成成功並保存為: {chart_filename}")
        plt.close(fig) # 關閉圖形以釋放內存

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_filename = "test_data_v1.7.xlsx"
    excel_path = os.path.join(script_dir, excel_filename)

    # 步驟 1: 生成測試數據 (如果文件不存在)
    if not os.path.exists(excel_path):
        print(f"未找到測試文件 {excel_filename}，正在生成...")
        test_df = generate_test_data(num_rows=500)
        test_df.to_excel(excel_path, index=False)
        print(f"✅ 測試數據已生成並保存到: {excel_path}")
    else:
        print(f"找到現有測試文件: {excel_path}")
        test_df = pd.read_excel(excel_path)

    # 步驟 2: 運行測試
    run_tests(test_df)
    
    print("\n🎉 所有測試用例執行完畢！")