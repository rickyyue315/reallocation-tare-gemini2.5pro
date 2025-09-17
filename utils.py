import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns

# 設置 matplotlib 支持中文
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False

def preprocess_data(df):
    """
    對輸入的 DataFrame 進行數據預處理、驗證和清理。
    - 強制轉換類型
    - 填充缺失值
    - 校正無效數據（如負庫存）
    - 記錄所有清理操作
    """
    logs = []
    required_cols = [
        'Article', 'Article Description', 'RP Type', 'Site', 'OM', 
        'SaSa Net Stock', 'Pending Received', 'Safety Stock', 
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    # 檢查必需欄位
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        error_msg = f"錯誤: Excel文件缺少以下必需欄位: {', '.join(missing_cols)}"
        logs.append(error_msg)
        st.error(error_msg)
        return None, logs

    # 添加Notes欄位
    df['Notes'] = ''

    # 1. Article欄位強制轉換為字串
    df['Article'] = df['Article'].astype(str)
    logs.append("Info: 'Article' 欄位已強制轉換為字串類型。")

    # 2. 數量欄位處理
    quantity_cols = [
        'SaSa Net Stock', 'Pending Received', 'Safety Stock', 
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    for col in quantity_cols:
        # 填充非數字值為0
        original_nan_mask = pd.to_numeric(df[col], errors='coerce').isnull()
        if original_nan_mask.any():
            df.loc[original_nan_mask, 'Notes'] += f'{col}非數字值已填充為0; '
            df.loc[original_nan_mask, col] = 0
            logs.append(f"Warning: '{col}' 欄位中的非數字值已填充為0。")
        
        df[col] = df[col].astype(int)

        # 3. 負值修正為0
        negative_mask = df[col] < 0
        if negative_mask.any():
            df.loc[negative_mask, 'Notes'] += f'{col}負值已修正為0; '
            df.loc[negative_mask, col] = 0
            logs.append(f"Warning: '{col}' 欄位中的負值已修正為0。")

        # 4. 銷量異常值處理
        if col in ['Last Month Sold Qty', 'MTD Sold Qty']:
            limit = 100000
            over_limit_mask = df[col] > limit
            if over_limit_mask.any():
                df.loc[over_limit_mask, 'Notes'] += f'{col}超過{limit}已限制為{limit}; '
                df.loc[over_limit_mask, col] = limit
                logs.append(f"Warning: '{col}' 中超過 {limit} 的值已限制為 {limit}。")

    # 5. 字串欄位空值填充
    string_cols = ['Article Description', 'RP Type', 'Site', 'OM']
    for col in string_cols:
        nan_mask = df[col].isnull() | (df[col] == '')
        if nan_mask.any():
            df.loc[nan_mask, 'Notes'] += f'{col}空值已填充; '
            df.loc[nan_mask, col] = ''
            logs.append(f"Info: '{col}' 欄位中的空值已填充。")

    # 業務邏輯驗證
    valid_rp_types = ['ND', 'RF']
    invalid_rp_mask = ~df['RP Type'].isin(valid_rp_types)
    if invalid_rp_mask.any():
        error_msg = f"錯誤: 'RP Type' 欄位包含無效值。只允許 {valid_rp_types}。"
        st.error(error_msg)
        logs.append(error_msg)
        # 可以在這裡決定是中止還是過濾掉這些行
        df = df[df['RP Type'].isin(valid_rp_types)]
        logs.append("Warning: 已過濾掉 'RP Type' 無效的行。")

    return df, logs

def generate_recommendations(df):
    """
    實現調貨建議的核心演算法。
    """
    recommendations = []
    
    # 計算有效銷量
    df['Effective Sold Qty'] = np.where(df['Last Month Sold Qty'] > 0, df['Last Month Sold Qty'], df['MTD Sold Qty'])

    # 按產品和OM分組處理
    grouped = df.groupby(['Article', 'OM'])

    for (article, om), group in grouped:
        # 找到該產品組中的最高銷量
        max_sales_in_group = group['Effective Sold Qty'].max()

        senders = []
        receivers = []

        # 識別轉出和接收候選
        for _, row in group.iterrows():
            stock = row['SaSa Net Stock']
            pending = row['Pending Received']
            safety_stock = row['Safety Stock']
            effective_sales = row['Effective Sold Qty']
            
            # --- 轉出候選識別規則 ---
            # 優先順序1 - ND類型完全轉出
            if row['RP Type'] == 'ND' and stock > 0:
                senders.append({
                    'type': 'ND轉出', 'priority': 1, 'data': row, 
                    'available_qty': stock
                })
            # 優先順序2 - RF類型過剩轉出
            elif row['RP Type'] == 'RF' and (stock + pending) > safety_stock and effective_sales < max_sales_in_group:
                base_transferable = (stock + pending) - safety_stock
                upper_limit = (stock + pending) * 0.2
                
                # 上限控制，但最少2件
                actual_transfer = min(base_transferable, max(upper_limit, 2))
                
                # 不能超過實際庫存
                actual_transfer = min(actual_transfer, stock)
                
                if actual_transfer > 0:
                    senders.append({
                        'type': 'RF過剩轉出', 'priority': 2, 'data': row,
                        'available_qty': int(np.floor(actual_transfer))
                    })

            # --- 接收候選識別規則 ---
            # 優先順序1 - 緊急缺貨補貨
            if row['RP Type'] == 'RF' and stock == 0 and effective_sales > 0:
                receivers.append({
                    'type': '緊急缺貨補貨', 'priority': 1, 'data': row,
                    'needed_qty': safety_stock
                })
            # 優先順序2 - 潛在缺貨補貨
            elif row['RP Type'] == 'RF' and (stock + pending) < safety_stock and effective_sales == max_sales_in_group and max_sales_in_group > 0:
                needed = safety_stock - (stock + pending)
                if needed > 0:
                    receivers.append({
                        'type': '潛在缺貨補貨', 'priority': 2, 'data': row,
                        'needed_qty': needed
                    })

        # 排序和匹配
        senders.sort(key=lambda x: x['priority'])
        receivers.sort(key=lambda x: x['priority'])

        for sender in senders:
            for receiver in receivers:
                if sender['available_qty'] > 0 and receiver['needed_qty'] > 0 and sender['data']['Site'] != receiver['data']['Site']:
                    transfer_qty = min(sender['available_qty'], receiver['needed_qty'])
                    
                    # 調貨數量優化：如果只有1件，嘗試調高到2件
                    if transfer_qty == 1:
                        # 檢查轉出2件後是否會低於安全庫存
                        if sender['data']['SaSa Net Stock'] >= 2 and (sender['data']['SaSa Net Stock'] - 2) >= sender['data']['Safety Stock']:
                            transfer_qty = 2
                    
                    if transfer_qty > 0:
                        # 確保最終轉出數量不超過發送方當前實際庫存
                        final_transfer_qty = min(transfer_qty, sender['data']['SaSa Net Stock'])
                        
                        # 再次檢查轉出後庫存
                        if sender['data']['SaSa Net Stock'] - final_transfer_qty >= 0:
                            recommendations.append({
                                'Article': article,
                                'Product Desc': sender['data']['Article Description'],
                                'OM': om,
                                'Transfer Site': sender['data']['Site'],
                                'Receive Site': receiver['data']['Site'],
                                'Transfer Qty': final_transfer_qty,
                                'Original Stock': sender['data']['SaSa Net Stock'],
                                'After Transfer Stock': sender['data']['SaSa Net Stock'] - final_transfer_qty,
                                'Safety Stock': sender['data']['Safety Stock'],
                                'Notes': f"{sender['type']} -> {receiver['type']}",
                                '_sender_type': sender['type'],
                                '_receiver_type': receiver['type']
                            })
                            sender['available_qty'] -= final_transfer_qty
                            receiver['needed_qty'] -= final_transfer_qty
                            # 更新庫存以反映調貨，避免重複計算
                            sender['data']['SaSa Net Stock'] -= final_transfer_qty

    if not recommendations:
        return pd.DataFrame(), {}, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    rec_df = pd.DataFrame(recommendations)

    # --- 統計分析功能 ---
    # 1. 基本KPI
    kpi_metrics = {
        "總調貨建議數量": len(rec_df),
        "總調貨件數": int(rec_df['Transfer Qty'].sum()),
        "涉及產品數量": int(rec_df['Article'].nunique()),
        "涉及OM數量": int(rec_df['OM'].nunique())
    }

    # 2. 按產品統計
    stats_by_article = rec_df.groupby('Article').agg(
        總調貨件數=('Transfer Qty', 'sum'),
        調貨行數=('Article', 'count'),
        涉及OM數量=('OM', 'nunique')
    ).reset_index().round(2)

    # 3. 按OM統計
    stats_by_om = rec_df.groupby('OM').agg(
        總調貨件數=('Transfer Qty', 'sum'),
        調貨行數=('OM', 'count'),
        涉及產品數量=('Article', 'nunique')
    ).reset_index().round(2)

    # 4. 轉出類型分佈
    transfer_type_dist = rec_df.groupby('_sender_type').agg(
        總件數=('Transfer Qty', 'sum'),
        涉及行數=('_sender_type', 'count')
    ).reset_index().round(2)

    # 5. 接收類型分佈
    receive_type_dist = rec_df.groupby('_receiver_type').agg(
        總件數=('Transfer Qty', 'sum'),
        涉及行數=('_receiver_type', 'count')
    ).reset_index().round(2)
    
    # 移除臨時列
    rec_df_final = rec_df.drop(columns=['_sender_type', '_receiver_type'])

    return rec_df_final, kpi_metrics, stats_by_article, stats_by_om, transfer_type_dist, receive_type_dist

def create_om_transfer_chart(recommendations_df):
    """
    創建 matplotlib 橫條圖進行數據視覺化。
    """
    if recommendations_df.empty:
        return plt.figure()

    # 按OM統計轉出和接收數量
    transfer_out = recommendations_df.groupby('OM')['Transfer Qty'].sum()
    
    # 接收數量與轉出數量相同，只是視角不同
    transfer_in = recommendations_df.groupby('OM')['Transfer Qty'].sum()

    chart_data = pd.DataFrame({'轉出數量': transfer_out, '接收數量': transfer_in}).fillna(0)
    
    fig, ax = plt.subplots(figsize=(12, 8))
    chart_data.plot(kind='bar', ax=ax)
    
    ax.set_title('OM Transfer vs Receive Analysis', fontsize=16)
    ax.set_xlabel('OM 單位', fontsize=12)
    ax.set_ylabel('調貨數量', fontsize=12)
    ax.tick_params(axis='x', rotation=45)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    return fig

def generate_excel_export(rec_df, kpis, stats_article, stats_om, transfer_dist, receive_dist):
    """
    實現Excel文件匯出功能。
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 工作表1 - 調貨建議
        # 確保欄位順序和英文名
        export_rec_df = rec_df[[
            'Article', 'Product Desc', 'OM', 'Transfer Site', 'Receive Site', 
            'Transfer Qty', 'Original Stock', 'After Transfer Stock', 
            'Safety Stock', 'Notes'
        ]]
        export_rec_df.to_excel(writer, sheet_name='調貨建議', index=False)

        # 工作表2 - 統計摘要
        summary_sheet = writer.sheets['統計摘要']
        start_row = 0

        # Helper function to write a dataframe with title
        def write_df_with_title(df, title, row):
            pd.DataFrame([title]).to_excel(writer, sheet_name='統計摘要', startrow=row, index=False, header=False)
            df.to_excel(writer, sheet_name='統計摘要', startrow=row + 2, index=False)
            return row + len(df) + 5 # 2 for title, df_len, 3 for spacing

        # KPI概覽
        kpi_df = pd.DataFrame([kpis])
        start_row = write_df_with_title(kpi_df, "KPI 概覽", start_row)
        
        # 按Article統計
        start_row = write_df_with_title(stats_article, "按 Article 統計", start_row)
        
        # 按OM統計
        start_row = write_df_with_title(stats_om, "按 OM 統計", start_row)
        
        # 轉出類型分佈
        start_row = write_df_with_title(transfer_dist, "轉出類型分佈", start_row)
        
        # 接收類型分佈
        write_df_with_title(receive_dist, "接收類型分佈", start_row)

    return output.getvalue()