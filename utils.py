import pandas as pd
import numpy as np
import streamlit as st

def preprocess_data(df):
    """
    對輸入的 DataFrame 進行數據預處理與驗證。
    """
    logs = []

    # 將 'Article Description' 欄位重命名為 'Product Desc'
    if 'Article Description' in df.columns:
        df.rename(columns={'Article Description': 'Product Desc'}, inplace=True)

    # 確保 Article 欄位為 12 位字符串
    if 'Article' in df.columns:
        df['Article'] = df['Article'].astype(str).str.zfill(12)
    else:
        logs.append("錯誤: 缺少 'Article' 欄位。")
        st.error("錯誤: 缺少 'Article' 欄位。")
        return None, logs

    # 定義關鍵欄位
    integer_cols = ['SaSa Net Stock', 'Pending Received', 'Safety Stock', 'Last Month Sold Qty', 'MTD Sold Qty']
    string_cols = ['OM', 'RP Type', 'Site']

    # 類型轉換與異常處理
    for col in integer_cols:
        if col in df.columns:
            original_type = df[col].dtype
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
            nan_mask = df[col].isnull()
            if nan_mask.any():
                logs.append(f"'{col}' 欄位中的非數字值已設為 0。")
            
            df[col] = df[col].fillna(0).astype(int)

            if col in ['Last Month Sold Qty', 'MTD Sold Qty']:
                # 銷量異常值校正
                large_sales_mask = df[col] > 100000
                if large_sales_mask.any():
                    logs.append(f"'{col}' 中大於 100,000 的值已校正為 100,000。")
                    df.loc[large_sales_mask, col] = 100000
                
                negative_sales_mask = df[col] < 0
                if negative_sales_mask.any():
                    logs.append(f"'{col}' 中小於 0 的值已校正為 0。")
                    df.loc[negative_sales_mask, col] = 0

            else: # SaSa Net Stock, Pending Received, Safety Stock
                negative_stock_mask = df[col] < 0
                if negative_stock_mask.any():
                    logs.append(f"'{col}' 中小於 0 的值已校正為 0。")
                    df.loc[negative_stock_mask, col] = 0
        else:
            df[col] = 0
            logs.append(f"警告: 缺少 '{col}' 欄位，已自動填充為 0。")

    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
        else:
            df[col] = ""
            logs.append(f"警告: 缺少 '{col}' 欄位，已自動填充為空字符串。")
            
    # 可選欄位
    if 'Product Desc' not in df.columns:
        df['Product Desc'] = ""

    return df, logs

def generate_recommendations(df):
    """
    根據處理後的數據生成調貨建議。
    """
    recommendations = []
    df['Effective Sold Qty'] = df[['Last Month Sold Qty', 'MTD Sold Qty']].max(axis=1)

    grouped = df.groupby(['Article', 'OM'])

    for (article, om), group in grouped:
        # 1. 識別最高銷量店鋪
        max_sales_site = None
        if not group.empty:
            max_sales_site = group.loc[group['Effective Sold Qty'].idxmax()]['Site']

        # 2. 識別轉出和接收店鋪
        senders = []
        receivers = []

        for _, row in group.iterrows():
            # 轉出規則
            if row['RP Type'] == 'ND':
                senders.append({
                    'type': 'ND',
                    'priority': 1,
                    'data': row,
                    'available_qty': row['SaSa Net Stock']
                })
            elif row['RP Type'] == 'RF' and (row['SaSa Net Stock'] + row['Pending Received'] > row['Safety Stock']) and (row['Site'] != max_sales_site):
                available_stock = row['SaSa Net Stock'] + row['Pending Received'] - row['Safety Stock']
                upper_limit = np.floor((row['SaSa Net Stock'] + row['Pending Received']) * 0.2)
                
                # 轉出數量最少為 2
                qty = min(available_stock, upper_limit)
                if qty < 2:
                    if row['SaSa Net Stock'] >= 2 and (row['SaSa Net Stock'] - 2 >= row['Safety Stock']):
                        qty = 2
                    else:
                        qty = 0 # 不滿足最少轉出2件的條件
                
                final_qty = min(qty, row['SaSa Net Stock'])
                if final_qty > 0:
                    senders.append({
                        'type': 'RF',
                        'priority': 2,
                        'data': row,
                        'available_qty': final_qty
                    })

            # 接收規則
            if row['RP Type'] == 'RF':
                if row['SaSa Net Stock'] == 0 and row['Effective Sold Qty'] > 0:
                    receivers.append({
                        'type': 'Urgent',
                        'priority': 1,
                        'data': row,
                        'needed_qty': row['Safety Stock']
                    })
                elif (row['SaSa Net Stock'] + row['Pending Received'] < row['Safety Stock']) and (row['Site'] == max_sales_site):
                    needed = row['Safety Stock'] - (row['SaSa Net Stock'] + row['Pending Received'])
                    if needed > 0:
                        receivers.append({
                            'type': 'Potential Shortage',
                            'priority': 2,
                            'data': row,
                            'needed_qty': needed
                        })

        # 3. 排序和匹配
        senders.sort(key=lambda x: x['priority'])
        receivers.sort(key=lambda x: x['priority'])

        for sender in senders:
            for receiver in receivers:
                if sender['available_qty'] > 0 and receiver['needed_qty'] > 0 and sender['data']['Site'] != receiver['data']['Site']:
                    transfer_qty = min(sender['available_qty'], receiver['needed_qty'])
                    note = ""

                    # 調貨數量為 1 時的調整規則
                    if transfer_qty == 1:
                        if sender['data']['SaSa Net Stock'] >= 2 and (sender['data']['SaSa Net Stock'] - 2 >= sender['data']['Safety Stock']):
                            transfer_qty = 2
                            note = "調貨數量從1調整為2"
                        else:
                            transfer_qty = 0 # 取消調貨

                    if transfer_qty > 0:
                        # 確保最終轉出數量不超過發送方當前實際庫存
                        transfer_qty = min(transfer_qty, sender['data']['SaSa Net Stock'])
                        # 確保轉出後庫存不低於安全庫存
                        if sender['data']['SaSa Net Stock'] - transfer_qty >= sender['data']['Safety Stock']:
                            final_note = f"{sender['type']}轉出"
                            if note:
                                final_note += f"; {note}"

                            recommendations.append({
                                'Article': article,
                                'Product Desc': sender['data']['Product Desc'],
                                'OM': om,
                                'Transfer Site': sender['data']['Site'],
                                'Receive Site': receiver['data']['Site'],
                                'Transfer Qty': transfer_qty,
                                'Transfer Site Original Stock': sender['data']['SaSa Net Stock'],
                                'Transfer Site After Stock': sender['data']['SaSa Net Stock'] - transfer_qty,
                                'Transfer Site Safety Stock': sender['data']['Safety Stock'],
                                'Notes': final_note,
                                'Transfer Type': sender['type'],
                                'Receive Priority': receiver['priority']
                            })
                            sender['available_qty'] -= transfer_qty
                            receiver['needed_qty'] -= transfer_qty
                            # 更新原始 DataFrame 中的庫存
                            df.loc[sender['data'].name, 'SaSa Net Stock'] -= transfer_qty
                            df.loc[receiver['data'].name, 'SaSa Net Stock'] += transfer_qty

    if not recommendations:
        return pd.DataFrame(), {}, {}

    # 4. 格式化輸出
    rec_df = pd.DataFrame(recommendations)
    
    # 統計摘要
    total_recommendations = len(rec_df)
    total_transfer_qty = rec_df['Transfer Qty'].sum()

    summary_kpis = {
        "總調貨涉及行數": total_recommendations,
        "總調貨件數": total_transfer_qty
    }

    summary_details = {
        'by_article': rec_df.groupby('Article').agg(總調貨件數=('Transfer Qty', 'sum'), 涉及的OM數量=('OM', 'nunique')).reset_index(),
        'by_om': rec_df.groupby('OM').agg(總調貨件數=('Transfer Qty', 'sum'), 涉及的Article數量=('Article', 'nunique')).reset_index(),
        'by_transfer_type': rec_df.groupby('Transfer Type').agg(涉及行數=('Article', 'count'), 總件數=('Transfer Qty', 'sum')).reset_index(),
        'by_receive_priority': rec_df.groupby('Receive Priority').agg(涉及行數=('Article', 'count'), 總件數=('Transfer Qty', 'sum')).reset_index()
    }

    # 移除臨時列
    rec_df = rec_df.drop(columns=['Transfer Type', 'Receive Priority'])

    return rec_df, summary_kpis, summary_details