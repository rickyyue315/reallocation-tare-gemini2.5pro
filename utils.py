import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.ticker import MaxNLocator

def preprocess_data(df):
    logs = []
    required_cols = [
        'Article', 'Article Description', 'RP Type', 'Site', 'OM',
        'SaSa Net Stock', 'Pending Received', 'Safety Stock',
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        error_msg = f"錯誤: Excel文件缺少以下必需欄位: {', '.join(missing_cols)}"
        logs.append(error_msg)
        st.error(error_msg)
        return None, logs

    if 'MOQ' not in df.columns:
        df['MOQ'] = 1
        logs.append("Warning: 缺少 'MOQ' 欄位，已預設為1。")

    df['Notes'] = ''

    df['Article'] = df['Article'].astype(str)
    logs.append("Info: 'Article' 欄位已強制轉換為字串類型。")

    quantity_cols = [
        'MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock',
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    for col in quantity_cols:
        original_nan_mask = pd.to_numeric(df[col], errors='coerce').isnull()
        if original_nan_mask.any():
            df.loc[original_nan_mask, 'Notes'] += f'{col}非數字值已填充為0; '
            df.loc[original_nan_mask, col] = 0
            logs.append(f"Warning: '{col}' 欄位中的非數字值已填充為0。")

        df[col] = df[col].astype(int)

        negative_mask = df[col] < 0
        if negative_mask.any():
            df.loc[negative_mask, 'Notes'] += f'{col}負值已修正為0; '
            df.loc[negative_mask, col] = 0
            logs.append(f"Warning: '{col}' 欄位中的負值已修正為0。")

        if col in ['Last Month Sold Qty', 'MTD Sold Qty']:
            limit = 100000
            over_limit_mask = df[col] > limit
            if over_limit_mask.any():
                df.loc[over_limit_mask, 'Notes'] += f'{col}超過{limit}已限制為{limit}; '
                df.loc[over_limit_mask, col] = limit
                logs.append(f"Warning: '{col}' 中超過 {limit} 的值已限制為 {limit}。")

    string_cols = ['Article Description', 'RP Type', 'Site', 'OM']
    for col in string_cols:
        nan_mask = df[col].isnull() | (df[col] == '')
        if nan_mask.any():
            df.loc[nan_mask, 'Notes'] += f'{col}空值已填充; '
            df.loc[nan_mask, col] = ''
            logs.append(f"Info: '{col}' 欄位中的空值已填充。")

    valid_rp_types = ['ND', 'RF']
    invalid_rp_mask = ~df['RP Type'].isin(valid_rp_types)
    if invalid_rp_mask.any():
        error_msg = f"錯誤: 'RP Type' 欄位包含無效值。只允許 {valid_rp_types}。"
        st.error(error_msg)
        logs.append(error_msg)
        df = df[df['RP Type'].isin(valid_rp_types)]
        logs.append("Warning: 已過濾掉 'RP Type' 無效的行。")

    return df, logs

def _calculate_candidates(df, transfer_mode):
    """
    內部輔助函數，根據業務規則識別轉出和接收候選。
    此函數不執行匹配。
    """
    mode = transfer_mode[0]  # 'A', 'B', or 'C'
    
    senders = []
    receivers = []
    
    grouped = df.groupby('Article')
    
    for article, group in grouped:
        max_sales_in_group = group['Effective Sold Qty'].max()
        
        sorted_group = group
        if mode == 'B':
            sorted_group = group.sort_values(by=['Last Month Sold Qty', 'MTD Sold Qty'], ascending=True)

        for _, row in sorted_group.iterrows():
            stock = row['SaSa Net Stock']
            pending = row['Pending Received']
            safety_stock = row['Safety Stock']
            effective_sales = row['Effective Sold Qty']
            moq = row['MOQ']
            
            # --- 轉出候選邏輯 ---
            if row['RP Type'] == 'ND' and stock > 0:
                senders.append({
                    'type': 'ND轉出', 'priority': 1, 'data': row, 
                    'available_qty': stock, 'current_stock': stock
                })
            
            if mode == 'A' or mode == 'C': # 在C模式下也允許RF過剩轉出
                if row['RP Type'] == 'RF' and (stock + pending) > safety_stock and effective_sales < max_sales_in_group:
                    base_transferable = (stock + pending) - safety_stock
                    upper_limit = (stock + pending) * 0.2
                    actual_transfer = min(base_transferable, max(upper_limit, 2))
                    actual_transfer = min(actual_transfer, stock)
                    
                    if actual_transfer > 0:
                        senders.append({
                            'type': 'RF過剩轉出', 'priority': 2, 'data': row,
                            'available_qty': int(np.floor(actual_transfer)),
                            'current_stock': stock
                        })
            
            elif mode == 'B':
                if row['RP Type'] == 'RF' and (stock + pending) > (moq + 1) and effective_sales < max_sales_in_group:
                    base_transferable = (stock + pending) - (moq + 1)
                    upper_limit = (stock + pending) * 0.5
                    actual_transfer = min(base_transferable, max(upper_limit, 2))
                    actual_transfer = min(actual_transfer, stock)

                    if actual_transfer > 0:
                        senders.append({
                            'type': 'RF加強轉出', 'priority': 2, 'data': row,
                            'available_qty': int(np.floor(actual_transfer)),
                            'current_stock': stock
                        })

            # --- 接收候選邏輯 ---
            if mode == 'C':
                if row['RP Type'] == 'RF' and (stock + pending) <= 1:
                    base_needed = max(row['Safety Stock'] * 0.5, 3)
                    if row['Safety Stock'] == 0:
                        base_needed = max(row['MOQ'], 3)
                    needed = int(base_needed)
                    if needed > 0:
                        receivers.append({
                            'type': 'C模式重點補0', 'priority': 0, 'data': row,
                            'needed_qty': needed, 'effective_sales': effective_sales
                        })
            else:
                if row['RP Type'] == 'RF':
                    if (stock + pending) < safety_stock:
                        needed = safety_stock - (stock + pending)
                        if needed > 0:
                            # 根據庫存狀況和銷售潛力定義接收類型
                            if stock == 0 and effective_sales > 0:
                                receivers.append({
                                    'type': '緊急缺貨補貨', 'priority': 1, 'data': row,
                                    'needed_qty': needed, 'effective_sales': effective_sales
                                })
                            else:
                                receivers.append({
                                    'type': '潛在缺貨補貨', 'priority': 2, 'data': row,
                                    'needed_qty': needed, 'effective_sales': effective_sales
                                })
                    # 針對安全庫存為0但存在缺貨的店鋪，補充起始需求
                    elif (stock + pending) == 0 and safety_stock == 0:
                        needed = int(max(row['MOQ'], 3))
                        receivers.append({
                            'type': '起始補貨需求', 'priority': 1, 'data': row,
                            'needed_qty': needed, 'effective_sales': effective_sales
                        })
                    
    return senders, receivers

def estimate_transfer_potential(df):
    """
    為兩種模式預先計算潛在的可轉出和需求數量，
    以便在運行完整分析前向用戶展示。
    """
    df_copy = df.copy()
    df_copy['Effective Sold Qty'] = np.where(df_copy['Last Month Sold Qty'] > 0, df_copy['Last Month Sold Qty'], df_copy['MTD Sold Qty'])

    senders_A, receivers_A = _calculate_candidates(df_copy, 'A: 保守轉貨')
    senders_B, _ = _calculate_candidates(df_copy, 'B: 加強轉貨')
    _, receivers_C = _calculate_candidates(df_copy, 'C: 重點補0')

    total_needed_A = sum(r['needed_qty'] for r in receivers_A)
    total_needed_C = sum(r['needed_qty'] for r in receivers_C)
    potential_A = sum(s['available_qty'] for s in senders_A)
    potential_B = sum(s['available_qty'] for s in senders_B)

    return {
        "potential_transfer_A": int(potential_A),
        "potential_transfer_B": int(potential_B),
        "total_needed_A": int(total_needed_A),
        "total_needed_C": int(total_needed_C)
    }

def identify_sources(df, transfer_mode):
    out = []
    for _, r in df.iterrows():
        total = int(r['SaSa Net Stock']) + int(r['Pending Received'])
        stock = int(r['SaSa Net Stock'])
        safety = int(r['Safety Stock'])
        rp = str(r['RP Type'])
        if rp == 'ND' and stock > 0:
            out.append({'site': r['Site'], 'om': r['OM'], 'rp_type': rp, 'transferable_qty': int(stock), 'priority': 1, 'original_stock': stock, 'effective_sold_qty': int(r['Effective Sold Qty']), 'source_type': 'ND轉出', 'row': r})
        if rp == 'RF' and stock > 0:
            if transfer_mode.startswith('A'):
                base = max(0, total - safety)
                upper = max(int(total * 0.4), 2)
                qty = min(base, upper, stock)
                if qty > 0 and (stock - qty + int(r['Pending Received'])) >= safety:
                    out.append({'site': r['Site'], 'om': r['OM'], 'rp_type': rp, 'transferable_qty': int(qty), 'priority': 2, 'original_stock': stock, 'effective_sold_qty': int(r['Effective Sold Qty']), 'source_type': 'RF過剩轉出', 'row': r})
            elif transfer_mode.startswith('B'):
                base = max(0, total - safety)
                upper = max(int(total * 0.8), 2)
                qty = min(max(0, upper), stock)
                if qty > 0:
                    remaining_total = stock - qty + int(r['Pending Received'])
                    stype = 'RF過剩轉出' if remaining_total >= safety else 'RF加強轉出'
                    out.append({'site': r['Site'], 'om': r['OM'], 'rp_type': rp, 'transferable_qty': int(qty), 'priority': 2, 'original_stock': stock, 'effective_sold_qty': int(r['Effective Sold Qty']), 'source_type': stype, 'row': r})
            else:
                upper = max(1, int(total * 0.5))
                qty = min(upper, stock)
                if qty > 0:
                    remaining_total = stock - qty + int(r['Pending Received'])
                    stype = 'RF過剩轉出' if remaining_total >= safety else 'RF加強轉出'
                    out.append({'site': r['Site'], 'om': r['OM'], 'rp_type': rp, 'transferable_qty': int(qty), 'priority': 2, 'original_stock': stock, 'effective_sold_qty': int(r['Effective Sold Qty']), 'source_type': stype, 'row': r})
    return out

def identify_destinations(df, transfer_mode):
    out = []
    max_sales = df.groupby('Article')['Effective Sold Qty'].max().to_dict()
    for _, r in df.iterrows():
        if str(r['RP Type']) != 'RF':
            continue
        stock = int(r['SaSa Net Stock'])
        pending = int(r['Pending Received'])
        total = stock + pending
        safety = int(r['Safety Stock'])
        eff = int(r['Effective Sold Qty'])
        need = 0
        dtype = None
        tgt = 0
        recvd = 0
        if transfer_mode.startswith('C') and total <= 1:
            tgt = max(int(safety * 0.5), 3)
            need = max(0, tgt - total)
            if need > 0:
                dtype = 'C模式重點補0'
        else:
            if total < safety:
                need = safety - total
                if stock == 0 and eff > 0:
                    dtype = '緊急缺貨補貨'
                else:
                    if eff >= max_sales.get(r['Article'], eff):
                        dtype = '潛在缺貨補貨'
        if dtype and need > 0:
            pri = 1 if dtype == '緊急缺貨補貨' else 2
            out.append({'site': r['Site'], 'om': r['OM'], 'rp_type': 'RF', 'needed_qty': int(need), 'priority': pri, 'current_stock': stock, 'pending_received': pending, 'safety_stock': safety, 'moq': int(r['MOQ']), 'effective_sold_qty': eff, 'dest_type': dtype, 'target_qty': tgt, 'received_qty': recvd, 'row': r})
    return out

def generate_recommendations(df, transfer_mode):
    recommendations = []
    df['Effective Sold Qty'] = np.where(df['Last Month Sold Qty'] > 0, df['Last Month Sold Qty'], df['MTD Sold Qty'])
    sources = identify_sources(df, transfer_mode)
    destinations = identify_destinations(df, transfer_mode)
    destinations = [d for d in destinations if d['rp_type'] == 'RF']
    def pair_rank(s, d):
        order = {
            ('ND轉出','緊急缺貨補貨'): 1,
            ('ND轉出','潛在缺貨補貨'): 2,
            ('RF過剩轉出','緊急缺貨補貨'): 3,
            ('RF過剩轉出','潛在缺貨補貨'): 4,
            ('RF加強轉出','緊急缺貨補貨'): 5,
            ('RF加強轉出','潛在缺貨補貨'): 6,
            ('RF過剩轉出','C模式重點補0'): 7,
            ('RF加強轉出','C模式重點補0'): 7
        }
        return order.get((s['source_type'], d['dest_type']), 99)
    sources.sort(key=lambda x: (x['priority'], -x['effective_sold_qty'], -x['transferable_qty']))
    destinations.sort(key=lambda x: (x['priority'], x['effective_sold_qty'], -x['current_stock']))
    locked = {}
    for s in sources:
        art = s['row']['Article']
        if art not in locked:
            locked[art] = set()
        cand = [d for d in destinations if d['row']['Article']==art and d['om']==s['om'] and d['site']!=s['site']]
        for d in sorted(cand, key=lambda x: pair_rank(s,x)):
            if s['transferable_qty'] <= 0 or d['needed_qty'] <= 0:
                continue
            if s['site'] in locked[art] or d['site'] in locked[art]:
                continue
            qty = min(int(s['transferable_qty']), int(d['needed_qty']))
            if qty <= 0:
                continue
            s['transferable_qty'] -= qty
            d['needed_qty'] -= qty
            d['received_qty'] += qty
            locked[art].add(s['site'])
            locked[art].add(d['site'])
            sender_type = s['source_type']
            receiver_type = d['dest_type']
            rec = {
                'Article': s['row']['Article'],
                'Product Desc': s['row']['Article Description'],
                'Transfer OM': s['om'],
                'Transfer Site': s['site'],
                'Receive OM': d['om'],
                'Receive Site': d['site'],
                'Transfer Qty': qty,
                'Original Stock': s['original_stock'],
                'After Transfer Stock': s['original_stock'] - qty,
                'Safety Stock': int(s['row']['Safety Stock']),
                'MOQ': int(s['row']['MOQ']),
                'Source Type': sender_type,
                'Destination Type': receiver_type,
                'Cumulative Received Qty': d['received_qty'],
                'Target Qty': d['target_qty'],
                'Remark': f"{sender_type} -> {receiver_type}",
                'Notes': '',
                '_sender_type': sender_type,
                '_receiver_type': receiver_type,
                'OM': s['om']
            }
            recommendations.append(rec)
    if not recommendations:
        return pd.DataFrame(), {}, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    rec_df = pd.DataFrame(recommendations)
    rec_df = rec_df[(rec_df['Transfer Qty'] > 0) & (rec_df['After Transfer Stock'] >= 0)]
    rec_df = rec_df[rec_df['Transfer Site'] != rec_df['Receive Site']]
    kpi_metrics = {
        "總調貨建議行數": len(rec_df),
        "總調貨件數": int(rec_df['Transfer Qty'].sum()),
        "涉及產品數量": int(rec_df['Article'].nunique()),
        "涉及OM數量": int(rec_df['OM'].nunique())
    }
    stats_by_article = rec_df.groupby('Article').agg(
        總調貨件數=('Transfer Qty','sum'),
        調貨行數=('Article','count'),
        涉及OM數量=('OM','nunique')
    ).reset_index().round(2)
    stats_by_om = rec_df.groupby('OM').agg(
        總調貨件數=('Transfer Qty','sum'),
        調貨行數=('OM','count'),
        涉及Article數量=('Article','nunique')
    ).reset_index().round(2)
    transfer_type_dist = rec_df.groupby('_sender_type').agg(總件數=('Transfer Qty','sum'), 建議數量=('_sender_type','count')).reset_index().round(2)
    receive_type_dist = rec_df.groupby('_receiver_type').agg(總件數=('Transfer Qty','sum'), 建議數量=('_receiver_type','count')).reset_index().round(2)
    return rec_df, kpi_metrics, stats_by_article, stats_by_om, transfer_type_dist, receive_type_dist

def create_om_transfer_chart(recommendations_df, transfer_mode):
    if recommendations_df.empty:
        return plt.figure()

    df = recommendations_df.copy()

    nd_transfer = df[df['_sender_type'] == 'ND轉出'].groupby('OM')['Transfer Qty'].sum()
    rf_surplus_transfer = df[df['_sender_type'] == 'RF過剩轉出'].groupby('OM')['Transfer Qty'].sum()

    transfer_data_dict = {
        'ND Transfer Out': nd_transfer,
        'RF Surplus Transfer Out': rf_surplus_transfer
    }

    if transfer_mode.startswith('B'):
        rf_enhanced_transfer = df[df['_sender_type'] == 'RF加強轉出'].groupby('OM')['Transfer Qty'].sum()
        transfer_data_dict['RF Enhanced Transfer Out'] = rf_enhanced_transfer
        
    urgent_receive = df[df['_receiver_type'] == '緊急缺貨補貨'].groupby('OM')['Transfer Qty'].sum()
    potential_receive = df[df['_receiver_type'] == '潛在缺貨補貨'].groupby('OM')['Transfer Qty'].sum()
    c_mode_receive = df[df['_receiver_type'] == 'C模式重點補0'].groupby('OM')['Transfer Qty'].sum()
    initial_receive = df[df['_receiver_type'].isin(['起始補貨需求', 'ND起始補貨'])].groupby('OM')['Transfer Qty'].sum()

    all_oms = df['OM'].unique()
    
    transfer_data = pd.DataFrame(transfer_data_dict).reindex(all_oms).fillna(0)
    
    receive_data_dict = {
        'Urgent Shortage Receive': urgent_receive, 
        'Potential Shortage Receive': potential_receive,
        'Initial Stock Receive': initial_receive
    }
    if transfer_mode.startswith('C'):
        receive_data_dict['C Mode Zero Fill'] = c_mode_receive

    receive_data = pd.DataFrame(receive_data_dict).reindex(all_oms).fillna(0)
    
    chart_data = pd.concat([transfer_data, receive_data], axis=1).fillna(0)

    fig, ax = plt.subplots(figsize=(18, 10))
    
    chart_data.plot(kind='bar', ax=ax, width=0.8)

    for p in ax.patches:
        if p.get_height() > 0:
            ax.annotate(f'{int(p.get_height())}', 
                        (p.get_x() + p.get_width() / 2., p.get_height()), 
                        ha='center', va='center', 
                        xytext=(0, 9), 
                        textcoords='offset points',
                        fontsize=9)

    ax.set_title('OM Transfer vs Receive Analysis', fontsize=18, weight='bold')
    ax.set_xlabel('OM Unit', fontsize=14)
    ax.set_ylabel('Transfer Quantity', fontsize=14)
    ax.tick_params(axis='x', rotation=45, labelsize=12)
    ax.tick_params(axis='y', labelsize=12)
    
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    
    ax.legend(title='Transfer/Receive Type', fontsize=12)
    ax.grid(axis='y', linestyle='--', alpha=0.7)

    plt.tight_layout()
    return fig

def generate_excel_export(rec_df, kpis, stats_article, stats_om, transfer_dist, receive_dist, transfer_mode):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        column_order = [
            'Article', 'Product Desc', 'Transfer OM', 'Transfer Site', 'Receive OM', 'Receive Site',
            'Transfer Qty', 'Original Stock', 'After Transfer Stock', 'Safety Stock', 'MOQ', 'Remark', 'Notes'
        ]

        export_rec_df = rec_df.copy()
        for col in column_order:
            if col not in export_rec_df.columns:
                export_rec_df[col] = ''

        export_rec_df = export_rec_df[column_order]
        export_rec_df.to_excel(writer, sheet_name='調貨建議', index=False)
        ws = writer.sheets['調貨建議']
        ws.set_column(0, 0, 15)
        ws.set_column(1, 1, 30)
        ws.set_column(2, 2, 15)
        ws.set_column(3, 3, 15)
        ws.set_column(4, 4, 15)
        ws.set_column(5, 5, 15)
        ws.set_column(6, 6, 12)
        ws.set_column(7, 7, 15)
        ws.set_column(8, 8, 18)
        ws.set_column(9, 9, 12)
        ws.set_column(10, 10, 8)
        ws.set_column(11, 11, 35)
        ws.set_column(12, 12, 60)

        summary_sheet_name = '統計摘要'
        workbook = writer.book
        worksheet = workbook.add_worksheet(summary_sheet_name)
        writer.sheets[summary_sheet_name] = worksheet

        title_format = workbook.add_format({'bold': True, 'font_size': 14})
        label_fmt = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
        value_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E2EFDA'})
        table_title_fmt = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})

        worksheet.write(0, 0, '調貨建議統計摘要', title_format)

        k_labels = ['總調貨建議行數', '總調貨件數', '涉及產品數量', '涉及OM數量']
        k_values = [kpis.get('總調貨建議行數', 0), kpis.get('總調貨件數', 0), kpis.get('涉及產品數量', 0), kpis.get('涉及OM數量', 0)]
        for i, (lbl, val) in enumerate(zip(k_labels, k_values)):
            worksheet.write(2 + i, 0, lbl, label_fmt)
            worksheet.write(2 + i, 1, val, value_fmt)

        worksheet.set_column(0, 0, 20)
        worksheet.set_column(1, 1, 12)
        worksheet.set_column(5, 9, 18)

        sa_start_row, sa_start_col = 8, 0
        worksheet.write(sa_start_row, sa_start_col, '按Article統計', table_title_fmt)
        stats_article.to_excel(writer, sheet_name=summary_sheet_name, startrow=sa_start_row + 2, startcol=sa_start_col, index=False)

        so_start_row, so_start_col = 8, 5
        worksheet.write(so_start_row, so_start_col, '按OM統計', table_title_fmt)
        stats_om.to_excel(writer, sheet_name=summary_sheet_name, startrow=so_start_row + 2, startcol=so_start_col, index=False)

        transfer_dist = transfer_dist.rename(columns={'涉及行數': '建議數量'})
        receive_dist = receive_dist.rename(columns={'涉及行數': '建議數量'})

        td_start_row, td_start_col = 20, 0
        worksheet.write(td_start_row, td_start_col, '轉出類型分析', table_title_fmt)
        transfer_dist.to_excel(writer, sheet_name=summary_sheet_name, startrow=td_start_row + 2, startcol=td_start_col, index=False)

        rd_start_row, rd_start_col = 20, 5
        worksheet.write(rd_start_row, rd_start_col, '接收類型分析', table_title_fmt)
        receive_dist.to_excel(writer, sheet_name=summary_sheet_name, startrow=rd_start_row + 2, startcol=rd_start_col, index=False)

    return output.getvalue()
