import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import io
import logging
from typing import List, Dict, Tuple, Optional, Any, Union
import re

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="📦 調貨建議生成系統",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Sidebar design
def render_sidebar():
    st.sidebar.header("系統資訊")
    st.sidebar.info("""
    **版本：v1.7**
    **開發者：Ricky**

    **核心功能：**  
    - ✅ ND/RF類型智能識別
    - ✅ 優先級調貨匹配
    - ✅ RF過剩轉出限制
    - ✅ 統計分析和圖表
    - ✅ Excel格式匯出
    """)

# Data validation and preprocessing
def validate_and_preprocess_data(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Validate and preprocess input data"""
    notes: List[str] = []
    
    # Required columns check
    required_columns = [
        'Article', 'Article Description', 'RP Type', 'Site', 'OM',
        'MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock',
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"缺少必需欄位: {', '.join(missing_columns)}")
    
    # Article field processing
    df['Article'] = df['Article'].astype(str).str.strip()
    notes.append("Article欄位已轉換為字串類型")
    
    # Numeric fields processing
    numeric_columns = [
        'MOQ', 'SaSa Net Stock', 'Pending Received', 'Safety Stock',
        'Last Month Sold Qty', 'MTD Sold Qty'
    ]
    
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        df[col] = df[col].fillna(0)
        
        # Correct negative values
        negative_mask = df[col] < 0
        if negative_mask.any():
            df.loc[negative_mask, col] = 0
            notes.append(f"{col}欄位負值已修正為0")
        
        # Limit extreme values for sales
        if col in ['Last Month Sold Qty', 'MTD Sold Qty']:
            extreme_mask = df[col] > 100000
            if extreme_mask.any():
                df.loc[extreme_mask, col] = 100000
                notes.append(f"{col}欄位異常值(>100000)已限制為100000")
    
    # Text fields processing
    text_columns = ['Article Description', 'RP Type', 'Site', 'OM']
    for col in text_columns:
        df[col] = df[col].astype(str).fillna("").str.strip()
    
    # RP Type validation
    invalid_rp_mask = ~df['RP Type'].isin(['ND', 'RF'])
    if invalid_rp_mask.any():
        df.loc[invalid_rp_mask, 'RP Type'] = 'RF'  # Default to RF
        notes.append("無效的RP Type值已修正為RF")
    
    # Add effective sales quantity
    df['Effective Sold Qty'] = df.apply(
        lambda row: row['Last Month Sold Qty'] if row['Last Month Sold Qty'] > 0 
        else row['MTD Sold Qty'], axis=1
    )
    
    # Add notes column
    df['Notes'] = ""
    
    return df, notes

# Core transfer logic - Conservative Transfer (Option A)
def calculate_conservative_transfer(df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, int]]:
    """Calculate transfer recommendations using conservative strategy"""
    transfer_out: List[Dict[str, Any]] = []  # Transfer-out candidates
    transfer_in: List[Dict[str, Any]] = []   # Receive candidates
    stats: Dict[str, int] = {
        'total_transfer_out': 0,
        'total_transfer_in': 0,
        'nd_transfer_count': 0,
        'rf_surplus_count': 0,
        'emergency_count': 0,
        'potential_count': 0
    }
    
    # Group by Article and OM
    grouped = df.groupby(['Article', 'OM'])
    
    for (article, om), group in grouped:
        # Calculate max sales quantity for the group
        max_sold_qty = group['Effective Sold Qty'].max()
        
        for _, row in group.iterrows():
            site = str(row['Site'])
            rp_type = str(row['RP Type'])
            net_stock = int(row['SaSa Net Stock'])
            pending = int(row['Pending Received'])
            safety_stock = int(row['Safety Stock'])
            sold_qty = int(row['Effective Sold Qty'])
            moq = int(row['MOQ'])
            
            # Priority 1: ND Type Complete Transfer-out
            if rp_type == 'ND' and net_stock > 0:
                transfer_out.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Transfer Qty': net_stock,
                    'Transfer Type': 'ND轉出',
                    'Priority': 1,
                    'Original Stock': net_stock,
                    'After Transfer Stock': 0,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['nd_transfer_count'] += 1
                stats['total_transfer_out'] += net_stock
            
            # Priority 2: RF Type Surplus Transfer-out
            elif (rp_type == 'RF' and 
                  (net_stock + pending) > safety_stock and
                  sold_qty < max_sold_qty):
                
                base_transferable = (net_stock + pending) - safety_stock
                upper_limit = max(int((net_stock + pending) * 0.2), 2)
                actual_transfer = min(base_transferable, upper_limit, net_stock)
                
                if actual_transfer >= 2:  # Minimum 2 pieces
                    transfer_out.append({
                        'Article': article,
                        'Article Description': str(row['Article Description']),
                        'OM': om,
                        'Site': site,
                        'RP Type': rp_type,
                        'Transfer Qty': actual_transfer,
                        'Transfer Type': 'RF過剩轉出',
                        'Priority': 2,
                        'Original Stock': net_stock,
                        'After Transfer Stock': net_stock - actual_transfer,
                        'Safety Stock': safety_stock,
                        'MOQ': moq
                    })
                    stats['rf_surplus_count'] += 1
                    stats['total_transfer_out'] += actual_transfer
            
            # Priority 1: Emergency Shortage Replenishment
            if (rp_type == 'RF' and 
                net_stock == 0 and 
                sold_qty > 0):
                
                transfer_in.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Needed Qty': safety_stock,
                    'Receive Type': '緊急缺貨補貨',
                    'Priority': 1,
                    'Current Stock': net_stock,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['emergency_count'] += 1
                stats['total_transfer_in'] += safety_stock
            
            # Priority 2: Potential Shortage Replenishment
            elif (rp_type == 'RF' and 
                  (net_stock + pending) < safety_stock and
                  sold_qty == max_sold_qty):
                
                needed_qty = safety_stock - (net_stock + pending)
                transfer_in.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Needed Qty': needed_qty,
                    'Receive Type': '潛在缺貨補貨',
                    'Priority': 2,
                    'Current Stock': net_stock,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['potential_count'] += 1
                stats['total_transfer_in'] += needed_qty
    
    return transfer_out, transfer_in, stats

# Core transfer logic - Enhanced Transfer (Option B)
def calculate_enhanced_transfer(df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, int]]:
    """Calculate transfer recommendations using enhanced strategy"""
    transfer_out: List[Dict[str, Any]] = []  # Transfer-out candidates
    transfer_in: List[Dict[str, Any]] = []   # Receive candidates
    stats: Dict[str, int] = {
        'total_transfer_out': 0,
        'total_transfer_in': 0,
        'nd_transfer_count': 0,
        'rf_enhanced_count': 0,
        'emergency_count': 0,
        'potential_count': 0
    }
    
    # Group by Article and OM
    grouped = df.groupby(['Article', 'OM'])
    
    for (article, om), group in grouped:
        # Calculate max sales quantity for the group
        max_sold_qty = group['Effective Sold Qty'].max()
        
        # Sort by sales quantity (low to high) for enhanced transfer
        group_sorted = group.sort_values('Effective Sold Qty', ascending=True)
        
        for _, row in group_sorted.iterrows():
            site = str(row['Site'])
            rp_type = str(row['RP Type'])
            net_stock = int(row['SaSa Net Stock'])
            pending = int(row['Pending Received'])
            safety_stock = int(row['Safety Stock'])
            sold_qty = int(row['Effective Sold Qty'])
            moq = int(row['MOQ'])
            
            # Priority 1: ND Type Complete Transfer-out
            if rp_type == 'ND' and net_stock > 0:
                transfer_out.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Transfer Qty': net_stock,
                    'Transfer Type': 'ND轉出',
                    'Priority': 1,
                    'Original Stock': net_stock,
                    'After Transfer Stock': 0,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['nd_transfer_count'] += 1
                stats['total_transfer_out'] += net_stock
            
            # Priority 2: RF Type Enhanced Transfer-out
            elif (rp_type == 'RF' and 
                  (net_stock + pending) > (moq + 1) and
                  sold_qty < max_sold_qty):
                
                base_transferable = (net_stock + pending) - (moq + 1)
                upper_limit = max(int((net_stock + pending) * 0.5), 2)
                actual_transfer = min(base_transferable, upper_limit, net_stock)
                
                if actual_transfer >= 2:  # Minimum 2 pieces
                    transfer_out.append({
                        'Article': article,
                        'Article Description': str(row['Article Description']),
                        'OM': om,
                        'Site': site,
                        'RP Type': rp_type,
                        'Transfer Qty': actual_transfer,
                        'Transfer Type': 'RF加強轉出',
                        'Priority': 2,
                        'Original Stock': net_stock,
                        'After Transfer Stock': net_stock - actual_transfer,
                        'Safety Stock': safety_stock,
                        'MOQ': moq
                    })
                    stats['rf_enhanced_count'] += 1
                    stats['total_transfer_out'] += actual_transfer
            
            # Priority 1: Emergency Shortage Replenishment
            if (rp_type == 'RF' and 
                net_stock == 0 and 
                sold_qty > 0):
                
                transfer_in.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Needed Qty': safety_stock,
                    'Receive Type': '緊急缺貨補貨',
                    'Priority': 1,
                    'Current Stock': net_stock,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['emergency_count'] += 1
                stats['total_transfer_in'] += safety_stock
            
            # Priority 2: Potential Shortage Replenishment
            elif (rp_type == 'RF' and 
                  (net_stock + pending) < safety_stock and
                  sold_qty == max_sold_qty):
                
                needed_qty = safety_stock - (net_stock + pending)
                transfer_in.append({
                    'Article': article,
                    'Article Description': str(row['Article Description']),
                    'OM': om,
                    'Site': site,
                    'RP Type': rp_type,
                    'Needed Qty': needed_qty,
                    'Receive Type': '潛在缺貨補貨',
                    'Priority': 2,
                    'Current Stock': net_stock,
                    'Safety Stock': safety_stock,
                    'MOQ': moq
                })
                stats['potential_count'] += 1
                stats['total_transfer_in'] += needed_qty
    
    return transfer_out, transfer_in, stats

# Match transfer out and in candidates (OM-based matching - for modes A, B, C)
def match_transfer_recommendations_om(transfer_out: List[Dict[str, Any]], transfer_in: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Match transfer out and receive candidates with OM matching"""
    recommendations: List[Dict[str, Any]] = []
    
    # Sort by priority
    transfer_out_sorted = sorted(transfer_out, key=lambda x: x['Priority'])
    transfer_in_sorted = sorted(transfer_in, key=lambda x: x['Priority'])
    
    # Matching logic - OM based (same Article and same OM)
    for receiver in transfer_in_sorted:
        remaining_need = receiver['Needed Qty']
        
        for supplier in transfer_out_sorted:
            if (supplier['Article'] == receiver['Article'] and 
                supplier['OM'] == receiver['OM'] and
                supplier['Site'] != receiver['Site'] and
                supplier['Transfer Qty'] > 0 and
                remaining_need > 0):
                
                transfer_amount = min(supplier['Transfer Qty'], remaining_need)
                
                # Quantity optimization: try to make it at least 2 pieces
                if transfer_amount == 1 and supplier['Transfer Qty'] >= 2:
                    transfer_amount = 2
                
                recommendation = {
                    'Article': supplier['Article'],
                    'Article Description': supplier['Article Description'],
                    'OM': supplier['OM'],
                    'Transfer Site': supplier['Site'],
                    'Receive Site': receiver['Site'],
                    'Transfer Qty': transfer_amount,
                    'Transfer Type': supplier['Transfer Type'],
                    'Receive Type': receiver['Receive Type'],
                    'Original Stock': supplier['Original Stock'],
                    'After Transfer Stock': supplier['After Transfer Stock'],
                    'Safety Stock': supplier['Safety Stock'],
                    'MOQ': supplier['MOQ'],
                    'Notes': f"Matched {supplier['Transfer Type']} with {receiver['Receive Type']}"
                }
                
                recommendations.append(recommendation)
                
                # Update quantities
                supplier['Transfer Qty'] -= transfer_amount
                remaining_need -= transfer_amount
                
                if remaining_need <= 0:
                    break
    
    return recommendations

# Match transfer out and in candidates (Cross-OM matching - for mode D)
def match_transfer_recommendations_cross_om(transfer_out: List[Dict[str, Any]], transfer_in: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Match transfer out and receive candidates with cross-OM matching"""
    recommendations: List[Dict[str, Any]] = []
    
    # Sort by priority
    transfer_out_sorted = sorted(transfer_out, key=lambda x: x['Priority'])
    transfer_in_sorted = sorted(transfer_in, key=lambda x: x['Priority'])
    
    # Group sites by Hong Kong/Macau pattern
    def get_site_group(site: str) -> str:
        """Group sites by HA/HB/HC pattern or HD pattern"""
        if site.startswith(('HA', 'HB', 'HC')):
            return 'HA_HB_HC'
        elif site.startswith('HD'):
            return 'HD'
        return 'OTHER'
    
    # Matching logic - Cross OM (same Article only, allow different OM within same site group)
    for receiver in transfer_in_sorted:
        remaining_need = receiver['Needed Qty']
        receiver_group = get_site_group(receiver['Site'])
        
        for supplier in transfer_out_sorted:
            supplier_group = get_site_group(supplier['Site'])
            
            if (supplier['Article'] == receiver['Article'] and 
                supplier_group == receiver_group and
                supplier['Site'] != receiver['Site'] and
                supplier['Transfer Qty'] > 0 and
                remaining_need > 0):
                
                transfer_amount = min(supplier['Transfer Qty'], remaining_need)
                
                # Quantity optimization: try to make it at least 2 pieces
                if transfer_amount == 1 and supplier['Transfer Qty'] >= 2:
                    transfer_amount = 2
                
                recommendation = {
                    'Article': supplier['Article'],
                    'Article Description': supplier['Article Description'],
                    'OM': f"{supplier['OM']}→{receiver['OM']}",
                    'Transfer Site': supplier['Site'],
                    'Receive Site': receiver['Site'],
                    'Transfer Qty': transfer_amount,
                    'Transfer Type': supplier['Transfer Type'],
                    'Receive Type': receiver['Receive Type'],
                    'Original Stock': supplier['Original Stock'],
                    'After Transfer Stock': supplier['After Transfer Stock'],
                    'Safety Stock': supplier['Safety Stock'],
                    'MOQ': supplier['MOQ'],
                    'Notes': f"Cross-OM: {supplier['Transfer Type']} with {receiver['Receive Type']}"
                }
                
                recommendations.append(recommendation)
                
                # Update quantities
                supplier['Transfer Qty'] -= transfer_amount
                remaining_need -= transfer_amount
                
                if remaining_need <= 0:
                    break
    
    return recommendations

# Generate statistics
def generate_statistics(recommendations: List[Dict[str, Any]], df: pd.DataFrame) -> Dict[str, Any]:
    """Generate comprehensive statistics"""
    if not recommendations:
        return {}
    
    rec_df = pd.DataFrame(recommendations)
    
    stats: Dict[str, Any] = {
        'total_recommendations': len(recommendations),
        'total_transfer_qty': int(rec_df['Transfer Qty'].sum()),
        'unique_articles': rec_df['Article'].nunique(),
        'unique_oms': rec_df['OM'].nunique(),
        
        # By article statistics
        'by_article': rec_df.groupby('Article').agg({
            'Transfer Qty': ['sum', 'count'],
            'OM': 'nunique'
        }).reset_index(),
        
        # By OM statistics
        'by_om': rec_df.groupby('OM').agg({
            'Transfer Qty': ['sum', 'count'],
            'Article': 'nunique'
        }).reset_index(),
        
        # Transfer type distribution
        'transfer_type_dist': rec_df.groupby('Transfer Type').agg({
            'Transfer Qty': ['sum', 'count']
        }).reset_index(),
        
        # Receive type distribution
        'receive_type_dist': rec_df.groupby('Receive Type').agg({
            'Transfer Qty': ['sum', 'count']
        }).reset_index()
    }
    
    return stats

# Create visualization
def create_visualization(stats: Dict, transfer_type: str):
    """Create matplotlib visualization"""
    if not stats:
        return None
    
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Prepare data for visualization
    om_stats = stats['by_om']
    transfer_stats = stats['transfer_type_dist']
    receive_stats = stats['receive_type_dist']
    
    # Create bar chart
    categories = om_stats['OM'].tolist()
    bar_width = 0.2
    x_pos = np.arange(len(categories))
    
    # Get values for each transfer type
    nd_values = []
    rf_values = []
    emergency_values = []
    potential_values = []
    
    for om in categories:
        om_data = stats['by_om'][stats['by_om']['OM'] == om]
        nd_data = stats['transfer_type_dist'][stats['transfer_type_dist']['Transfer Type'] == 'ND轉出']
        rf_data = stats['transfer_type_dist'][stats['transfer_type_dist']['Transfer Type'] == 
                 ('RF過剩轉出' if transfer_type == 'A' else 'RF加強轉出')]
        emergency_data = stats['receive_type_dist'][stats['receive_type_dist']['Receive Type'] == '緊急缺貨補貨']
        potential_data = stats['receive_type_dist'][stats['receive_type_dist']['Receive Type'] == '潛在缺貨補貨']
        
        nd_values.append(nd_data[('Transfer Qty', 'sum')].iloc[0] if not nd_data.empty else 0)
        rf_values.append(rf_data[('Transfer Qty', 'sum')].iloc[0] if not rf_data.empty else 0)
        emergency_values.append(emergency_data[('Transfer Qty', 'sum')].iloc[0] if not emergency_data.empty else 0)
        potential_values.append(potential_data[('Transfer Qty', 'sum')].iloc[0] if not potential_data.empty else 0)
    
    # Plot bars
    ax.bar(x_pos - 1.5*bar_width, nd_values, bar_width, label='ND Transfer Out', alpha=0.8)
    ax.bar(x_pos - 0.5*bar_width, rf_values, bar_width, 
           label='RF Surplus Transfer Out' if transfer_type == 'A' else 'RF Enhanced Transfer Out', alpha=0.8)
    ax.bar(x_pos + 0.5*bar_width, emergency_values, bar_width, label='Emergency Receive', alpha=0.8)
    ax.bar(x_pos + 1.5*bar_width, potential_values, bar_width, label='Potential Receive', alpha=0.8)
    
    # Customize chart
    ax.set_xlabel('OM Units')
    ax.set_ylabel('Transfer Quantity')
    ax.set_title('OM Transfer vs Receive Analysis')
    ax.set_xticks(x_pos)
    ax.set_xticklabels(categories, rotation=45)
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

# Generate Excel export
def generate_excel_export(recommendations: List[Dict], stats: Dict, notes: List[str]) -> bytes:
    """Generate Excel file with multiple sheets"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Transfer Recommendations
        if recommendations:
            rec_df = pd.DataFrame(recommendations)
            rec_df.to_excel(writer, sheet_name='Transfer Recommendations', index=False)
        
        # Sheet 2: Statistical Summary
        summary_data = {
            'Metric': [
                'Total Recommendations',
                'Total Transfer Quantity',
                'Unique Articles',
                'Unique OMs'
            ],
            'Value': [
                stats.get('total_recommendations', 0),
                stats.get('total_transfer_qty', 0),
                stats.get('unique_articles', 0),
                stats.get('unique_oms', 0)
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Statistical Summary', startrow=0, index=False)
        
        # Add detailed statistics
        row_offset = len(summary_df) + 3
        
        # By Article statistics
        if 'by_article' in stats and not stats['by_article'].empty:
            article_stats = stats['by_article'].copy()
            article_stats.columns = ['Article', 'Total Transfer Qty', 'Number of Transfers', 'Number of OMs']
            article_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=row_offset, index=False)
            row_offset += len(article_stats) + 3
        
        # By OM statistics
        if 'by_om' in stats and not stats['by_om'].empty:
            om_stats = stats['by_om'].copy()
            om_stats.columns = ['OM', 'Total Transfer Qty', 'Number of Transfers', 'Number of Articles']
            om_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=row_offset, index=False)
            row_offset += len(om_stats) + 3
        
        # Transfer Type distribution
        if 'transfer_type_dist' in stats and not stats['transfer_type_dist'].empty:
            transfer_stats = stats['transfer_type_dist'].copy()
            transfer_stats.columns = ['Transfer Type', 'Total Quantity', 'Number of Transfers']
            transfer_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=row_offset, index=False)
            row_offset += len(transfer_stats) + 3
        
        # Receive Type distribution
        if 'receive_type_dist' in stats and not stats['receive_type_dist'].empty:
            receive_stats = stats['receive_type_dist'].copy()
            receive_stats.columns = ['Receive Type', 'Total Quantity', 'Number of Transfers']
            receive_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=row_offset, index=False)
        
        # Sheet 3: Data Processing Notes
        if notes:
            notes_df = pd.DataFrame({'Processing Notes': notes})
            notes_df.to_excel(writer, sheet_name='Processing Notes', index=False)
    
    output.seek(0)
    return output.getvalue()

# Main application
def main():
    st.title("📦 調貨建議生成系統")
    st.markdown("---")
    
    render_sidebar()
    
    # File upload section
    st.subheader("📁 數據上傳")
    uploaded_file = st.file_uploader(
        "上傳Excel文件", 
        type=['xlsx', 'xls'],
        help="请上传包含库存数据的Excel文件"
    )
    
    if uploaded_file is not None:
        try:
            # Read and validate data
            with st.spinner("正在讀取和驗證數據..."):
                df = pd.read_excel(uploaded_file)
                df, notes = validate_and_preprocess_data(df)
            
            # Data preview section
            st.subheader("📊 數據預覽")
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("總數據行數", len(df))
                st.metric("唯一產品數", df['Article'].nunique())
            
            with col2:
                st.metric("OM單位數", df['OM'].nunique())
                st.metric("店鋪數", df['Site'].nunique())
            
            st.dataframe(df.head(10), use_container_width=True)
            
            # Show processing notes
            if notes:
                with st.expander("數據預處理備註"):
                    for note in notes:
                        st.info(f"• {note}")
            
            # Transfer strategy selection
            st.subheader("⚙️ 調貨策略選擇")
            transfer_type = st.radio(
                "選擇調貨策略:",
                options=['A:保守轉貨', 'B:加強轉貨'],
                horizontal=True
            )
            
            if st.button("🚀 開始分析", type="primary"):
                with st.spinner("正在計算調貨建議..."):
                    if transfer_type == 'A:保守轉貨':
                        transfer_out, transfer_in, transfer_stats = calculate_conservative_transfer(df)
                    else:
                        transfer_out, transfer_in, transfer_stats = calculate_enhanced_transfer(df)
                    
                    recommendations = match_transfer_recommendations_om(transfer_out, transfer_in)
                    stats = generate_statistics(recommendations, df)
                
                # Display results
                st.subheader("📈 分析結果")
                
                if recommendations:
                    # KPI metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("總調貨建議數", stats['total_recommendations'])
                    with col2:
                        st.metric("總調貨件數", f"{stats['total_transfer_qty']:,.0f}")
                    with col3:
                        st.metric("涉及產品數", stats['unique_articles'])
                    with col4:
                        st.metric("涉及OM數", stats['unique_oms'])
                    
                    # Transfer recommendations table
                    st.subheader("📋 調貨建議明細")
                    rec_df = pd.DataFrame(recommendations)
                    st.dataframe(rec_df, use_container_width=True)
                    
                    # Statistics tables
                    st.subheader("📊 統計分析")
                    
                    tab1, tab2, tab3, tab4 = st.tabs([
                        "按產品統計", "按OM統計", "轉出類型分佈", "接收類型分佈"
                    ])
                    
                    with tab1:
                        st.dataframe(stats['by_article'], use_container_width=True)
                    
                    with tab2:
                        st.dataframe(stats['by_om'], use_container_width=True)
                    
                    with tab3:
                        st.dataframe(stats['transfer_type_dist'], use_container_width=True)
                    
                    with tab4:
                        st.dataframe(stats['receive_type_dist'], use_container_width=True)
                    
                    # Visualization
                    st.subheader("📊 可視化分析")
                    fig = create_visualization(stats, transfer_type)
                    if fig:
                        st.pyplot(fig)
                    
                    # Export functionality
                    st.subheader("💾 匯出結果")
                    excel_data = generate_excel_export(recommendations, stats, notes)
                    
                    st.download_button(
                        label="📥 下載Excel報告",
                        data=excel_data,
                        file_name=f"調貨建議_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ 分析完成！")
                
                else:
                    st.info("ℹ️ 當前數據無需調貨建議")
        
        except Exception as e:
            st.error(f"❌ 處理文件時發生錯誤: {str(e)}")
            logger.error(f"Error processing file: {str(e)}")
    
    else:
        st.info("ℹ️ 請上傳Excel文件開始分析")
        
        # File format requirements
        with st.expander("📋 文件格式要求"):
            st.markdown("""
            ### 必需欄位:
            - **Article**: 產品編號 (文字)
            - **Article Description**: 產品描述 (文字)
            - **RP Type**: 補貨類型 (ND/RF)
            - **Site**: 店鋪編號 (文字)
            - **OM**: 營運管理單位 (文字)
            - **MOQ**: 最低派貨數量 (數字)
            - **SaSa Net Stock**: 現有庫存數量 (數字)
            - **Pending Received**: 在途訂單數量 (數字)
            - **Safety Stock**: 安全庫存數量 (數字)
            - **Last Month Sold Qty**: 上月銷量 (數字)
            - **MTD Sold Qty**: 本月至今銷量 (數字)
            
            ### 數據預處理規則:
            - Article欄位強制轉換為字串類型
            - 所有數量欄位轉換為整數，無效值填充為0
            - 負值庫存和銷量自動修正為0
            - 銷量異常值(>100000)限制為100000
            - 字串欄位空值填充為空字串
            """)

if __name__ == "__main__":
    main()