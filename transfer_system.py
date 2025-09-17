import pandas as pd
import numpy as np
from datetime import datetime
import logging
from typing import List, Dict, Tuple, Optional, Any, Union
import re

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TransferOptimizer:
    def __init__(self):
        self.transfer_recommendations = []
        self.quality_checks = []
    
    def read_and_validate_data(self, file_path: str) -> pd.DataFrame:
        """Read Excel file and perform data validation and transformation"""
        try:
            # Read Excel file
            df = pd.read_excel(file_path)
            logger.info(f"Successfully read file: {file_path}, shape: {df.shape}")
            
            # Data preprocessing and validation
            df = self._preprocess_data(df)
            
            return df
            
        except Exception as e:
            logger.error(f"Error reading file: {str(e)}")
            raise
    
    def _preprocess_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Data preprocessing and validation"""
        # Article field forced to 12-digit text format
        if 'Article' in df.columns:
            df['Article'] = df['Article'].astype(str).str.strip()
            # Remove non-digit characters, then pad to 12 digits
            df['Article'] = df['Article'].apply(lambda x: re.sub(r'\D', '', x).zfill(12))
        
        # Numeric field processing
        numeric_columns = ['SaSa Net Stock', 'Pending Received', 'Safety Stock', 
                          'Last Month Sold Qty', 'MTD Sold Qty']
        
        for col in numeric_columns:
            if col in df.columns:
                # Convert non-numeric values to NaN
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # Fill missing values
                df[col] = df[col].fillna(0)
                # Correct outliers (negative values to 0)
                df[col] = df[col].clip(lower=0)
                
                # Special handling for sales fields
                if col in ['Last Month Sold Qty', 'MTD Sold Qty']:
                    df[col] = df[col].clip(upper=100000)
        
        # Text field processing
        text_columns = ['OM', 'RP Type', 'Site']
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).fillna("").str.strip()
        
        # Add effective sales quantity field
        df['Effective Sold Qty'] = df.apply(
            lambda row: row['Last Month Sold Qty'] if row['Last Month Sold Qty'] > 0 
            else row['MTD Sold Qty'], axis=1
        )
        
        return df
    
    def identify_transfer_candidates(self, df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Identify transfer-out candidates and receive candidates"""
        suppliers: List[Dict[str, Any]] = []  # Transfer-out candidates
        receivers: List[Dict[str, Any]] = []  # Receive candidates
        
        # Process by Article+OM grouping
        grouped = df.groupby(['Article', 'OM'])
        
        for (article, om), group in grouped:
            # Calculate maximum sales quantity within group
            max_sold_qty = group['Effective Sold Qty'].max()
            
            for _, row in group.iterrows():
                site = row.get('Site', '')
                rp_type = row.get('RP Type', '')
                net_stock = row.get('SaSa Net Stock', 0)
                pending_received = row.get('Pending Received', 0)
                safety_stock = row.get('Safety Stock', 0)
                sold_qty = row.get('Effective Sold Qty', 0)
                
                # Transfer-out rule - Priority 1: ND type transfer-out
                if rp_type == 'ND':
                    suppliers.append({
                        'article': article,
                        'om': om,
                        'site': site,
                        'rp_type': rp_type,
                        'transferable_qty': net_stock,
                        'priority': 1,
                        'original_stock': net_stock
                    })
                
                # Transfer-out rule - Priority 2: RF type surplus transfer-out
                elif (rp_type == 'RF' and 
                      (float(net_stock) + float(pending_received)) > float(safety_stock) and
                      float(sold_qty) != float(max_sold_qty)):
                    # Calculate base transferable quantity
                    transferable = float(net_stock) + float(pending_received) - float(safety_stock)
                    
                    # Apply 20% upper limit: max 20% of (net_stock + pending_received)
                    max_transferable = int((float(net_stock) + float(pending_received)) * 0.2)
                    transferable = min(transferable, max_transferable)
                    
                    # Apply minimum 2 pieces requirement
                    if transferable < 2:
                        transferable = 0
                    
                    if transferable > 0:
                        suppliers.append({
                            'article': article,
                            'om': om,
                            'site': site,
                            'rp_type': rp_type,
                            'transferable_qty': int(transferable),
                            'priority': 2,
                            'original_stock': int(net_stock)
                        })
                
                # Receive rule - Priority 1: Emergency shortage replenishment
                if (rp_type == 'RF' and 
                    net_stock == 0 and 
                    sold_qty > 0):
                    receivers.append({
                        'article': article,
                        'om': om,
                        'site': site,
                        'rp_type': rp_type,
                        'needed_qty': safety_stock,
                        'priority': 1,
                        'current_stock': net_stock
                    })
                
                # Receive rule - Priority 2: Potential shortage replenishment
                elif (rp_type == 'RF' and 
                      (net_stock + pending_received) < safety_stock and
                      sold_qty == max_sold_qty):
                    needed = safety_stock - (net_stock + pending_received)
                    receivers.append({
                        'article': article,
                        'om': om,
                        'site': site,
                        'rp_type': rp_type,
                        'needed_qty': needed,
                        'priority': 2,
                        'current_stock': net_stock
                    })
        
        return suppliers, receivers
    
    def match_transfers(self, suppliers: List[Dict[str, Any]], receivers: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Match suppliers and receivers"""
        transfer_suggestions: List[Dict[str, Any]] = []
        
        # Sort by priority
        suppliers_sorted = sorted(suppliers, key=lambda x: x['priority'])
        receivers_sorted = sorted(receivers, key=lambda x: x['priority'])
        
        # Matching logic
        for receiver in receivers_sorted:
            remaining_need = receiver['needed_qty']
            
            for supplier in suppliers_sorted:
                if (supplier['article'] == receiver['article'] and 
                    supplier['om'] == receiver['om'] and
                    supplier['site'] != receiver['site'] and
                    supplier['transferable_qty'] > 0 and
                    remaining_need > 0):
                    
                    transfer_amount = min(supplier['transferable_qty'], remaining_need)
                    
                    # Create transfer suggestion
                    suggestion = {
                        'Article': supplier['article'],
                        'OM': supplier['om'],
                        'Transfer Site': supplier['site'],
                        'Receive Site': receiver['site'],
                        'Transfer Qty': transfer_amount,
                        'Transfer Type': 'ND' if supplier['priority'] == 1 else 'RF',
                        'Receive Priority': 'Emergency' if receiver['priority'] == 1 else 'Potential',
                        'Original Stock': supplier['original_stock'],
                        'Current Need': receiver['needed_qty']
                    }
                    
                    transfer_suggestions.append(suggestion)
                    
                    # Update remaining quantity
                    supplier['transferable_qty'] -= transfer_amount
                    remaining_need -= transfer_amount
                    
                    if remaining_need <= 0:
                        break
        
        return transfer_suggestions
    
    def run_quality_checks(self, transfer_suggestions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Perform quality checks"""
        quality_checks: List[Dict[str, Any]] = []
        
        for i, transfer in enumerate(transfer_suggestions):
            checks = {
                'index': i,
                'article_om_match': True,
                'positive_transfer_qty': True,
                'not_exceed_original_stock': True,
                'different_sites': True,
                'article_format_12_digit': True
            }
            
            # Check 1: Article and OM must match exactly
            if transfer['Article'] != transfer_suggestions[0]['Article'] or \
               transfer['OM'] != transfer_suggestions[0]['OM']:
                checks['article_om_match'] = False
            
            # Check 2: Transfer Qty must be positive integer
            if transfer['Transfer Qty'] <= 0 or not isinstance(transfer['Transfer Qty'], int):
                checks['positive_transfer_qty'] = False
            
            # Check 3: Transfer Qty cannot exceed supplier's original SaSa Net Stock
            if transfer['Transfer Qty'] > transfer['Original Stock']:
                checks['not_exceed_original_stock'] = False
            
            # Check 4: Transfer Site and Receive Site cannot be the same
            if transfer['Transfer Site'] == transfer['Receive Site']:
                checks['different_sites'] = False
            
            # Check 5: Article field must be 12-digit text format
            if len(str(transfer['Article'])) != 12 or not str(transfer['Article']).isdigit():
                checks['article_format_12_digit'] = False
            
            quality_checks.append(checks)
        
        return quality_checks
    
    def generate_output(self, df: pd.DataFrame, transfer_suggestions: List[Dict[str, Any]]) -> str:
        """Generate output file"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'transfer_suggestions_{timestamp}.xlsx'
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Worksheet 1: Transfer Suggestions
            transfer_df = pd.DataFrame(transfer_suggestions)
            if not transfer_df.empty:
                transfer_df.to_excel(writer, sheet_name='Transfer Suggestions', index=False)
            
            # Worksheet 2: Statistical Summary
            self._generate_summary_dashboard(writer, transfer_suggestions, df)
        
        logger.info(f"Output file generated: {output_file}")
        return output_file
    
    def _generate_summary_dashboard(self, writer, 
                                  transfer_suggestions: List[Dict[str, Any]], 
                                  original_df: pd.DataFrame):
        """Generate statistical summary"""
        if not transfer_suggestions:
            return
        
        transfer_df = pd.DataFrame(transfer_suggestions)
        
        # KPI Banner
        summary_data = {
            'Metric': ['Total Transfer Suggestions', 'Total Transfer Quantity'],
            'Value': [len(transfer_suggestions), transfer_df['Transfer Qty'].sum()]
        }
        kpi_df = pd.DataFrame(summary_data)
        kpi_df.to_excel(writer, sheet_name='Statistical Summary', startrow=0, index=False)
        
        # Statistics by Article
        article_stats = transfer_df.groupby('Article').agg({
            'Transfer Qty': 'sum',
            'OM': 'nunique'
        }).reset_index()
        article_stats.columns = ['Article', 'Total Transfer Quantity', 'Number of OMs Involved']
        article_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=5, index=False)
        
        # Statistics by OM
        om_stats = transfer_df.groupby('OM').agg({
            'Transfer Qty': 'sum',
            'Article': 'nunique'
        }).reset_index()
        om_stats.columns = ['OM', 'Total Transfer Quantity', 'Number of Articles Involved']
        om_stats.to_excel(writer, sheet_name='Statistical Summary', startrow=5 + len(article_stats) + 2, index=False)
        
        # Transfer Type Analysis
        transfer_type_stats = transfer_df.groupby('Transfer Type').agg({
            'Transfer Qty': ['count', 'sum']
        }).reset_index()
        transfer_type_stats.columns = ['Transfer Type', 'Number of Suggestions', 'Total Quantity']
        transfer_type_stats.to_excel(writer, sheet_name='Statistical Summary', 
                                   startrow=5 + len(article_stats) + len(om_stats) + 4, index=False)
        
        # Receive Priority Analysis
        priority_stats = transfer_df.groupby('Receive Priority').agg({
            'Transfer Qty': ['count', 'sum']
        }).reset_index()
        priority_stats.columns = ['Receive Priority', 'Number of Suggestions', 'Total Quantity']
        priority_stats.to_excel(writer, sheet_name='Statistical Summary', 
                              startrow=5 + len(article_stats) + len(om_stats) + len(transfer_type_stats) + 6, 
                              index=False)
    
    def process_file(self, file_path: str):
        """Process Excel file and generate transfer suggestions"""
        try:
            # 1. Read and validate data
            df = self.read_and_validate_data(file_path)
            
            # 2. Identify transfer candidates
            suppliers, receivers = self.identify_transfer_candidates(df)
            
            # 3. Match transfers
            transfer_suggestions = self.match_transfers(suppliers, receivers)
            
            # 4. Quality checks
            quality_checks = self.run_quality_checks(transfer_suggestions)
            
            # 5. Generate output
            output_file = self.generate_output(df, transfer_suggestions)
            
            # 6. Print summary
            self._print_summary(transfer_suggestions, quality_checks)
            
            return output_file, transfer_suggestions
            
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            raise
    
    def _print_summary(self, transfer_suggestions: List[Dict[str, Any]], quality_checks: List[Dict[str, Any]]):
        """Print processing result summary"""
        print("=" * 60)
        print("Transfer System Processing Result Summary")
        print("=" * 60)
        
        if transfer_suggestions:
            print(f"Total Transfer Suggestions: {len(transfer_suggestions)}")
            total_qty = sum(t['Transfer Qty'] for t in transfer_suggestions)
            print(f"Total Transfer Quantity: {total_qty}")
            
            # Check quality check results
            all_passed = all(
                all(check.values()) 
                for check in quality_checks 
                if isinstance(check, dict)
            )
            
            if all_passed:
                print("✅ All quality checks passed")
            else:
                print("⚠️  Some quality checks failed, please check detailed report")
        
        else:
            print("ℹ️  No transfer suggestions generated")

# Usage Example
if __name__ == "__main__":
    optimizer = TransferOptimizer()
    
    # Process sample file
    input_file = r"C:\Users\BestO\Dropbox\SASA\ELE_08Sep2025.XLSX"
    
    try:
        output_file, suggestions = optimizer.process_file(input_file)
        print(f"\nProcessing completed! Output file: {output_file}")
        
    except Exception as e:
        print(f"Error during processing: {str(e)}")