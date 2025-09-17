import pandas as pd
import numpy as np
from datetime import datetime
import os

def generate_test_data():
    """Generate test data for the transfer recommendation system"""
    
    # Create sample data
    np.random.seed(42)
    
    # Sample data
    articles = [f'ART{i:03d}' for i in range(1, 21)]
    article_descriptions = [f'Product {i}' for i in range(1, 21)]
    sites = [f'S{i:03d}' for i in range(1, 11)]
    oms = [f'OM{i:02d}' for i in range(1, 6)]
    rp_types = ['ND', 'RF']
    
    data = []
    
    for article, desc in zip(articles, article_descriptions):
        for site in sites:
            for om in oms:
                rp_type = np.random.choice(rp_types, p=[0.3, 0.7])
                
                # Generate realistic values
                moq = np.random.choice([1, 2, 3, 5, 10, 12, 24])
                net_stock = np.random.randint(0, 50)
                pending = np.random.randint(0, 20)
                safety_stock = np.random.randint(5, 30)
                last_month = np.random.randint(0, 100)
                mtd = np.random.randint(0, 50)
                
                # Create some edge cases
                if np.random.random() < 0.1:
                    net_stock = 0  # Out of stock
                if np.random.random() < 0.1:
                    pending = 0  # No pending orders
                if np.random.random() < 0.1:
                    last_month = 0  # No sales last month
                
                data.append({
                    'Article': article,
                    'Article Description': desc,
                    'RP Type': rp_type,
                    'Site': site,
                    'OM': om,
                    'MOQ': moq,
                    'SaSa Net Stock': net_stock,
                    'Pending Received': pending,
                    'Safety Stock': safety_stock,
                    'Last Month Sold Qty': last_month,
                    'MTD Sold Qty': mtd
                })
    
    df = pd.DataFrame(data)
    
    # Save to Excel
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'test_data_{timestamp}.xlsx'
    df.to_excel(filename, index=False)
    
    print(f"Test data generated: {filename}")
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"RP Type distribution:")
    print(df['RP Type'].value_counts())
    print(f"\nSample data:")
    print(df.head())
    
    return filename

if __name__ == "__main__":
    generate_test_data()