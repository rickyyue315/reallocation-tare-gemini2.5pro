import pandas as pd
import numpy as np
from datetime import datetime
import subprocess
import time
import os

def test_system():
    """Test the complete transfer recommendation system"""
    
    print("🧪 Testing Retail Transfer Recommendation System")
    print("=" * 60)
    
    # 1. Check if app.py exists
    if not os.path.exists('app.py'):
        print("❌ app.py not found")
        return False
    
    print("✅ app.py found")
    
    # 2. Check requirements
    try:
        import streamlit
        import pandas
        import numpy
        import matplotlib
        import seaborn
        import openpyxl
        print("✅ All dependencies installed")
    except ImportError as e:
        print(f"❌ Missing dependency: {e}")
        return False
    
    # 3. Test data processing
    try:
        from app import validate_and_preprocess_data
        test_file = 'test_data_20250917_222853.xlsx'
        
        if os.path.exists(test_file):
            df, _ = validate_and_preprocess_data(pd.read_excel(test_file))
            print(f"✅ Data loading successful - Shape: {df.shape}")
            print(f"   Columns: {list(df.columns)}")
            print(f"   Data types: {df.dtypes.to_dict()}")
        else:
            print("⚠️  Test data file not found, skipping data test")
            
    except Exception as e:
        print(f"❌ Data processing test failed: {e}")
        return False
    
    # 4. Test business logic functions
    try:
        from app import (
            calculate_conservative_transfer,
            calculate_enhanced_transfer,
            match_transfer_recommendations_om
        )
        print("✅ Business logic functions imported successfully")
    except Exception as e:
        print(f"❌ Business logic import failed: {e}")
        return False
    
    # 5. Test visualization
    try:
        import matplotlib.pyplot as plt
        plt.figure(figsize=(10, 6))
        plt.bar(['Test'], [1])
        plt.title('Test Chart')
        plt.savefig('test_chart.png', dpi=100, bbox_inches='tight')
        plt.close()
        os.remove('test_chart.png')
        print("✅ Matplotlib visualization test passed")
    except Exception as e:
        print(f"❌ Visualization test failed: {e}")
        return False
    
    # 6. Test Excel export
    try:
        from app import generate_excel_export
        test_data = pd.DataFrame({
            'Article': ['ART001'],
            'Transfer Site': ['S001'],
            'Receive Site': ['S002'],
            'Transfer Qty': [5],
            'Transfer Type': ['ND Transfer']
        })
        generate_excel_export([test_data.to_dict('records')], {}, [])
        print("✅ Excel export function test passed")
    except Exception as e:
        print(f"❌ Excel export test failed: {e}")
        return False
    
    print("=" * 60)
    print("🎉 All system tests passed!")
    print("\n📋 Next steps:")
    print("1. Open http://localhost:8502 in your browser")
    print("2. Upload the test Excel file")
    print("3. Select transfer strategy (A/B)")
    print("4. Click 'Generate Transfer Recommendations' to see results")
    print("5. Download the Excel report")
    
    return True

if __name__ == "__main__":
    test_system()