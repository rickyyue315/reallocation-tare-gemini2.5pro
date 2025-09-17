import pandas as pd
import numpy as np
from app import validate_and_preprocess_data, calculate_conservative_transfer, calculate_enhanced_transfer
from app import match_transfer_recommendations_om, generate_statistics, create_visualization

def test_functionality():
    print("🧪 Testing System Functionality")
    print("=" * 50)
    
    # Load test data
    print("📊 Loading test data...")
    df = pd.read_excel("test_data_20250918_000610.xlsx")
    df, notes = validate_and_preprocess_data(df)
    print(f"✅ Data loaded - Shape: {df.shape}")
    print(f"   Columns: {list(df.columns)}")
    
    # Test both transfer modes
    print("\n⚙️ Testing Transfer Modes:")
    
    # Test A: Conservative Transfer
    print("\nA: 保守轉貨模式")
    transfer_out_a, transfer_in_a, transfer_stats_a = calculate_conservative_transfer(df)
    recommendations_a = match_transfer_recommendations_om(transfer_out_a, transfer_in_a)
    stats_a = generate_statistics(recommendations_a, df)
    
    print(f"   ✅ 轉出候選: {len(transfer_out_a)} 條")
    print(f"   ✅ 接收候選: {len(transfer_in_a)} 條") 
    print(f"   ✅ 匹配建議: {len(recommendations_a)} 條")
    
    # Test B: Enhanced Transfer
    print("\nB: 加強轉貨模式")
    transfer_out_b, transfer_in_b, transfer_stats_b = calculate_enhanced_transfer(df)
    recommendations_b = match_transfer_recommendations_om(transfer_out_b, transfer_in_b)
    stats_b = generate_statistics(recommendations_b, df)
    
    print(f"   ✅ 轉出候選: {len(transfer_out_b)} 條")
    print(f"   ✅ 接收候選: {len(transfer_in_b)} 條")
    print(f"   ✅ 匹配建議: {len(recommendations_b)} 條")
    
    # Test visualization
    print("\n📈 Testing Visualization:")
    fig_a = create_visualization(stats_a, 'A:保守轉貨')
    fig_b = create_visualization(stats_b, 'B:加強轉貨')
    
    if fig_a and fig_b:
        print("✅ A/B模式圖表生成成功")
        print("   A模式標籤: RF過剩轉出")
        print("   B模式標籤: RF加強轉貨")
    else:
        print("❌ 圖表生成失敗")
    
    # Check statistics
    print("\n📊 Checking Statistics:")
    print(f"   A模式總調貨件數: {stats_a['total_transfer_qty']:,.0f}")
    print(f"   B模式總調貨件數: {stats_b['total_transfer_qty']:,.0f}")
    
    print("\n🎉 所有功能測試完成！")
    print("📍 請訪問 http://localhost:8502 查看完整界面")

if __name__ == "__main__":
    test_functionality()