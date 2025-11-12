import sys
import pandas as pd
from utils import preprocess_data, estimate_transfer_potential, generate_recommendations, generate_excel_export


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\BestO\AI\Sep2025_App\Trea reallocation_Gemini 2.5Pro\MAY_12Nov2025.XLSX"
    engine = 'openpyxl' if path.lower().endswith('xlsx') else 'xlrd'
    df = pd.read_excel(path, engine=engine)

    processed_df, logs = preprocess_data(df.copy())
    print('預處理日誌:')
    for l in logs:
        print('-', l)
    if processed_df is None:
        print('預處理失敗，終止測試')
        return

    potential = estimate_transfer_potential(processed_df.copy())
    print('潛在調貨量預估:')
    print(potential)

    modes = ['A: 保守轉貨', 'B: 加強轉貨', 'C: 重點補0']
    for mode in modes:
        print(f'\n模式: {mode}')
        rec_df, kpi_metrics, stats_by_article, stats_by_om, transfer_type_dist, receive_type_dist = generate_recommendations(processed_df.copy(), mode)
        if rec_df.empty:
            print('無建議')
        else:
            print('KPI:')
            print(kpi_metrics)
            print('樣本:')
            print(rec_df[['Article','OM','Transfer Site','Receive Site','Transfer Qty','Notes']].head(20))
            excel_bytes = generate_excel_export(rec_df, kpi_metrics, stats_by_article, stats_by_om, transfer_type_dist, receive_type_dist, mode)
            out_name = f"測試輸出_調貨建議_{mode[0]}.xlsx"
            with open(out_name, 'wb') as f:
                f.write(excel_bytes)
            print(f'已輸出報告: {out_name}')


if __name__ == '__main__':
    main()

