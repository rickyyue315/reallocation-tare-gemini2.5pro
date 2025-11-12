import pandas as pd
from utils import generate_recommendations, estimate_transfer_potential


def make_df():
    data = [
        # RF: receiver zero stock, zero safety -> should create initial demand
        {
            'Article': 'SKU1', 'Article Description': 'RF Item', 'RP Type': 'RF', 'Site': 'S1', 'OM': 'OM1',
            'MOQ': 2, 'SaSa Net Stock': 0, 'Pending Received': 0, 'Safety Stock': 0,
            'Last Month Sold Qty': 5, 'MTD Sold Qty': 2
        },
        # RF: sender with surplus
        {
            'Article': 'SKU1', 'Article Description': 'RF Item', 'RP Type': 'RF', 'Site': 'S2', 'OM': 'OM1',
            'MOQ': 2, 'SaSa Net Stock': 10, 'Pending Received': 0, 'Safety Stock': 5,
            'Last Month Sold Qty': 1, 'MTD Sold Qty': 0
        },
        # ND: receiver zero stock -> should create ND initial demand
        {
            'Article': 'SKU2', 'Article Description': 'ND Item', 'RP Type': 'ND', 'Site': 'S3', 'OM': 'OM2',
            'MOQ': 3, 'SaSa Net Stock': 0, 'Pending Received': 0, 'Safety Stock': 0,
            'Last Month Sold Qty': 0, 'MTD Sold Qty': 0
        },
        # ND: sender with stock
        {
            'Article': 'SKU2', 'Article Description': 'ND Item', 'RP Type': 'ND', 'Site': 'S4', 'OM': 'OM2',
            'MOQ': 3, 'SaSa Net Stock': 8, 'Pending Received': 0, 'Safety Stock': 0,
            'Last Month Sold Qty': 0, 'MTD Sold Qty': 0
        },
    ]
    return pd.DataFrame(data)


if __name__ == '__main__':
    df = make_df()
    print('Potential (A/B/C):')
    print(estimate_transfer_potential(df))

    for mode in ['A: 保守轉貨', 'B: 加強轉貨', 'C: 重點補0']:
        rec_df, *_ = generate_recommendations(df.copy(), mode)
        print(f'\nMode: {mode}')
        if rec_df.empty:
            print('No recommendations')
        else:
            print(rec_df[['Article','OM','Transfer Site','Receive Site','Transfer Qty','Notes']])

