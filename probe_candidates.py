import sys
import pandas as pd
from utils import _calculate_candidates


def main():
    path = sys.argv[1]
    article = sys.argv[2]
    mode = sys.argv[3]
    engine = 'openpyxl' if path.lower().endswith('xlsx') else 'xlrd'
    df = pd.read_excel(path, engine=engine)
    df['Effective Sold Qty'] = df.apply(lambda r: r['Last Month Sold Qty'] if r['Last Month Sold Qty'] > 0 else r['MTD Sold Qty'], axis=1)
    df = df[df['Article'].astype(str) == article]
    senders, receivers = _calculate_candidates(df, mode)
    print('Receivers:')
    for r in receivers:
        d = r['data']
        print(d['Site'], d['OM'], d['RP Type'], r['needed_qty'], r['type'], r['effective_sales'])


if __name__ == '__main__':
    main()

