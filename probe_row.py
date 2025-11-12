import sys
import pandas as pd


def main():
    path = sys.argv[1]
    article = sys.argv[2]
    site = sys.argv[3]
    engine = 'openpyxl' if path.lower().endswith('xlsx') else 'xlrd'
    df = pd.read_excel(path, engine=engine)
    mask = (df['Article'].astype(str) == article) & (df['Site'].astype(str) == site)
    cols = ['Article','Article Description','RP Type','Site','OM','SaSa Net Stock','Pending Received','Safety Stock','MOQ','Last Month Sold Qty','MTD Sold Qty']
    print(df.loc[mask, cols].to_string(index=False))


if __name__ == '__main__':
    main()

