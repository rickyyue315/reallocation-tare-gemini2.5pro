import sys
import pandas as pd
from utils import preprocess_data, generate_recommendations


def main():
    path = sys.argv[1]
    article = sys.argv[2]
    site = sys.argv[3]
    mode = sys.argv[4]
    engine = 'openpyxl' if path.lower().endswith('xlsx') else 'xlrd'
    df = pd.read_excel(path, engine=engine)
    processed_df, _ = preprocess_data(df.copy())
    rec_df, *_ = generate_recommendations(processed_df.copy(), mode)
    m = (rec_df['Article'].astype(str) == article) & (rec_df['Receive Site'].astype(str) == site)
    print(rec_df.loc[m].to_string(index=False))


if __name__ == '__main__':
    main()

