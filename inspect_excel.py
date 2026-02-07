import pandas as pd

try:
    df = pd.read_excel('wzór.xlsx')
    print("Columns:", df.columns.tolist())
    print("\nFirst 5 rows:")
    print(df.head())
    print("\nData Types:")
    print(df.dtypes)
except Exception as e:
    print(f"Error: {e}")
