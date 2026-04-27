import pandas as pd

df = pd.read_excel("data/base.xlsx", header=3, nrows=3)
print("=== 전체 컬럼 목록 ===")
for i, col in enumerate(df.columns):
    print(f"{i:3d}  {col}  |  {df.iloc[0,i]}")
