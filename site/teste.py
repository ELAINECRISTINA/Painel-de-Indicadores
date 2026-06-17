import pandas as pd
df_raw = pd.read_excel('dados_teste/VBP.xlsx', sheet_name=0)
print(df_raw.columns.tolist())
print(df_raw.head(3))