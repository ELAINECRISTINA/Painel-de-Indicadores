# teste_safra.py
import pandas as pd

arquivo = 'dados_teste/Safras.xlsx'
xl = pd.ExcelFile(arquivo)

for aba in xl.sheet_names:
    for h in range(0, 8):
        df = pd.read_excel(arquivo, sheet_name=aba, header=h)
        colunas = [c for c in df.columns if 'unnamed' not in str(c).lower()]
        if len(colunas) >= 2:
            print(f"\n✅ ABA='{aba}' | header={h}")
            print(f"   Colunas: {df.columns.tolist()}")
            print(df.head(3).to_string())
            break