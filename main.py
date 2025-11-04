# import scraper 
import os
import time
import shutil
import pandas as pd

def pega_valores_cana() -> dict:
    safra_cana_df = pd.read_excel('./geral/Cana.xlsx', sheet_name='Área, produtividade e produção')
    for i, r in safra_cana_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL':
            return {'area': r['Unnamed: 2'], 'produtividade': r['Unnamed: 5'], 'producao': r['Unnamed: 8']}
            
    # A condição `safra_cana_df.iloc[:, 0] == 'Brasil'` é usada para filtrar as linhas.
    # O resultado em `brasil_df` conterá as LINHAS INTEIRAS (com todas as colunas) que correspondem à condição.
def pega_valores_graos() -> dict:
    safra_graos_df = pd.read_excel('./geral/Graos.xlsx', sheet_name='Brasil - Total por Produto')
    for i, r in safra_graos_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL (2)':
            return {'area': r['Unnamed: 2'], 'produtividade': r['Unnamed: 5'], 'producao': r['Unnamed: 8']}

def pega_valores_cafe() -> dict:
    safra_cafe_df = pd.read_excel('./geral/Café.xls', sheet_name='1 Café Total')
    for i, r in safra_cafe_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL':
            return {'area': r['Unnamed: 2'], 'produtividade': r['Unnamed: 5'], 'producao': r['Unnamed: 8']}

    
valores_cana = pega_valores_cana()
print(valores_cana)
valores_graos = pega_valores_graos()
print(valores_graos)
valores_cafe = pega_valores_cafe()
print(valores_cafe)
