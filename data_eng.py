# import scraper 
import os
import time
import shutil
import numpy as np
import pandas as pd
import statistics
import string
from pprint import pprint
from tools.dict_tools import add_values_in_months, fix_dict, monthly_median, indicadores_linhas
from tools.treatments_tools import convert_br_number
from datetime import datetime
from tools.conab_dict import month_numeric
from tools.llm import init_llm
from tools.passa_captcha_llm import *
import re 
os.environ["HF_HUB_DISABLE_SYMLINKS"] = "1"
from docling.document_converter import DocumentConverter
import torch

def pega_valores_cana() -> dict:
    safra_cana_df = pd.read_excel('./geral/Cana.xlsx', sheet_name='Área, produtividade e produção')
    area = ''
    produtividade = ''
    producao = ''
    quando = ''
    for i, r in safra_cana_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL':
            area = r['Unnamed: 2']
            produtividade = r['Unnamed: 5']
            producao = r['Unnamed: 8']

        if 'Estimativa' in str(r['Unnamed: 0']):
            quando = r['Unnamed: 0']
            quando = quando.split(' ')[-1]
            mes = quando.split('/')[0][0:3].title()
            ano = quando.split('/')[-1].replace('.','')
    return {'cana':{'area':area,'produtividade':produtividade, 'producao':producao, 'mes':mes, 'ano':ano}}  

    # A condição `safra_cana_df.iloc[:, 0] == 'Brasil'` é usada para filtrar as linhas.
    # O resultado em `brasil_df` conterá as LINHAS INTEIRAS (com todas as colunas) que correspondem à condição.
def pega_valores_graos() -> dict:
    safra_graos_df = pd.read_excel('./geral/Graos.xlsx', sheet_name='Brasil - Total por Produto')
    area = ''
    produtividade = ''
    producao = ''
    quando = ''
    for i, r in safra_graos_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL (2)':
            area = r['Unnamed: 2']
            produtividade = r['Unnamed: 5']
            producao = r['Unnamed: 8']

        if 'Estimativa' in str(r['Unnamed: 0']):
            quando = r['Unnamed: 0']
            quando = quando.split(' ')[-1]
            mes = quando.split('/')[0][0:3].title()
            ano = quando.split('/')[-1].replace('.','')
    return {'graos':{'area': area, 'produtividade': produtividade, 'producao': producao, 'mes':mes, 'ano':ano}}
def pega_valores_cafe() -> dict:
    safra_cafe_df = pd.read_excel('./geral/Café.xls', sheet_name='1 Café Total')
    area = ''
    produtividade = ''
    producao = ''
    quando = ''
    for i, r in safra_cafe_df.iterrows():
        if r['Unnamed: 0'] == 'BRASIL':
            area = r['Unnamed: 2']
            produtividade = r['Unnamed: 5']
            producao = r['Unnamed: 8']

        if 'Estimativa' in str(r['Unnamed: 0']):
            quando = r['Unnamed: 0']
            quando = quando.split(' ')[-1]
            mes = quando.split('/')[0][0:3].title()
            ano = quando.split('/')[-1].replace('.','')
    return {'cafe':{'area': area, 'produtividade': produtividade, 'producao': producao, 'mes':mes, 'ano':ano}}


def clean_conab_arch(df):
    df = df.drop(df.index[0:14])
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    
    # Filtrar diretamente as colunas que não são NaN
    df = df.loc[:, ~pd.isna(df.columns)]

    df = df.dropna(how='all')

    for i, r in df.iterrows():
        produto = r['Produto']
        if produto.lower() == 'produto':
            df = df.drop(index=i)
    return df

def catch_values_conab(path: str, choice: str) -> list:
    """
    Processa dados de produtos agrícolas do arquivo CONAB.
    
    Args:
        path: Caminho do arquivo Excel
        choice: Categoria dos produtos ('inseticidas', 'herbicidas', 'fertilizantes')
    
    Returns:
        Lista de dicionários com dados processados de cada produto
    """
    df = pd.read_excel(path)
    df = clean_conab_arch(df)
    
    # Obtém as listas de produtos
    lista_inseticidas, lista_fungicidas, lista_herbicidas = indicadores_linhas()
    
    # Mapeamento de categorias para listas de produtos
    categorias_produtos = {
        'inseticidas': lista_inseticidas,
        'fungicidas': lista_fungicidas,
        'herbicidas': lista_herbicidas,
        'fertilizantes': lista_inseticidas  # Ajuste se necessário
    }
    
    # Valida a escolha
    if choice not in categorias_produtos:
        raise ValueError(f'Choice {choice} not valid. Escolhas válidas: {list(categorias_produtos.keys())}')
    
    # Mapeamento de normalização de nomes (do DataFrame -> chave do dict)
    normalizacao_nomes = {
        'imidacloprid nortox': 'imidacloprid',
        '2,4-d': '2,4-D',
        'uréia': 'ureia',
        'cloreto de potássio': 'cloreto_de_potassio',
        'superfosfato simples': 'superfosfato_simples',
        'sulfato de amônio': 'sulfato_de_amonio',
        '20-05-20': 'NPK (20-05-20)'
    }
    
    # Obtém lista de produtos para a categoria escolhida
    lista_produtos = categorias_produtos[choice]
    
    # Inicializa estruturas para cada produto
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
             'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    
    produtos_dict = {}
    produtos_lower_map = {}  # Mapeia produto.lower() -> chave normalizada
    
    for produto in lista_produtos:
        produto_lower = produto.lower().strip()
        produto_key = normalizacao_nomes.get(produto_lower, produto_lower)
        
        produtos_lower_map[produto_lower] = produto_key
        
        produtos_dict[produto_key] = {
            produto_key: {
                'unidade': [],
                'precos': {mes: [] for mes in meses},
                'mediana_mensal': {mes: [] for mes in meses},
                'média': [],
                'mediana': []
            }
        }
    
    # Processa cada linha do DataFrame
    for i, r in df.iterrows():
        produto = r['Produto'].lower().strip()
        unidade = r['Unidade']
        
        # Verifica se o produto está na lista de produtos desejados
        if produto in produtos_lower_map:
            produto_key = produtos_lower_map[produto]
            
            # Adiciona unidade
            produtos_dict[produto_key][produto_key]['unidade'].append(unidade)
            
            # Adiciona valores mensais
            produtos_dict[produto_key] = add_values_in_months(
                produtos_dict[produto_key], 
                produto_key, 
                unidade, 
                r
            )
    
    # Processa cada produto (fix_dict e monthly_median) e monta lista final
    lista_resultado = []
    for produto in lista_produtos:
        produto_lower = produto.lower().strip()
        produto_key = normalizacao_nomes.get(produto_lower, produto_lower)
        
        # Aplica transformações
        produtos_dict[produto_key] = fix_dict(produtos_dict[produto_key], produto_key)
        produtos_dict[produto_key] = monthly_median(produtos_dict[produto_key], produto_key)
        
        # Adiciona à lista de resultado
        lista_resultado.append(produtos_dict[produto_key])
    
    return lista_resultado

def catch_dolar_median(path) -> dict:    
    df = pd.read_csv(path, header=None, sep=';')
    df['Dolar'] = df[5].apply(convert_br_number)
    media = float(df['Dolar'].mean())
    return {'Dolar':{'media':media,'mes':datetime.now().month}}
    
def catch_vbp_values(path) -> dict:
    df = pd.read_excel(path, sheet_name='Variação', header=None)
    for i, r in df.iterrows():
        if r[0] == 'VBP TOTAL':
            id_data = i + 1
            valor = r[6]

    data = df.iloc[id_data][0]
    data = data.split(';')[1]
    data = data.split(' a ')[-1] 
    
    mes = data.split('/')[0][0:3]
    mes_numeric = month_numeric[data.split('/')[0].title()]
    print(mes)
    print(mes_numeric)
    return {'VBP':{'valor':valor,'mes':mes,'mes_num':mes_numeric,'ano':data.split('/')[-1]}}

def catch_credito_rural_values(path) -> dict:
    df = pd.read_excel(path, header=None)
    df = df.iloc[3:]
    df.drop(columns=[0], inplace=True)
    
    df.columns = df.iloc[0]
    df = df.iloc[1:-1]

    df = df.loc[df['Mês/Ano Protocolo'].str.contains('2025')]
    
    df = df.groupby('Mês/Ano Protocolo')['Valor comprometido R$'].sum().reset_index()

    return df  # Retorna o dataframe processado para análise

def process_credito_rural_dashboard(path='./dados_teste/Credito_Rural.xlsx') -> dict:
    """
    Processa dados de Crédito Rural para criar dashboards com:
    - Gráfico de valores por mês (últimos 12 meses)
    - Gráfico de valores por ano
    """
    df = pd.read_excel(path)
    
    # Filtrar dados válidos (remover NaN)
    df_clean = df.dropna(subset=['Valor'])
    
    # Dados por ano (valores anuais)
    df_anos = df_clean[df_clean['Mês'].isna()].copy()
    df_anos = df_anos.sort_values('Ano')
    
    # Dados por mês (últimos 12 meses disponíveis)
    df_meses = df_clean[df_clean['Mês'].notna()].copy()
    df_meses = df_meses.sort_values(['Ano', 'Nº Mês'])
    
    # Pegar os últimos 12 meses
    df_meses_12 = df_meses.tail(12)
    
    # Criar labels para os meses (Mês/Ano)
    df_meses_12['Label'] = df_meses_12['Mês'] + '/' + df_meses_12['Ano'].astype(str)
    
    # Dados para gráficos
    dados_anos = {
        'anos': df_anos['Ano'].tolist(),
        'valores': df_anos['Valor'].tolist(),
        'variacoes': df_anos['Variação'].fillna(0).tolist()
    }
    
    dados_meses = {
        'labels': df_meses_12['Label'].tolist(),
        'valores': df_meses_12['Valor'].tolist(),
        'variacoes_mes': df_meses_12['Variação Mês'].fillna(0).tolist(),
        'meses': df_meses_12['Mês'].tolist(),
        'anos': df_meses_12['Ano'].tolist()
    }
    
    # Estatísticas gerais
    total_registros = len(df_clean)
    valor_medio_anual = df_anos['Valor'].mean() if len(df_anos) > 0 else 0
    valor_medio_mensal = df_meses['Valor'].mean() if len(df_meses) > 0 else 0
    
    return {
        'dados_anos': dados_anos,
        'dados_meses': dados_meses,
        'estatisticas': {
            'total_registros': total_registros,
            'valor_medio_anual': valor_medio_anual,
            'valor_medio_mensal': valor_medio_mensal,
            'anos_disponiveis': len(dados_anos['anos']),
            'meses_disponiveis': len(dados_meses['labels'])
        }
    }
# # df = pd.read_excel('./geral/Fertilizantes_Fungicida.xls')

def catch_podas(path) -> dict:
    df = pd.read_excel(path)
    df['produto'] =df['produto'].str.lower()
    df_pivot = df.pivot_table(index='produto', columns='variavel', values='dado', aggfunc='first').reset_index()
    
    df_pivot['categoria'] = df_pivot['produto'].apply(tentativa_fuzzy)


    produtos_pendentes = df_pivot[df_pivot['categoria'] == 'Pendente']['produto'].tolist()


    if produtos_pendentes:
   
        # Chamada única à API
        resultados_llm = categorizar_em_lote_com_llm(produtos_pendentes)
        
        # 5. Mapeie os resultados de volta para o DataFrame
        # O .map() é perfeito para isso
        df_pivot['categoria'] = df_pivot.apply(
            lambda row: resultados_llm.get(row['produto'], row['categoria']) if row['categoria'] == 'Pendente' else row['categoria'],
            axis=1
        )


    df = df_pivot.loc[df_pivot['categoria'] != 'Não Categorizado']

    df['Quantidade produzida'] = df['Quantidade produzida'].str.replace(' Toneladas', '').str.replace('.', '', regex=False)
    df['Quantidade produzida'] = pd.to_numeric(df['Quantidade produzida'], errors='coerce')
    
    df['Área colhida'] = df['Área colhida'].str.replace(' Hectares', '').str.replace('.', '', regex=False)
    df['Área colhida'] = pd.to_numeric(df['Área colhida'], errors='coerce')

    df['Custo por Hectare'] = pd.to_numeric(df['Custo por Hectare'], errors='coerce')

    df = df.groupby('categoria')[['Custo por Hectare', 'Quantidade produzida', 'Área colhida']].sum()

    return df

def catch_cultives_prices(path) -> dict:
    df = pd.read_excel(path)

    df = pd.read_excel(path, header=None)
    df = df.iloc[12:]
    df.drop(columns=[0], inplace=True)
    
    df.columns = df.iloc[0]
    df = df.iloc[1:-1]

    produtos_antigos_df = pd.read_excel('./dados_teste/Preço Cultivos.xlsx')
    produtos_antigos = produtos_antigos_df['Produto'].unique()
    
    print(df.shape)
    df = df.dropna()
    print(df.shape)
    
    dados = []
    for i, r in df.iterrows():

        product_f_apnd = None
        if isinstance(r['Produto'],str):
            if 'kg' in r['Produto']:
                qtd_kg = r['Produto'].split('(')[-1].replace(')','').replace('kg','')
                if '-' in qtd_kg:
                    qtd_kg = qtd_kg.split('-')[0]
                if qtd_kg == '':
                    qtd_kg = 1
                qtd_kg = float(qtd_kg)
            
            for product in produtos_antigos:
            
                if product.lower() in r['Produto'].lower():
                    product_f_apnd = product

        if isinstance(r['Preço medio'],str):        
            valor = convert_br_number(r['Preço medio'].replace('R$',''))
        
        if not isinstance(r['Período'], float):
            mes, ano = r['Período'].split('/')
            # print(mes, ano)


        valor_real = valor/qtd_kg

        if product_f_apnd:
            dados.append({'Produto':product_f_apnd,
                        'Nivel de comercialização':r['Nível de comercialização'],
                        'UF':r['UF'],
                        'Mês':mes,
                        'Ano':ano,
                        'Preço Médio (R$/Kg)':valor_real})
    return dados
        # print(f'Valor Por kg de {r['Produto']} = {valor_real} | valor antigo: {r['Preço medio']}')

def catch_pib_prices(path) -> dict:
 

    if torch.cuda.is_available():
        print("Sucesso! O PyTorch consegue usar a GPU (CUDA).")
        print(f"Nome da GPU: {torch.cuda.get_device_name(0)}")
    else:
        print("PROBLEMA: O PyTorch NÃO consegue encontrar o CUDA. Esta é a causa do seu problema.")
    source = path # file path or URL
    converter = DocumentConverter()
    doc = converter.convert(source).document
    markdown = doc.export_to_markdown

    pib = retorna_pib_atualizado(markdown)
    return pib

# herbicidas = catch_values_conab('./geral/Agrotoxicos_Herbicida.xls', 'herbicidas')
# print(herbicidas)
# inseticidas = catch_values_conab('./geral/Agrotoxicos_Inseticida.xls', 'inseticidas')
# print(herbicidas)
# herbicidas = catch_values_conab('./geral/Fertilizantes_Fungicida.xls', 'fungicidas')
# print(herbicidas)
# # # # print(herbicidas)
# print(fertilizantes)

# preços_podas = catch_cultives_prices('./geral/Preços Cultivos.xlsx')
# print(preços_podas)

# valores_cana = pega_valores_cana()
# print(valores_cana)
# valores_graos = pega_valores_graos()
# print(valores_graos)
# valores_cafe = pega_valores_cafe()
# print(valores_cafe)

# dolar = catch_dolar_median('./geral/Dolar.csv')
# print(dolar)

# VBP = catch_vbp_values('./geral/VBP.xlsx')
# print(VBP)

# credito_rural = catch_credito_rural_values('./geral/Credito Rural.xlsx')
# print(credito_rural)

# pib = catch_pib_prices('./geral/Pib Agro.pdf')
# print(pib)


