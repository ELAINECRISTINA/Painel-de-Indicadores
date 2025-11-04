import statistics
import pandas as pd
import tools.treatments_tools as treatments_tools
import string

def add_values_in_months(dict_thing:dict, thing:str, unit, r:pd.Series) -> dict:
    """
        Pega os valores do dicionario 'coisa' e adiciona os preços nos meses
    """
    unit = unit.strip()

    try: 
        unit = float(unit)
    except Exception as E:
        if unit.isalpha():
            if unit == 'T':
                unit = 1000
        else:
            unit = treatments_tools.convert_br_number(unit)

    for i in dict_thing[thing]['precos']:
        temp_ = f"{i} "

      
        if isinstance(unit, str):
            dict_thing[thing]['precos'][i].append(treatments_tools.convert_br_number(r[temp_]))
        elif isinstance(unit, float) or isinstance(unit, int):
            value = treatments_tools.convert_br_number(r[temp_])
            n_f_app = value / unit
            
            dict_thing[thing]['precos'][i].append(n_f_app)    
        else:
            print("CASUcjhABSCNja smcoa vcsaov ainsca ivboo")
            raise ValueError
    return dict_thing

def fix_dict(dict_thing:dict, thing:str) -> dict:
    """
        Pega os valores do dicionario 'coisa' e adiciona a média e mediana
    """
    ids_to_explode = []
    units = dict_thing[thing]['unidade']
    for n, v in enumerate(units):
        if v == 'UN':
            ids_to_explode.append(n)
    

    valores = dict_thing[thing]['precos']
    
    for id in ids_to_explode:
        dict_thing[thing]['unidade'].pop(id)
    for i, k in valores.items():
        if len(k) == 0:
            return dict_thing
        if len(ids_to_explode) > 0:
            for id in ids_to_explode:
                dict_thing[thing]['precos'][i].pop(id)

   
    media = mean(valores)
    mediana = median(valores)
    dict_thing[thing]['média'].append(media)
    dict_thing[thing]['mediana'].append(mediana)
    return dict_thing

def monthly_median(dict_thing: dict, thing: str) -> float:
    
    precos = dict_thing[thing]['precos']
    for i, k in precos.items():
        if len(k) == 0:
            return dict_thing
        else: 
            list_to_statistic = []
            for valor in k:
                if valor > 0:
                    list_to_statistic.append(valor)
            if len(list_to_statistic) > 0:
                dict_thing[thing]['mediana_mensal'][i].append(statistics.median(list_to_statistic))
            else:
                dict_thing[thing]['mediana_mensal'][i].append(0)
    return dict_thing

def mean(dict) -> float:
    """
        Media personalizada
    """
    valores_sem_0 = []
    for k, i in dict.items():
        for valor in i:
            if valor > 0:
                valores_sem_0.append(valor)
    return sum(valores_sem_0) / len(valores_sem_0)

def median(dict) -> float:
    """
        Mediana personalizada
    """
    valores_sem_0 = []
    for k, i in dict.items():
        for valor in i:
            if valor > 0:
                valores_sem_0.append(valor)
    return statistics.median(valores_sem_0)

def indicadores_linhas():
    df = pd.read_excel('./dados_teste/temp_ind_linhas.xlsx')
    prds = df['Produto'].unique()
    
    lista_inseticida = []
    lista_fungicida = []
    lista_herbicida = []
    for i,r in df.iterrows():
        if r['Sub-Grupo'] == 'Inseticida':
            lista_inseticida.append(r['Produto'])
        elif r['Sub-Grupo'] == 'Fungicida':
            lista_fungicida.append(r['Produto'])
        elif r['Sub-Grupo'] == 'Herbicida':
            lista_herbicida.append(r['Produto'])
        
    return list(set(lista_inseticida)), list(set(lista_fungicida)), list(set(lista_herbicida))
