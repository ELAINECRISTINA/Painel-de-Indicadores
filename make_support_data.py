
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException, NoAlertPresentException
from selenium.webdriver.common.keys import Keys

import pandas as pd
import time
import numpy as np

from webdriver_manager.chrome import ChromeDriverManager

from tools.treatments_tools import convert_br_number

def make_driver():
    chrome_options = Options()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

def add_comma_after_two_digits(value):
    """
    Adiciona vírgula após as duas primeiras casas de um número
    Exemplos: 1000 -> 10,00 | 921312 -> 9213,12 | 50 -> 0,50
    """
    if pd.isna(value) or value == '':
        return value
    
    try:
        # Converte para string e remove espaços
        str_value = str(value).strip()
        
        # Remove qualquer formatação existente (pontos, vírgulas)
        clean_value = str_value.replace('.', '').replace(',', '')
        
        # Verifica se é um número (apenas dígitos)
        if not clean_value.isdigit():
            return value
        
        # Se tem menos de 3 dígitos, adiciona zeros à esquerda
        if len(clean_value) < 3:
            clean_value = clean_value.zfill(3)
        
        # Insere vírgula após as duas primeiras casas
        result = clean_value[:-2] + ',' + clean_value[-2:]
        
        return result
    except:
        return value


def format_last_columns(df):
    """
    Aplica formatação de vírgula nas últimas 4-5 colunas do dataframe
    """
    # Identifica quantas colunas processar (últimas 4 ou 5)
    total_cols = len(df.columns)
    cols_to_process = min(5, total_cols)  # Processa no máximo 5 colunas
    
    # Pega as últimas colunas
    last_columns = df.columns[-cols_to_process:]
    
    for col in last_columns:
        # Aplica a formatação em cada valor da coluna
        formatted_values = []
        for value in df[col]:
            formatted_values.append(add_comma_after_two_digits(value))
        
        df[col] = formatted_values
    
    return df


def make_cultures_data():
    driver = make_driver()
    dataframes = []
    
    try:
        # Navega para a página
        driver.get("https://www.hfbrasil.org.br/br")
        
        # Encontra os links das culturas
        links_culturas = [i.get_attribute('href') for i in driver.find_elements(By.XPATH,'//div[@class="imagenet-box-categoria imagenet-scroll"]/a')]
        print(f"Links encontrados: {len(links_culturas)}")
        
        for link in links_culturas:
            driver.get(link)
            time.sleep(2)

            try:
                driver.find_element(By.ID,'tbl-abrir')
                driver.find_element(By.ID,'tbl-abrir').click()
            except NoSuchElementException:
                pass
            time.sleep(1)
            # Processa as tabelas da página
            if 'uva' in link:
                all_tables = driver.find_elements(By.TAG_NAME, "table")
                if len(all_tables) > 1:
                    # Força pd.read_html a tratar tudo como string primeiro
                    df = pd.read_html(all_tables[1].get_attribute('outerHTML'))[0]
                    # Aplica formatação de vírgula nas últimas colunas
                    df = format_last_columns(df)
                    etiqueta_value = link.split('/')[-1].replace('.aspx','')
                    df['Etiqueta'] = [etiqueta_value] * len(df)
                    dataframes.append(df)
            else:
                table = driver.find_element(By.TAG_NAME, "table")
                # Força pd.read_html a tratar tudo como string primeiro
                df = pd.read_html(table.get_attribute('outerHTML'))[0]
                # Aplica formatação de vírgula nas últimas colunas
                df = format_last_columns(df)
                # Adiciona coluna de etiqueta
                etiqueta_value = link.split('/')[-1].replace('.aspx','')
                df['Etiqueta'] = [etiqueta_value] * len(df)
                print(df.head(5))
                print('='*50)
                dataframes.append(df)
            
            
            time.sleep(2)
            
           
    finally:
        driver.quit()

    final_dataframe = []
    for df in dataframes:
       

        column_index = df.columns.get_loc('Unidade')
        last_column_index = len(df.columns) -1 #Para não pegar a coluna de etiqueta

        df['Preço Médio'] = np.nan


        for idx, r in df.iterrows():  # Mudei 'i' para 'idx' para evitar conflito
            print(r['Unidade']) 
            prices = r.iloc[column_index + 1:last_column_index ].values
            divisor = 1  # Valor padrão para evitar erro
            
            if 'kg' in r['Unidade'].lower():
                print('CAIU AQUI')
                for unit_part in r['Unidade'].split():  # Mudei 'i' para 'unit_part'
                    if unit_part[0].isdigit():
                        if unit_part.lower().endswith('kg'):
                            unit_part = unit_part[:-2]

                        unit_part = convert_br_number(unit_part)
                        print('CAIU NO DIGITO', unit_part)
                        divisor = float(unit_part)  # Converte diretamente para float
                          # Sai do loop após encontrar o primeiro número
                
                valid_prices = [float(convert_br_number(p))/divisor for p in prices if pd.notna(p) and p != '- - -']
                print(valid_prices)
                
            elif 'Alface' in r['Produto']:
                for unit_part in r['Unidade'].split():  # Mudei 'i' para 'unit_part'
                    if unit_part.isdigit():
                        # Corrigido: usar idx (índice da linha) em vez de unit_part
                        df.at[idx, 'Unidade'] = f'{float(unit_part) * 0.3} KG'  # peso médio por pé de alface
                        divisor = float(unit_part) * 0.3
                          # Sai do loop após encontrar o primeiro número
                
                valid_prices = [float(convert_br_number(p))/divisor for p in prices if pd.notna(p) and p != '- - -']
                
            else:
                valid_prices = [float(convert_br_number(p)) for p in prices if pd.notna(p) and p != '- - -']
        
            if valid_prices:
                df.at[idx, 'Preço Médio'] = np.mean(valid_prices)  # Usar idx em vez de i

        median = df['Preço Médio'].median()
        
        if df['Etiqueta'][0] != 'folhosas':
            final_dataframe.append([df['Etiqueta'][0], median])
        else: 
            final_dataframe.append(['alface', median])
    
    colunas = ['Produto','Preço Médio']
    df_final = pd.DataFrame(final_dataframe, columns=colunas)
    return df_final


def make_relation_agrotoxics_cultures():
    # Mudar o Caminho se necessário
    agrotoxics = pd.read_excel(r'C:\Users\ddemico\Projetos\portable_news\automate\dados_teste\Agrotóxicos.xlsx')

    culturas = make_cultures_data()
    print(culturas)    


    list_dicts = []
    cultivos = agrotoxics.groupby('Cultivo')
    for id, df in cultivos:
        if str(id) in [i.title() for i in culturas['Produto'].to_list()]:
            id_price = culturas['Produto'].to_list().index(id.lower())
            price = culturas.iloc[id_price]['Preço Médio']
            list_dicts.append({id:df['Produto'].unique().tolist(), 'Preço Médio':float(price)})
        
    print(list_dicts)
    return list_dicts
    

# make_relation_agrotoxics_cultures()
# make_cultures_data()




    