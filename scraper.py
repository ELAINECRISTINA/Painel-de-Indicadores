import os
from pydoc import ispath
import time
import logging
from datetime import datetime, timedelta
import locale
import pyautogui
import pandas as pd

import requests

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException, NoAlertPresentException
from selenium.webdriver.common.keys import Keys


from webdriver_manager.chrome import ChromeDriverManager

from tools.passa_captcha_llm import analisar_imagem_com_gemini, retorna_labels_certos
from tools.conab_dict import conab_values
from tools.treatments_tools import convert_br_number
# 1. Importações necessárias


# ==========================================================
# =========== INÍCIO DAS CONFIGURAÇÕES DO DRIVER ===========
# ==========================================================

# ETAPA 1: Configurar as Opções do Chrome (Options)
# Aqui você define como o navegador deve se comportar.
cookie_consentimento = {
    'name': 'cookie-consent',
    'value': 'true',
    'domain': '.bcb.gov.br'  # Use o domínio que você encontrou
}

chrome_options = Options()
chrome_options.add_argument("--start-maximized")         
chrome_options.add_argument("--disable-notifications")    # Desabilita as notificações do navegador.
chrome_options.add_argument("--disable-gpu")              # Recomendado ao usar o modo headless.
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-web-security")
chrome_options.add_argument("--allow-running-insecure-content")

path_geral = "C:\\Users\\ddemico\\Projetos\\portable_news\\automate\\geral"
prefs = {
    "download.default_directory": path_geral,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,
    "safebrowsing.enabled": False,
    "safebrowsing.disable_download_protection": True,
    "profile.default_content_setting_values.notifications": 2,
    "profile.default_content_settings.popups": 0,
    "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
    "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
    "PluginsAllowedForUrls": "https://www.gov.br",
    "profile.default_content_setting_values.automatic_downloads": 1
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome( options=chrome_options)
wait = WebDriverWait(driver, 20)

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def download():
    try:
        time.sleep(2)

        pyautogui.press('enter')
        print('Download Confirmado')
    except Exception as e:
        print('Ocorreu um erro ao confirmar o download: ',e)

def verify_new_download(new_name):
    global arquivos
    arquivos_new = os.listdir(path_geral)
    print(len(arquivos_new),len(arquivos))
    if not (len(arquivos_new) > len(arquivos)):
        raise FileNotFoundError("Nenhum novo arquivo foi baixado")
    else:
        print("="*20)
        print(arquivos, arquivos_new)
        arquivo_novo = ""
        for i in arquivos_new:
            if i not in arquivos:
                arquivo_novo = i
                break
        print(arquivo_novo)
        extensao = arquivo_novo.split(".")[-1]
        print(extensao)
        print("="*20)
        os.rename(os.path.join(path_geral, arquivo_novo), os.path.join(path_geral, f"{new_name}.{extensao}"))

        arquivos = os.listdir(path_geral)
        
        return print(f"{arquivo_novo} ->{new_name}.xlsx ARRUMADO COM SUCESSO")
        
def get_safras(driver, link, tipo):
    try:
        driver.get(link)
        time.sleep(2)
        links = driver.find_elements(By.XPATH, '//a[contains(text(),"Levantamento")]')
        new_link = links[0].get_attribute("href")
        driver.get(new_link)
        
        tabela_dados = driver.find_element(By.XPATH, '//a[contains(text(),"Tabela de dados")]')
        driver.get(tabela_dados.get_attribute("href"))
        time.sleep(5)
        
    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)
    
def get_VBP(driver,link):
    try:
        driver.get(link)
        time.sleep(2)
        links = driver.find_elements(By.XPATH, '//a[contains(text(),"VBP Brasil")]')
        # print(links)

        link = links[0].get_attribute("href")
        driver.get(link)
        time.sleep(5)

    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)

def get_dolar(driver,link):
    try:
        driver.get(link)

        cookie = driver.find_element(By.XPATH, '//button[contains(text(),"Aceitar")]')
        cookie.click()
        time.sleep(2)
        iframe_locator = (By.XPATH, "//iframe[contains(@src, 'ptax.bcb.gov.br')]")
        wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))

        button = driver.find_element(By.XPATH, '/html/body/div/form/div/input')
        button.click()
        time.sleep(1)

        link_csv = driver.find_element(By.XPATH,'//a[contains(normalize-space(.),"CSV")]').get_attribute("href")
        driver.get(link_csv)
        time.sleep(2)
        


    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)

def get_conab(driver:object, link:str, item:dict):
    try:
        driver.get(link)
        time.sleep(3)
        max_tentativas = 10
        tentativa = 0
        while tentativa < max_tentativas:
                select_grupo = Select(driver.find_element(By.XPATH, '//select[@id="idGrupo"]'))
                select_grupo.select_by_value(item['grupo_value'])
                time.sleep(1)
                select_sub_grupo = Select(driver.find_element(By.XPATH, '//select[@id="idSubGrupo"]'))
                select_sub_grupo.select_by_value(item['sub_grupo_value'])
                time.sleep(1)

                select_ano = Select(driver.find_element(By.XPATH, '//select[@id="anos"]'))
                select_ano_final = Select(driver.find_element(By.XPATH, '//select[@id="anoFinal"]'))
                select_ano.select_by_value('2025')
                select_ano_final.select_by_value('2025')
                time.sleep(1)

                input_captcha = driver.find_element(By.XPATH, '//input[@name="jcaptcha"]')
                consultar_btn = driver.find_element(By.XPATH, '//input[@value="Consultar"]')
                captcha = driver.find_element(By.XPATH, '//img[@src="jcaptcha.jpg"]')
                captcha.screenshot('captcha.png')
                texto = analisar_imagem_com_gemini('captcha.png', 'gemini-2.5-flash')
                print(texto)
                os.remove('captcha.png')
                input_captcha.send_keys(texto)
                consultar_btn.click()
                time.sleep(1)

                try:
                    excel = driver.find_element(By.XPATH, '//span[@title="Excel"]')
                    excel.click()
                    break
                except UnexpectedAlertPresentException:
                    try:
                        # Mude o foco para o alerta
                        alert = driver.switch_to.alert

                        # Verifique se o texto do alerta é o esperado
                        if "O código de segurança informado é inválido." in alert.text:
                            print(f"Alerta de segurança inválido detectado: '{alert.text}'. Ignorando.")
                            # Aceite o alerta para fechá-lo
                            alert.accept()
                            tentativa += 1 
                            continue
                        else:
                            # Se for um alerta diferente, você pode querer tratá-lo de outra forma
                            print(f"Alerta inesperado encontrado: '{alert.text}'.")
                            alert.dismiss() # ou alert.accept() dependendo do caso

                    except NoAlertPresentException:
                        # Caso o alerta já tenha sido fechado por algum motivo
                        print("Nenhum alerta presente para manipular.")

                except Exception as e:
                    print(f"Erro na tentativa {tentativa + 1}: {e}")
                    tentativa += 1
                    time.sleep(2)
                    continue

                time.sleep(5)
        time.sleep(5)

    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)
    
def get_credito_rural(driver,link):

    try:
        driver.get(link)
        time.sleep(2)
        
        link_doc = driver.find_element(By.XPATH, '//a[contains(text(),"Plano Agrícola")]').get_attribute("href")

        driver.get(link_doc)
        time.sleep(2)

    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)

def get_podas_tables(driver, link):
    try:

        #### Primeiro pega os links de todos os produtos
        driver.get(link)
        driver.implicitly_wait(10)

        prods = driver.find_elements(By.XPATH,'//a[@prod]')
        
        # Extrai os atributos dos produtos antes do loop para evitar StaleElementReferenceException
        produtos_info = []
        for prod in prods:
            try:
                produto_nome = prod.get_attribute('prod')
                if produto_nome:
                    produtos_info.append(produto_nome)
            except Exception as e:
                print(f"Erro ao extrair atributo do produto: {e}")
                continue

        
        dados_area_quantidade = []
        for produto_nome in produtos_info:
            info_link = f"https://www.ibge.gov.br/explica/producao-agropecuaria/{produto_nome}/br"

            try:
                driver.get(info_link)
            

                variaveis = driver.find_elements(By.CLASS_NAME,'variavel-nome')
                dados = driver.find_elements(By.CLASS_NAME,'variavel-dado')
                periodo = driver.find_elements(By.CLASS_NAME,'variavel-periodo')

                for var, dado, per in zip(variaveis, dados, periodo):
                    try:
                        var_text = var.text
                        dado_text = dado.text
                        periodo_text = per.text
                        
                        print(var_text, dado_text, periodo_text)
                        if var_text.lower() == 'área colhida' or var_text.lower() == 'quantidade produzida':
                            dados_area_quantidade.append({
                                'produto': produto_nome,
                                'variavel': var_text,
                                'dado': dado_text,
                                'periodo': periodo_text
                            })
                    except Exception as e:
                        print(f"Erro ao processar dados do produto {produto_nome}: {e}")
                        continue
                
                print("Terminou de pegar o", produto_nome)
                
            except Exception as e:
                print(f"Erro ao processar produto {produto_nome}: {e}")
                continue
                
        print(dados_area_quantidade)
        df = pd.DataFrame(dados_area_quantidade)
        print(df)

        # Parte para pegar o Crédito Rural e depois fazer o cruzamento de tudo
        link_2 = 'https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/custos-de-producao/planilhas-de-custos-de-producao'
        driver.get(link_2)

        links_arquivos = driver.find_elements(By.XPATH,'//td/p/a[@href]')
   
        
        dados_poda_path = os.path.join(path_geral, 'dados_poda')
        os.makedirs(dados_poda_path, exist_ok=True)
     


        
        for link in links_arquivos:
            link_arch = link.get_attribute('href')
            arquivos = os.listdir(path_geral)
            
            name = f'{link.text.replace("/","_")}'

            try:
                response = requests.get(link_arch, stream=True, timeout=10)
                if response.status_code == 200:
                    with open(os.path.join(dados_poda_path, name), 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                else:
                    print(f"Erro ao baixar {name}: Status Code {response.status_code}")
            except Exception as e:
                print(f"Erro ao baixar {name}: {e}")

        
        arquivos_baixados = os.listdir(dados_poda_path)      
        dados_custo = []
        for arquivo in arquivos_baixados:
            caminho = os.path.join(dados_poda_path, arquivo)
            # print(arquivo)
            try:
                df = pd.read_excel(caminho, sheet_name=None)
            except:
                continue 
            custo_hectare = 0
            for nome_planilha, df in df.items():
                # print(nome_planilha)
                if nome_planilha.split('-')[-1] == '2024':
                    new_df = pd.read_excel(caminho, sheet_name=nome_planilha)
                    # print(new_df)

                    
                    for i,r in new_df.iterrows():
                        vlr = r.iloc[0]
                        vlr_segunda_coluna = r.iloc[1]
                        
                        if 'Tratores e Colheitadeiras' in str([vlr]): 
                            # print('Custo por Hectare:', convert_br_number(vlr_segunda_coluna))
                            custo_hectare += convert_br_number(vlr_segunda_coluna)
                     
            
            if custo_hectare:
                print(f'Cultura: {arquivo} | Custo por Hectare {custo_hectare}')
                dados_area_quantidade.append({
                                'produto': arquivo,
                                'variavel': 'Custo por Hectare',
                                'dado': custo_hectare,
                                'periodo': '2024'
                            })

           
           
            print("+"*22)
    except Exception as e:
        print('Ocorreu um erro na execução do script: ',e)

    dados_area_quantidade_df = pd.DataFrame(dados_area_quantidade)
    dados_area_quantidade_df.to_excel('./geral/Dados_para_poda.xlsx')
    dados_baixados = os.listdir(dados_poda_path)
    for i in dados_baixados:
        os.remove(os.path.join(dados_poda_path, i))
    os.rmdir(dados_poda_path)

def get_cultures_prices(driver, link):
    
    driver.get(link)
    
    time.sleep(2)

    _button = driver.find_element(By.XPATH,'//*[@id="main-content"]/app-home/form/div[1]/nav/ul/li[2]/button').click()
    wait = WebDriverWait(driver, 10)
    time.sleep(5)
    

    print("Abrindo o seletor de mês...")
    dropdown_trigger = driver.find_element(By.ID, "mesInicial")
    dropdown_trigger.click()
    janeiro_label = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//label[text()='Janeiro']"))
    )
    janeiro_label.click()
    
    dropdown_trigger = driver.find_element(By.ID, "anoInicial")
    dropdown_trigger.click()
    ano_inicial_select = str(datetime.now().year)
    ano_inicial = wait.until(
        EC.element_to_be_clickable((By.XPATH,f"//label[text()='{ano_inicial_select}']"))
    )
    ano_inicial.click()


   
    # hoje = datetime.now()
    # _primeiro_dia = hoje.replace(day=1)
    # _data_mes_anterior = _primeiro_dia - timedelta(days=1)
    # mes_anterior = _data_mes_anterior.strftime('%B').capitalize()
    # print(mes_anterior)

    mes = datetime.now().month - 2
    str_mes = f'mesFinal{mes}'

    dropdown_trigger = driver.find_element(By.ID, "mesFinal")
    dropdown_trigger.click()
    final_mes_label = wait.until(
        EC.element_to_be_clickable((By.XPATH, f"//label[@for='{str_mes}']"))
    )
    final_mes_label.click()
    
    dropdown_trigger = driver.find_element(By.ID, "anoFinal")
    dropdown_trigger.click()
    ano_inicial_select = str(datetime.now().year)
    ano_inicial = wait.until(
        EC.element_to_be_clickable((By.XPATH,f"//label[@for='anoFinal0']"))
    )
    ano_inicial.click()
    
    _button_download = driver.find_element(By.XPATH,'//*[@id="main-content"]/app-home/form/div[2]/div[1]/div[3]/button').click()

    time.sleep(2)

    products_list = driver.find_element(By.ID, "produto")
    products_dropdown = products_list.find_element(By.XPATH,'.//div[@tabindex]')
    driver.execute_script("arguments[0].setAttribute('expanded', '');", products_dropdown)

    all_products = driver.find_elements(By.XPATH,"//label[starts-with(@for,'produto')]")

    produtos = []
    for product in all_products:
        produtos.append({'produto':product.text,'produto_id':product.get_attribute('for')})

    labels = retorna_labels_certos(produtos)    
    for lb in labels:
        _checkbox = wait.until(EC.element_to_be_clickable((By.XPATH,f"//label[@for='{lb}']")))
        driver.execute_script("arguments[0].click();", _checkbox)
        time.sleep(1)
        
    driver.execute_script("arguments[0].removeAttribute('expanded', '');", products_dropdown)
    
    time.sleep(1)

    comercialization_list = driver.find_element(By.ID, "nivelComercializacao")
    comercialization_dropdown = comercialization_list.find_element(By.XPATH,'.//div[@tabindex]')
    driver.execute_script("arguments[0].setAttribute('expanded', '');", comercialization_dropdown)
    _all_comercialization = driver.find_element(By.XPATH,'//label[@for="nivelComercializacao-all"]').click()
    driver.execute_script("arguments[0].removeAttribute('expanded', '');", comercialization_dropdown)

    time.sleep(1)

    location_dropdown = driver.find_element(By.ID,'unidadeFederacao').find_element(By.XPATH,'.//div[@tabindex]')
    driver.execute_script("arguments[0].setAttribute('expanded', '');", location_dropdown)
    _all_locals = driver.find_element(By.XPATH,"//label[@for='unidadeFederacao-all']").click()
    driver.execute_script("arguments[0].removeAttribute('expanded', '');", location_dropdown)
    
    button_consult = driver.find_element(By.XPATH,'//button[text()=" Consultar "]').click()
    time.sleep(1)
    button_download = driver.find_element(By.XPATH,'//button[@aria-label="Excel"]').click()
    time.sleep(5)
  

def get_pib_values(driver, link):
    driver.get(link)
    time.sleep(2)
    archive = driver.find_element(By.XPATH,'//*[@id="imagenet-content"]/div[2]/div[1]/div/div/div[1]/div[1]/p[9]/a').click()
    # driver.get(archive.get_attribute('href'))
    time.sleep(2)
        


    

link_graos = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-graos"
link_cana = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cana-de-acucar"
link_cafe = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cafe"
link_VBP = "https://www.gov.br/agricultura/pt-br/assuntos/politica-agricola/valor-bruto-da-producao-agropecuaria-vbp"
link_dolar = 'https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes'
link_conab = 'https://consultaweb.conab.gov.br/consultas/consultaInsumo.do?method=acaoCarregarConsulta'
link_credito_rural = 'https://www.bndes.gov.br/wps/portal/site/home/transparencia/consulta-operacoes-bndes/credito-rural-desempenho-operacional'
link_podas = 'https://www.ibge.gov.br/explica/producao-agropecuaria/br'
link_precos_cultivos = 'https://consultaprecosdemercado.conab.gov.br/#/home'
link_pib_agro = 'https://www.cepea.org.br/br/pib-do-agronegocio-brasileiro.aspx'





try:
    archs = os.listdir(path_geral)
    if len(archs) > 0:
        for i in archs:
            if not ispath(f"{path_geral}/{i}"):
                complete_path = f"{path_geral}/{i}"
                print(complete_path)
                os.remove(complete_path)

    
    arquivos = os.listdir(path_geral)

    tempo_inicial = datetime.now()
    # for item in conab_values:
    #     get_conab(driver, link_conab, item)
    #     time.sleep(2)
    #     verify_new_download(f"{item['grupo'].title()}_{item['sub_grupo'].title()}")
    # get_safras(driver, link_cafe, "café") 
    # verify_new_download("Café")

    # get_safras(driver, link_graos, "graos") 
    # verify_new_download("Graos")

    # get_safras(driver, link_cana, "cana") 
    # verify_new_download("Cana")

    # get_VBP(driver, link_VBP)
    # verify_new_download("VBP")

    # get_credito_rural(driver, link_credito_rural)
    # verify_new_download("Credito Rural")

    # get_cultures_prices(driver, link_precos_cultivos)
    # verify_new_download("Preços Cultivos")
    # get_podas_tables(driver, link_podas)

    get_pib_values(driver, link_pib_agro)
    verify_new_download('Pib Agro')
    
    

 
    

except FileNotFoundError as e:
    print("Um dos arquivos não foi Baixado") 

except Exception as e:
    print('Ocorreu um erro na execução do script: ',e)
# finally:
#     tempo_final = datetime.now()
#     print("Tempo de Execução: ", tempo_final - tempo_inicial)
#     driver.quit() 