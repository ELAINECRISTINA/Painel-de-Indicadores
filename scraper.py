import os
import time
import shutil
import logging
from datetime import datetime, timedelta
import locale
import pyautogui
import pandas as pd
import openpyxl

import requests

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException, NoAlertPresentException, WebDriverException, InvalidSessionIdException
from selenium.webdriver.common.keys import Keys

from webdriver_manager.chrome import ChromeDriverManager

from tools.passa_captcha_llm import analisar_imagem_com_gemini, retorna_labels_certos
from tools.conab_dict import conab_values
from tools.treatments_tools import convert_br_number


# ==========================================================
# =========== INÍCIO DAS CONFIGURAÇÕES DO DRIVER ===========
# ==========================================================

cookie_consentimento = {
    'name': 'cookie-consent',
    'value': 'true',
    'domain': '.bcb.gov.br'
}

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-web-security")
chrome_options.add_argument("--allow-running-insecure-content")
# chrome_options.add_argument("--headless")

chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36")

# ==========================================================
# =================== PROTEÇÃO DE CAMINHOS ==================
# ==========================================================
# 🛑 FIX CRÍTICO: path_geral é a raiz do projeto (onde ficam os .py).
# Antes, o Chrome baixava arquivos DIRETO nessa pasta e um bloco de
# limpeza no início do script apagava "tudo que estivesse lá" — o que
# incluía o próprio código-fonte (scraper.py, app.py, etc).
#
# Agora: path_geral continua sendo a raiz do projeto (só leitura/referência),
# e downloads_path é uma subpasta EXCLUSIVA para arquivos baixados pelo
# Chrome. Nenhuma rotina de limpeza/remoção deve operar em path_geral.
path_geral = r"C:\Users\ESILVA65\OneDrive - JACTO\Documentos\automate"
downloads_path = os.path.join(path_geral, '_downloads_temp')
os.makedirs(downloads_path, exist_ok=True)

dados_teste_path = os.path.join(path_geral, 'dados_teste')
os.makedirs(dados_teste_path, exist_ok=True)

# Extensões que NUNCA podem ser removidas por rotina automática,
# mesmo que algum caminho seja configurado incorretamente no futuro.
EXTENSOES_PROIBIDAS = {'.py', '.json', '.env', '.cfg', '.ini', '.toml', '.yaml', '.yml'}


def remove_seguro(caminho):
    """
    Wrapper de os.remove que recusa apagar arquivos de código/configuração.
    Última linha de defesa: mesmo que downloads_path/dados_teste_path um dia
    voltem a apontar para a pasta errada, isso evita apagar o projeto.
    """
    _, ext = os.path.splitext(caminho)
    if ext.lower() in EXTENSOES_PROIBIDAS:
        raise RuntimeError(f"🛑 Tentativa de remover arquivo protegido, abortando: {caminho}")
    os.remove(caminho)


prefs = {
    "download.default_directory": downloads_path,  # 🛑 FIX: era path_geral
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

driver = webdriver.Chrome(options=chrome_options)
driver.set_page_load_timeout(30)
wait = WebDriverWait(driver, 20)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')


# ==========================================================
# =================== FUNÇÕES AUXILIARES ===================
# ==========================================================

def download():
    try:
        time.sleep(2)
        pyautogui.press('enter')
        print('Download Confirmado')
    except Exception as e:
        print('Ocorreu um erro ao confirmar o download: ', e)


def verify_new_download(new_name):
    """
    Detecta o arquivo recém-baixado em downloads_path comparando a lista de
    arquivos antes/depois. Quando há mais de um arquivo novo (ex: cópias
    numeradas '(1)', '(2)' deixadas por execuções anteriores que não foram
    limpas), escolhe o de modificação mais recente em vez do primeiro da
    lista, já que a ordem de os.listdir não é garantida.

    🛑 FIX: antes a verificação usava `len(arquivos_new) > len(arquivos)`
    como guarda de entrada. Isso falha sempre que um arquivo sai da pasta
    (ex: move_to_dados_teste movendo o download anterior) na mesma janela
    em que outro entra: a contagem líquida fica igual (ex: 2 e 2) mesmo
    com um arquivo genuinamente novo presente, e o erro
    "Nenhum novo arquivo foi baixado" disparava com o arquivo já ali,
    intacto, como aconteceu com a Cana. A checagem correta é olhar
    diretamente a diferença de conjuntos (`novos`), não o tamanho das
    listas.
    """
    global arquivos
    arquivos_new = os.listdir(downloads_path)  # 🛑 FIX: era path_geral
    print(len(arquivos_new), len(arquivos))

    novos = [i for i in arquivos_new if i not in arquivos]

    if not novos:
        raise FileNotFoundError("Nenhum novo arquivo foi baixado")
    else:
        print("=" * 20)
        print(f"Arquivos novos detectados: {novos}")

        # Escolhe o mais recente entre os novos, em caso de mais de um
        arquivo_novo = max(
            novos,
            key=lambda f: os.path.getmtime(os.path.join(downloads_path, f))
        )
        print(f"Arquivo novo escolhido: {arquivo_novo}")
        extensao = arquivo_novo.split(".")[-1]
        print(extensao)
        print("=" * 20)

        destino = os.path.join(downloads_path, f"{new_name}.{extensao}")

        if os.path.exists(destino):
            remove_seguro(destino)
            print(f"🗑️ Arquivo antigo removido: {new_name}.{extensao}")

        os.rename(os.path.join(downloads_path, arquivo_novo), destino)
        arquivos = os.listdir(downloads_path)
        print(f"{arquivo_novo} -> {new_name}.{extensao} ARRUMADO COM SUCESSO")
        return f"{new_name}.{extensao}"


def aguarda_arquivo_especifico(pasta, nome_arquivo, timeout=60, intervalo=1):
    """
    Aguarda até que um arquivo com nome específico apareça completo na pasta.
    Usada quando o site sempre gera o download com o mesmo nome fixo
    (ex: 'Consulta-precos-mensal.xlsx'), tornando a detecção mais confiável
    do que comparar listas de arquivos antes/depois.
    """
    destino = os.path.join(pasta, nome_arquivo)
    temp_destino = destino + '.crdownload'

    tempo_decorrido = 0
    while tempo_decorrido < timeout:
        if os.path.exists(destino) and not os.path.exists(temp_destino):
            # Confirma que o arquivo não está mais sendo escrito
            # (tamanho estável por um curto intervalo)
            tamanho_antes = os.path.getsize(destino)
            time.sleep(0.5)
            if os.path.exists(destino) and os.path.getsize(destino) == tamanho_antes:
                print(f"✅ Arquivo '{nome_arquivo}' detectado e completo após {tempo_decorrido}s")
                return True
        time.sleep(intervalo)
        tempo_decorrido += intervalo

    print(f"⏱️ Timeout de {timeout}s atingido, '{nome_arquivo}' não apareceu")
    return False


def aguarda_download_completo(pasta, arquivos_antes, timeout=60, intervalo=1,
                                estabilidade=1.0):
    """
    Espera ativamente um novo download terminar, sem depender de nome fixo
    nem de time.sleep() arbitrário. Substitui os antigos `time.sleep(5)` ao
    final de get_safras/get_VBP/get_credito_rural/get_pib_values, que
    falhavam quando o site demorava mais que o tempo fixo (caso real: o
    arquivo de cana, mais lento, ficava incompleto e a próxima verificação
    via `verify_new_download` não detectava nada de novo).

    Critérios para considerar o download concluído:
      1. Existe pelo menos um arquivo novo em `pasta` (que não estava em
         `arquivos_antes`).
      2. Nenhum arquivo novo termina em `.crdownload` (indicador nativo do
         Chrome de download em andamento).
      3. O tamanho do arquivo novo mais recente está estável por um
         intervalo curto (`estabilidade` segundos) — proteção extra para
         servidores que finalizam o nome antes de realmente terminar de
         escrever todos os bytes.

    Retorna True se detectou e confirmou o download; False se atingiu o
    timeout sem sucesso (quem chamar decide se trata como erro fatal).
    """
    tempo_decorrido = 0
    while tempo_decorrido < timeout:
        arquivos_atuais = os.listdir(pasta)
        novos = [f for f in arquivos_atuais if f not in arquivos_antes]
        novos_completos = [f for f in novos if not f.endswith('.crdownload')]

        if novos_completos:
            candidato = max(
                novos_completos,
                key=lambda f: os.path.getmtime(os.path.join(pasta, f))
            )
            caminho_candidato = os.path.join(pasta, candidato)
            tamanho_antes = os.path.getsize(caminho_candidato)
            time.sleep(estabilidade)

            # Reconfirma que ainda existe e não apareceu um .crdownload novo
            # para o mesmo arquivo nesse intervalo (alguns sites recriam o
            # arquivo final só depois de finalizar o .crdownload).
            ainda_em_download = any(
                f.endswith('.crdownload') for f in os.listdir(pasta)
                if f not in arquivos_antes
            )
            if not ainda_em_download and os.path.exists(caminho_candidato):
                tamanho_depois = os.path.getsize(caminho_candidato)
                if tamanho_depois == tamanho_antes and tamanho_depois > 0:
                    print(f"✅ Download concluído e estável: {candidato} ({tamanho_depois} bytes) após {tempo_decorrido}s")
                    return True

        time.sleep(intervalo)
        tempo_decorrido += intervalo

    print(f"⏱️ Timeout de {timeout}s atingido aguardando download completar em {pasta}")
    return False


def limpa_arquivo_antigo(pasta, nome_arquivo):
    """
    Remove o arquivo (e variantes numeradas tipo 'nome (1).xlsx') antes de
    um novo download, evitando que o Chrome crie uma cópia renomeada por já
    existir um arquivo com aquele nome, e evitando falsos positivos na
    detecção do novo download.
    """
    destino = os.path.join(pasta, nome_arquivo)
    if os.path.exists(destino):
        remove_seguro(destino)
        print(f"🗑️ Removido antes do novo download: {nome_arquivo}")

    nome_base, ext = os.path.splitext(nome_arquivo)
    for f in os.listdir(pasta):
        if f.startswith(nome_base) and f.endswith(ext) and f != nome_arquivo:
            try:
                remove_seguro(os.path.join(pasta, f))
                print(f"🗑️ Variante antiga removida: {f}")
            except Exception as e:
                print(f"⚠️ Não foi possível remover {f}: {e}")


def move_to_dados_teste(nome_origem, nome_destino=None):
    """Move (e opcionalmente renomeia) arquivo de downloads_path para dados_teste/"""
    global arquivos
    nome_destino = nome_destino or nome_origem
    origem = os.path.join(downloads_path, nome_origem)  # 🛑 FIX: era path_geral
    destino = os.path.join(dados_teste_path, nome_destino)

    if not os.path.exists(origem):
        print(f"❌ Arquivo origem não encontrado: {origem}")
        return

    if os.path.exists(destino):
        remove_seguro(destino)
        print(f"🗑️ Arquivo antigo removido em dados_teste: {nome_destino}")

    shutil.move(origem, destino)

    # 🛑 FIX: resincroniza o snapshot global usado por verify_new_download.
    # Sem isso, o arquivo recém-movido (que saiu de downloads_path) ficava
    # "fantasma" na lista `arquivos`, e a checagem antiga baseada em
    # comparação de tamanho de listas mascarava downloads novos legítimos.
    arquivos = os.listdir(downloads_path)

    print(f"✅ Movido: {nome_origem} → dados_teste/{nome_destino}")


# ==========================================================
# =================== FUNÇÕES DE SCRAPING ==================
# ==========================================================

def get_safras(driver, link, tipo):
    try:
        wait_local = WebDriverWait(driver, 60)
        print(f"🔍 Acessando: {link}")

        try:
            driver.get(link)
        except TimeoutException:
            print(f"⚠️ Timeout no carregamento, continuando...")
            driver.execute_script("window.stop();")
            time.sleep(3)

        print(f"URL atual: {driver.current_url}")
        print(f"Título: {driver.title}")

        try:
            links = wait_local.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//a[contains(translate(text(),"LEVANTAMENTO","levantamento"),"levantamento")]')
                )
            )
        except TimeoutException:
            links = driver.find_elements(By.XPATH, '//a[contains(@href,"levantamento")]')

        if not links:
            print(f"❌ Nenhum link encontrado")
            return

        print(f"✅ {len(links)} links encontrados, acessando o primeiro...")
        new_link = links[0].get_attribute("href")
        print(f"Link: {new_link}")

        try:
            driver.get(new_link)
        except TimeoutException:
            print(f"⚠️ Timeout na segunda página, continuando...")
            driver.execute_script("window.stop();")
            time.sleep(3)

        print(f"URL segunda página: {driver.current_url}")
        driver.save_screenshot(f'safra_{tipo}_2.png')

        try:
            tabela_dados = wait_local.until(
                EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(text(),"Tabela de dados")]')
                )
            )
        except TimeoutException:
            print("⚠️ Tentando XPath alternativo para tabela...")
            tabela_dados = driver.find_element(By.PARTIAL_LINK_TEXT, "Tabela")

        print(f"📊 Tabela encontrada!")
        tabela_link = tabela_dados.get_attribute("href")
        print(f"Link tabela: {tabela_link}")

        # 🛑 FIX: snapshot ANTES do clique que dispara o download, para a
        # espera ativa saber o que é "novo". O antigo time.sleep(5) fixo
        # falhava em arquivos maiores/mais lentos (ex: cana-de-açúcar),
        # deixando o download incompleto e quebrando verify_new_download
        # na chamada seguinte.
        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(tabela_link)
        except TimeoutException:
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(downloads_path, arquivos_antes_download, timeout=90)
        if not sucesso:
            print(f"⚠️ Download de {tipo} não confirmado dentro do timeout")
            driver.save_screenshot(f'safra_{tipo}_download_timeout.png')

        print(f"✅ Download concluído para {tipo}")

    except TimeoutException as e:
        print(f"⏱️ Timeout em get_safras({tipo})")
        driver.save_screenshot(f'safra_{tipo}_timeout.png')
    except Exception as e:
        print(f'❌ Erro em get_safras({tipo}): {e}')
        driver.save_screenshot(f'safra_{tipo}_erro.png')


def get_VBP(driver, link):
    try:
        wait_local = WebDriverWait(driver, 60)
        print(f"🔍 Acessando VBP: {link}")

        try:
            driver.get(link)
        except TimeoutException:
            print(f"⚠️ Timeout no carregamento VBP, continuando...")
            driver.execute_script("window.stop();")
            time.sleep(3)

        print(f"URL atual: {driver.current_url}")
        print(f"Título: {driver.title}")
        driver.save_screenshot('vbp_1.png')

        try:
            links = wait_local.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//a[contains(text(),"VBP Brasil")]')
                )
            )
        except TimeoutException:
            print("⚠️ Tentando XPath alternativo para VBP...")
            links = driver.find_elements(By.PARTIAL_LINK_TEXT, "VBP")

        if not links:
            print("❌ Link VBP não encontrado")
            driver.save_screenshot('vbp_sem_links.png')
            return

        link_vbp = links[0].get_attribute("href")
        print(f"Link VBP encontrado: {link_vbp}")

        # 🛑 FIX: snapshot antes do clique que dispara o download
        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(link_vbp)
        except TimeoutException:
            print(f"⚠️ Timeout ao acessar VBP, continuando...")
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(downloads_path, arquivos_antes_download, timeout=90)
        if not sucesso:
            print("⚠️ Download de VBP não confirmado dentro do timeout")
            driver.save_screenshot('vbp_download_timeout.png')

        print("✅ VBP concluído")

    except Exception as e:
        print(f'❌ Erro em get_VBP: {e}')
        driver.save_screenshot('vbp_erro.png')


def get_dolar(driver, link):
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

        link_csv = driver.find_element(By.XPATH, '//a[contains(normalize-space(.),"CSV")]').get_attribute("href")
        driver.get(link_csv)
        time.sleep(2)

    except Exception as e:
        print('Ocorreu um erro na execução do script: ', e)


def get_credito_rural(driver, link):
    try:
        wait_local = WebDriverWait(driver, 60)
        print(f"🔍 Acessando Crédito Rural: {link}")

        try:
            driver.get(link)
        except TimeoutException:
            print(f"⚠️ Timeout no carregamento, continuando...")
            driver.execute_script("window.stop();")
            time.sleep(3)

        print(f"URL atual: {driver.current_url}")
        print(f"Título: {driver.title}")
        driver.save_screenshot('credito_rural_1.png')

        try:
            link_doc = wait_local.until(
                EC.presence_of_element_located(
                    (By.XPATH, '//a[contains(text(),"Plano Agrícola")]')
                )
            )
        except TimeoutException:
            print("⚠️ Tentando XPath alternativo...")
            link_doc = driver.find_element(By.PARTIAL_LINK_TEXT, "Plano")

        href = link_doc.get_attribute("href")
        print(f"Link encontrado: {href}")

        # 🛑 FIX: snapshot antes do clique que dispara o download
        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(href)
        except TimeoutException:
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(downloads_path, arquivos_antes_download, timeout=90)
        if not sucesso:
            print("⚠️ Download de Crédito Rural não confirmado dentro do timeout")
            driver.save_screenshot('credito_rural_download_timeout.png')

        print("✅ Crédito Rural concluído")

    except Exception as e:
        print(f'❌ Erro em get_credito_rural: {e}')
        driver.save_screenshot('credito_rural_erro.png')


def get_podas_tables(driver, link):
    try:
        driver.get(link)
        driver.implicitly_wait(10)

        prods = driver.find_elements(By.XPATH, '//a[@prod]')

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

                variaveis = driver.find_elements(By.CLASS_NAME, 'variavel-nome')
                dados = driver.find_elements(By.CLASS_NAME, 'variavel-dado')
                periodo = driver.find_elements(By.CLASS_NAME, 'variavel-periodo')

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

        link_2 = 'https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/custos-de-producao/planilhas-de-custos-de-producao'
        driver.get(link_2)

        links_arquivos = driver.find_elements(By.XPATH, '//td/p/a[@href]')

        # 🛑 FIX: pasta temporária de poda criada dentro de downloads_path,
        # nunca em path_geral (raiz do projeto).
        dados_poda_path = os.path.join(downloads_path, 'dados_poda')
        os.makedirs(dados_poda_path, exist_ok=True)

        for link in links_arquivos:
            link_arch = link.get_attribute('href')
            arquivos = os.listdir(downloads_path)

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
            try:
                df = pd.read_excel(caminho, sheet_name=None)
            except:
                continue
            custo_hectare = 0
            for nome_planilha, df in df.items():
                if nome_planilha.split('-')[-1] == '2024':
                    new_df = pd.read_excel(caminho, sheet_name=nome_planilha)
                    for i, r in new_df.iterrows():
                        vlr = r.iloc[0]
                        vlr_segunda_coluna = r.iloc[1]
                        if 'Tratores e Colheitadeiras' in str([vlr]):
                            custo_hectare += convert_br_number(vlr_segunda_coluna)

            if custo_hectare:
                print(f'Cultura: {arquivo} | Custo por Hectare {custo_hectare}')
                dados_area_quantidade.append({
                    'produto': arquivo,
                    'variavel': 'Custo por Hectare',
                    'dado': custo_hectare,
                    'periodo': '2024'
                })

            print("+" * 22)

    except Exception as e:
        print('Ocorreu um erro na execução do script: ', e)

    dados_area_quantidade_df = pd.DataFrame(dados_area_quantidade)
    destino_poda = os.path.join(dados_teste_path, 'Indicadores_Poda.xlsx')
    if os.path.exists(destino_poda):
        remove_seguro(destino_poda)
    dados_area_quantidade_df.to_excel(destino_poda, index=False)
    print(f"✅ Indicadores_Poda.xlsx salvo em dados_teste/")

    dados_baixados = os.listdir(dados_poda_path)
    for i in dados_baixados:
        remove_seguro(os.path.join(dados_poda_path, i))
    os.rmdir(dados_poda_path)


def get_cultures_prices(driver, link):
    global arquivos
    NOME_ARQUIVO_PRECOS = "Consulta-precos-mensal.xlsx"

    try:
        driver.get(link)
        time.sleep(2)

        _button = driver.find_element(By.XPATH, '//*[@id="main-content"]/app-home/form/div[1]/nav/ul/li[2]/button').click()
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
            EC.element_to_be_clickable((By.XPATH, f"//label[text()='{ano_inicial_select}']"))
        )
        ano_inicial.click()
        print("✅ Mês/ano inicial selecionados")

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
            EC.element_to_be_clickable((By.XPATH, f"//label[@for='anoFinal0']"))
        )
        ano_inicial.click()
        print("✅ Mês/ano final selecionados")

        _button_download = driver.find_element(By.XPATH, '//*[@id="main-content"]/app-home/form/div[2]/div[1]/div[3]/button').click()
        time.sleep(2)
        print("✅ Botão intermediário clicado")

        products_list = driver.find_element(By.ID, "produto")
        products_dropdown = products_list.find_element(By.XPATH, './/div[@tabindex]')
        driver.execute_script("arguments[0].setAttribute('expanded', '');", products_dropdown)

        all_products = driver.find_elements(By.XPATH, "//label[starts-with(@for,'produto')]")
        print(f"✅ {len(all_products)} produtos encontrados na lista")

        produtos = []
        for product in all_products:
            produtos.append({'produto': product.text, 'produto_id': product.get_attribute('for')})

        labels = retorna_labels_certos(produtos)
        print(f"✅ Labels retornados pela LLM: {labels}")
        for lb in labels:
            _checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[@for='{lb}']")))
            driver.execute_script("arguments[0].click();", _checkbox)
            time.sleep(1)

        driver.execute_script("arguments[0].removeAttribute('expanded', '');", products_dropdown)
        time.sleep(1)
        print("✅ Produtos selecionados")

        comercialization_list = driver.find_element(By.ID, "nivelComercializacao")
        comercialization_dropdown = comercialization_list.find_element(By.XPATH, './/div[@tabindex]')
        driver.execute_script("arguments[0].setAttribute('expanded', '');", comercialization_dropdown)
        _all_comercialization = driver.find_element(By.XPATH, '//label[@for="nivelComercializacao-all"]').click()
        driver.execute_script("arguments[0].removeAttribute('expanded', '');", comercialization_dropdown)
        time.sleep(1)
        print("✅ Nível de comercialização selecionado")

        location_dropdown = driver.find_element(By.ID, 'unidadeFederacao').find_element(By.XPATH, './/div[@tabindex]')
        driver.execute_script("arguments[0].setAttribute('expanded', '');", location_dropdown)
        _all_locals = driver.find_element(By.XPATH, "//label[@for='unidadeFederacao-all']").click()
        driver.execute_script("arguments[0].removeAttribute('expanded', '');", location_dropdown)
        print("✅ UFs selecionadas")

        limpa_arquivo_antigo(downloads_path, NOME_ARQUIVO_PRECOS)  # 🛑 FIX: era path_geral
        arquivos = os.listdir(downloads_path)

        button_consult = driver.find_element(By.XPATH, '//button[text()=" Consultar "]').click()
        time.sleep(1)
        print("✅ Consulta executada, baixando Excel...")
        button_download = driver.find_element(By.XPATH, '//button[@aria-label="Excel"]').click()
        print("✅ Botão Excel clicado, aguardando download...")

        sucesso = aguarda_arquivo_especifico(downloads_path, NOME_ARQUIVO_PRECOS, timeout=60)  # 🛑 FIX: era path_geral
        if not sucesso:
            driver.save_screenshot('precos_cultivos_timeout.png')
            raise FileNotFoundError(f"Download de '{NOME_ARQUIVO_PRECOS}' não foi concluído a tempo")

        print("✅ Download de Preço Cultivos concluído")

    except Exception as e:
        print(f"❌ Erro em get_cultures_prices: {e}")
        driver.save_screenshot('precos_cultivos_erro.png')
        raise


def pagina_tem_cloudflare_challenge(driver):
    """
    Detecta heuristicamente se a página atual é uma tela de verificação
    anti-bot do Cloudflare (ex: 'Confirme que é humano'), em vez do
    conteúdo real esperado. Usado em get_pib_values porque o domínio do
    CEPEA (cepea.org.br / cepea.esalq.usp.br são espelhos do mesmo site)
    às vezes serve esse desafio antes de liberar a página com o link do
    Excel do PIB.
    """
    try:
        texto = driver.page_source.lower()
    except Exception:
        return False

    indicadores = [
        "verificação de segurança",
        "confirme que é humano",
        "checking your browser",
        "cloudflare",
        "cf-challenge",
        "just a moment",
    ]
    return any(termo in texto for termo in indicadores)


def get_pib_values(driver, link, tentativas=3, espera_challenge=15):
    """
    🛑 FIX: o domínio do CEPEA por vezes redireciona para uma versão
    protegida por Cloudflare ('Confirme que é humano'), o que travava o
    wait_local.until() até dar timeout e quebrar a sessão do ChromeDriver
    (erro de stacktrace gigante observado em produção). Agora a função
    detecta esse desafio e tenta novamente após uma pausa, em vez de
    insistir cegamente num elemento que nunca vai aparecer naquela carga.
    """
    for tentativa in range(1, tentativas + 1):
        try:
            # ✅ FIX: usa wait local para não depender do estado do wait global
            wait_local = WebDriverWait(driver, 30)

            driver.get(link)
            time.sleep(3)

            if pagina_tem_cloudflare_challenge(driver):
                print(f"⚠️ Verificação anti-bot detectada (tentativa {tentativa}/{tentativas}), aguardando {espera_challenge}s...")
                driver.save_screenshot(f'pib_cloudflare_tentativa{tentativa}.png')
                time.sleep(espera_challenge)
                driver.get(link)  # tenta de novo, às vezes resolve sozinho ou cai no domínio sem desafio
                time.sleep(3)

                if pagina_tem_cloudflare_challenge(driver):
                    print(f"⚠️ Verificação anti-bot persistiu na tentativa {tentativa}/{tentativas}")
                    if tentativa < tentativas:
                        continue  # parte para a próxima tentativa do laço externo
                    else:
                        print("❌ Não foi possível passar da verificação anti-bot do CEPEA após todas as tentativas")
                        return

            archive = wait_local.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//a[contains(@href, ".xlsx") or contains(@href, ".xls")]')
                )
            )

            # 🛑 FIX: snapshot antes do clique que dispara o download
            arquivos_antes_download = os.listdir(downloads_path)

            archive.click()

            sucesso = aguarda_download_completo(downloads_path, arquivos_antes_download, timeout=90)
            if sucesso:
                print("✅ PIB clicado")
                return
            else:
                print(f"⚠️ Download de PIB não confirmado dentro do timeout (tentativa {tentativa}/{tentativas})")
                driver.save_screenshot(f'pib_download_timeout_tentativa{tentativa}.png')

        except TimeoutException as e:
            print(f"⏱️ Timeout em get_pib_values (tentativa {tentativa}/{tentativas}): {e}")
            driver.save_screenshot(f'pib_timeout_tentativa{tentativa}.png')
        except (InvalidSessionIdException, WebDriverException) as e:
            # 🛑 A sessão do ChromeDriver morreu (ex: após o stacktrace gigante
            # observado quando o Cloudflare travava a página por muito tempo).
            # Continuar tentando não vai funcionar; melhor abortar com um
            # diagnóstico claro do que insistir em chamadas que vão falhar igual.
            print(f"💀 Sessão do ChromeDriver parece ter morrido em get_pib_values: {e}")
            print("💀 Abortando tentativas de PIB — pode ser necessário reiniciar o driver.")
            return
        except Exception as e:
            print(f'❌ Erro em get_pib_values (tentativa {tentativa}/{tentativas}): {e}')
            driver.save_screenshot(f'pib_erro_tentativa{tentativa}.png')

    print("❌ Todas as tentativas de obter o PIB falharam")


# ==========================================================
# ===================== LINKS ==================================
# ==========================================================

link_graos = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-graos"
link_cana = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cana-de-acucar"
link_cafe = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cafe"
link_VBP = "https://www.gov.br/agricultura/pt-br/assuntos/politica-agricola/valor-bruto-da-producao-agropecuaria-vbp"
link_dolar = 'https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes'
link_credito_rural = 'https://www.bndes.gov.br/wps/portal/site/home/transparencia/consulta-operacoes-bndes/credito-rural-desempenho-operacional'
link_podas = 'https://www.ibge.gov.br/explica/producao-agropecuaria/br'
link_precos_cultivos = 'https://consultaprecosdemercado.conab.gov.br/#/home'
# ✅ FIX: domínio atualizado de cepea.org.br → cepea.esalq.usp.br
link_pib_agro = 'https://www.cepea.esalq.usp.br/br/pib-do-agronegocio-brasileiro.aspx'


# ==========================================================
# ===================== EXECUÇÃO ===========================
# ==========================================================

try:
    # 🛑 FIX: a limpeza inicial agora opera em downloads_path (pasta exclusiva
    # de downloads), nunca mais em path_geral (raiz do projeto/código-fonte).
    archs = os.listdir(downloads_path)
    if len(archs) > 0:
        for i in archs:
            complete_path = os.path.join(downloads_path, i)
            if os.path.isfile(complete_path):
                print(f"🗑️ Removendo arquivo antigo: {complete_path}")
                remove_seguro(complete_path)

    arquivos = os.listdir(downloads_path)
    tempo_inicial = datetime.now()

    get_safras(driver, link_cafe, "cafe")
    verify_new_download("Cafe")
    move_to_dados_teste("Cafe.xlsx")

    get_safras(driver, link_graos, "graos")
    verify_new_download("Graos")
    move_to_dados_teste("Graos.xlsx")

    get_safras(driver, link_cana, "cana")
    verify_new_download("Cana")
    move_to_dados_teste("Cana.xlsx")

    # --- Consolida os 3 em um único Safras.xlsx ---
    safras_origem = {
        "Cafe":  os.path.join(dados_teste_path, "Cafe.xlsx"),
        "Graos": os.path.join(dados_teste_path, "Graos.xlsx"),
        "Cana":  os.path.join(dados_teste_path, "Cana.xlsx"),
    }
    safras_dest = os.path.join(dados_teste_path, "Safras.xlsx")

    if os.path.exists(safras_dest):
        remove_seguro(safras_dest)
        print("🗑️ Safras.xlsx antigo removido")

    with pd.ExcelWriter(safras_dest, engine="openpyxl") as writer:
        nomes_usados = []

        for cultura, caminho in safras_origem.items():
            if not os.path.exists(caminho):
                print(f"⚠️ Arquivo não encontrado, pulando: {caminho}")
                continue

            xl = pd.ExcelFile(caminho, engine="openpyxl")

            for aba in xl.sheet_names:
                df_aba = pd.read_excel(caminho, sheet_name=aba, engine="openpyxl", header=None)

                if df_aba.dropna(how="all").empty:
                    print(f"⏭️ Aba vazia ignorada: [{cultura}] {aba}")
                    continue

                prefixo = cultura[:3].upper()
                nome_base = f"{prefixo} - {aba}"[:31]

                nome_final = nome_base
                contador = 1
                while nome_final in nomes_usados:
                    sufixo = f"_{contador}"
                    nome_final = nome_base[:31 - len(sufixo)] + sufixo
                    contador += 1

                nomes_usados.append(nome_final)
                df_aba.to_excel(writer, sheet_name=nome_final, index=False, header=False)
                print(f"✅ Aba adicionada: {nome_final}")

    print("✅ Safras.xlsx consolidado com sucesso!")

    get_VBP(driver, link_VBP)
    verify_new_download("VBP")
    move_to_dados_teste("VBP.xlsx")

    get_credito_rural(driver, link_credito_rural)
    verify_new_download("Credito Rural")
    move_to_dados_teste("Credito Rural.xlsx", "Credito_Rural.xlsx")

    get_cultures_prices(driver, link_precos_cultivos)
    move_to_dados_teste("Consulta-precos-mensal.xlsx", "Preço Cultivos.xlsx")

    get_pib_values(driver, link_pib_agro)
    # ✅ FIX: captura extensão real do arquivo (pode ser .xls ou .xlsx)
    nome_pib = verify_new_download("PIB")
    move_to_dados_teste(nome_pib, "PIB.xlsx")

    get_dolar(driver, link_dolar)
    # ✅ FIX: captura extensão real do arquivo (é .csv, não .xlsx)
    nome_dolar = verify_new_download("Dolar")
    move_to_dados_teste(nome_dolar, "Dolar.csv")

    # Podas já salva direto em dados_teste/ dentro da própria função
    get_podas_tables(driver, link_podas)

except FileNotFoundError as e:
    print(f"Um dos arquivos não foi Baixado: {e}")
    raise

except Exception as e:
    print('Ocorreu um erro na execução do script: ', e)

# finally:
#     tempo_final = datetime.now()
#     print("Tempo de Execução: ", tempo_final - tempo_inicial)
#     driver.quit()