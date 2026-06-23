import os
import time
import shutil
import logging
from datetime import datetime, timedelta
import locale
import pyautogui
import pandas as pd
import openpyxl
import re

import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    UnexpectedAlertPresentException, NoAlertPresentException,
    WebDriverException, InvalidSessionIdException
)
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
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# ==========================================================
# =================== PROTEÇÃO DE CAMINHOS ==================
# ==========================================================
path_geral = r"C:\Users\ESILVA65\OneDrive - JACTO\Documentos\automate"
downloads_path = os.path.join(path_geral, '_downloads_temp')
os.makedirs(downloads_path, exist_ok=True)

dados_teste_path = os.path.join(path_geral, 'dados_teste')
os.makedirs(dados_teste_path, exist_ok=True)

EXTENSOES_PROIBIDAS = {'.py', '.json', '.env', '.cfg', '.ini', '.toml', '.yaml', '.yml'}


def remove_seguro(caminho):
    """
    Wrapper de os.remove que recusa apagar arquivos de código/configuração.
    """
    _, ext = os.path.splitext(caminho)
    if ext.lower() in EXTENSOES_PROIBIDAS:
        raise RuntimeError(
            f"🛑 Tentativa de remover arquivo protegido, abortando: {caminho}"
        )
    os.remove(caminho)


prefs = {
    "download.default_directory": downloads_path,
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
driver.execute_script(
    "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
)

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
    Detecta o arquivo recém-baixado em downloads_path comparando a lista
    de arquivos antes/depois. Escolhe o de modificação mais recente quando
    há mais de um arquivo novo.
    """
    global arquivos
    arquivos_new = os.listdir(downloads_path)
    print(len(arquivos_new), len(arquivos))

    novos = [i for i in arquivos_new if i not in arquivos]

    if not novos:
        raise FileNotFoundError("Nenhum novo arquivo foi baixado")

    print("=" * 20)
    print(f"Arquivos novos detectados: {novos}")

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
    """
    destino = os.path.join(pasta, nome_arquivo)
    temp_destino = destino + '.crdownload'

    tempo_decorrido = 0
    while tempo_decorrido < timeout:
        if os.path.exists(destino) and not os.path.exists(temp_destino):
            tamanho_antes = os.path.getsize(destino)
            time.sleep(0.5)
            if os.path.exists(destino) and os.path.getsize(destino) == tamanho_antes:
                print(
                    f"✅ Arquivo '{nome_arquivo}' detectado e completo "
                    f"após {tempo_decorrido}s"
                )
                return True
        time.sleep(intervalo)
        tempo_decorrido += intervalo

    print(f"⏱️ Timeout de {timeout}s atingido, '{nome_arquivo}' não apareceu")
    return False


def aguarda_download_completo(pasta, arquivos_antes, timeout=60, intervalo=1,
                               estabilidade=1.0):
    """
    Espera ativamente um novo download terminar, sem depender de nome fixo
    nem de time.sleep() arbitrário.
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

            ainda_em_download = any(
                f.endswith('.crdownload') for f in os.listdir(pasta)
                if f not in arquivos_antes
            )
            if not ainda_em_download and os.path.exists(caminho_candidato):
                tamanho_depois = os.path.getsize(caminho_candidato)
                if tamanho_depois == tamanho_antes and tamanho_depois > 0:
                    print(
                        f"✅ Download concluído e estável: {candidato} "
                        f"({tamanho_depois} bytes) após {tempo_decorrido}s"
                    )
                    return True

        time.sleep(intervalo)
        tempo_decorrido += intervalo

    print(
        f"⏱️ Timeout de {timeout}s atingido aguardando download "
        f"completar em {pasta}"
    )
    return False


def limpa_arquivo_antigo(pasta, nome_arquivo):
    """
    Remove o arquivo (e variantes numeradas tipo 'nome (1).xlsx') antes
    de um novo download.
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
    origem = os.path.join(downloads_path, nome_origem)
    destino = os.path.join(dados_teste_path, nome_destino)

    if not os.path.exists(origem):
        print(f"❌ Arquivo origem não encontrado: {origem}")
        return

    if os.path.exists(destino):
        remove_seguro(destino)
        print(f"🗑️ Arquivo antigo removido em dados_teste: {nome_destino}")

    shutil.move(origem, destino)
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
                    (By.XPATH,
                     '//a[contains(translate(text(),"LEVANTAMENTO","levantamento"),'
                     '"levantamento")]')
                )
            )
        except TimeoutException:
            links = driver.find_elements(
                By.XPATH, '//a[contains(@href,"levantamento")]'
            )

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

        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(tabela_link)
        except TimeoutException:
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(
            downloads_path, arquivos_antes_download, timeout=90
        )
        if not sucesso:
            print(f"⚠️ Download de {tipo} não confirmado dentro do timeout")
            driver.save_screenshot(f'safra_{tipo}_download_timeout.png')

        print(f"✅ Download concluído para {tipo}")

    except TimeoutException:
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

        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(link_vbp)
        except TimeoutException:
            print(f"⚠️ Timeout ao acessar VBP, continuando...")
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(
            downloads_path, arquivos_antes_download, timeout=90
        )
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

        cookie = driver.find_element(
            By.XPATH, '//button[contains(text(),"Aceitar")]'
        )
        cookie.click()
        time.sleep(2)

        iframe_locator = (By.XPATH, "//iframe[contains(@src, 'ptax.bcb.gov.br')]")
        wait.until(EC.frame_to_be_available_and_switch_to_it(iframe_locator))

        button = driver.find_element(By.XPATH, '/html/body/div/form/div/input')
        button.click()
        time.sleep(1)

        link_csv = driver.find_element(
            By.XPATH, '//a[contains(normalize-space(.),"CSV")]'
        ).get_attribute("href")
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

        arquivos_antes_download = os.listdir(downloads_path)

        try:
            driver.get(href)
        except TimeoutException:
            driver.execute_script("window.stop();")

        sucesso = aguarda_download_completo(
            downloads_path, arquivos_antes_download, timeout=90
        )
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
            info_link = (
                f"https://www.ibge.gov.br/explica/producao-agropecuaria"
                f"/{produto_nome}/br"
            )

            try:
                driver.get(info_link)

                variaveis = driver.find_elements(By.CLASS_NAME, 'variavel-nome')
                dados    = driver.find_elements(By.CLASS_NAME, 'variavel-dado')
                periodo  = driver.find_elements(By.CLASS_NAME, 'variavel-periodo')

                for var, dado, per in zip(variaveis, dados, periodo):
                    try:
                        var_text     = var.text
                        dado_text    = dado.text
                        periodo_text = per.text

                        print(var_text, dado_text, periodo_text)
                        if var_text.lower() in ('área colhida', 'quantidade produzida'):
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

        link_2 = (
            'https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias'
            '/custos-de-producao/planilhas-de-custos-de-producao'
        )
        driver.get(link_2)

        links_arquivos = driver.find_elements(By.XPATH, '//td/p/a[@href]')

        dados_poda_path = os.path.join(downloads_path, 'dados_poda')
        os.makedirs(dados_poda_path, exist_ok=True)

        for link in links_arquivos:
            link_arch = link.get_attribute('href')
            name = f'{link.text.replace("/","_")}'

            try:
                response = requests.get(link_arch, stream=True, timeout=10)
                if response.status_code == 200:
                    with open(os.path.join(dados_poda_path, name), 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                else:
                    print(
                        f"Erro ao baixar {name}: "
                        f"Status Code {response.status_code}"
                    )
            except Exception as e:
                print(f"Erro ao baixar {name}: {e}")

        arquivos_baixados = os.listdir(dados_poda_path)
        for arquivo in arquivos_baixados:
            caminho = os.path.join(dados_poda_path, arquivo)
            try:
                df_sheets = pd.read_excel(caminho, sheet_name=None)
            except Exception:
                continue

            custo_hectare = 0
            for nome_planilha, df_sheet in df_sheets.items():
                if nome_planilha.split('-')[-1] == '2024':
                    new_df = pd.read_excel(
                        caminho, sheet_name=nome_planilha
                    )
                    for i, r in new_df.iterrows():
                        vlr               = r.iloc[0]
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

        driver.find_element(
            By.XPATH,
            '//*[@id="main-content"]/app-home/form/div[1]/nav/ul/li[2]/button'
        ).click()
        wait_local = WebDriverWait(driver, 10)
        time.sleep(5)

        # ── Mês inicial: Janeiro ──────────────────────────────────────────────
        print("Abrindo o seletor de mês inicial...")
        driver.find_element(By.ID, "mesInicial").click()
        wait_local.until(
            EC.element_to_be_clickable((By.XPATH, "//label[text()='Janeiro']"))
        ).click()

        # ── Ano inicial: 2024 (fixo) ──────────────────────────────────────────
        driver.find_element(By.ID, "anoInicial").click()
        wait_local.until(
            EC.element_to_be_clickable((By.XPATH, "//label[text()='2024']"))
        ).click()
        print("✅ Mês/ano inicial selecionados: Janeiro/2024")

        # ── Mês final: mês atual - 2 ──────────────────────────────────────────
        mes = datetime.now().month - 2
        str_mes = f'mesFinal{mes}'
        driver.find_element(By.ID, "mesFinal").click()
        wait_local.until(
            EC.element_to_be_clickable((By.XPATH, f"//label[@for='{str_mes}']"))
        ).click()

        # ── Ano final: ano atual ──────────────────────────────────────────────
        ano_final_select = str(datetime.now().year)
        driver.find_element(By.ID, "anoFinal").click()
        wait_local.until(
            EC.element_to_be_clickable((By.XPATH, "//label[@for='anoFinal0']"))
        ).click()
        print(f"✅ Mês/ano final selecionados: mês {mes}/{ano_final_select}")

        driver.find_element(
            By.XPATH,
            '//*[@id="main-content"]/app-home/form/div[2]/div[1]/div[3]/button'
        ).click()
        time.sleep(2)
        print("✅ Botão intermediário clicado")

        # ── Seleção de produtos via LLM ───────────────────────────────────────
        products_list    = driver.find_element(By.ID, "produto")
        products_dropdown = products_list.find_element(By.XPATH, './/div[@tabindex]')
        driver.execute_script(
            "arguments[0].setAttribute('expanded', '');", products_dropdown
        )

        all_products = driver.find_elements(
            By.XPATH, "//label[starts-with(@for,'produto')]"
        )
        print(f"✅ {len(all_products)} produtos encontrados na lista")

        produtos = [
            {'produto': p.text, 'produto_id': p.get_attribute('for')}
            for p in all_products
        ]

        labels = retorna_labels_certos(produtos)
        print(f"✅ Labels retornados pela LLM: {labels}")
        for lb in labels:
            _checkbox = wait_local.until(
                EC.element_to_be_clickable((By.XPATH, f"//label[@for='{lb}']"))
            )
            driver.execute_script("arguments[0].click();", _checkbox)
            time.sleep(1)

        driver.execute_script(
            "arguments[0].removeAttribute('expanded', '');", products_dropdown
        )
        time.sleep(1)
        print("✅ Produtos selecionados")

        # ── Nível de comercialização: todos ───────────────────────────────────
        comercialization_list     = driver.find_element(By.ID, "nivelComercializacao")
        comercialization_dropdown = comercialization_list.find_element(
            By.XPATH, './/div[@tabindex]'
        )
        driver.execute_script(
            "arguments[0].setAttribute('expanded', '');", comercialization_dropdown
        )
        driver.find_element(
            By.XPATH, '//label[@for="nivelComercializacao-all"]'
        ).click()
        driver.execute_script(
            "arguments[0].removeAttribute('expanded', '');", comercialization_dropdown
        )
        time.sleep(1)
        print("✅ Nível de comercialização selecionado")

        # ── UFs: todas ────────────────────────────────────────────────────────
        location_dropdown = driver.find_element(
            By.ID, 'unidadeFederacao'
        ).find_element(By.XPATH, './/div[@tabindex]')
        driver.execute_script(
            "arguments[0].setAttribute('expanded', '');", location_dropdown
        )
        driver.find_element(
            By.XPATH, "//label[@for='unidadeFederacao-all']"
        ).click()
        driver.execute_script(
            "arguments[0].removeAttribute('expanded', '');", location_dropdown
        )
        print("✅ UFs selecionadas")

        # ── Download ──────────────────────────────────────────────────────────
        limpa_arquivo_antigo(downloads_path, NOME_ARQUIVO_PRECOS)
        arquivos = os.listdir(downloads_path)

        driver.find_element(By.XPATH, '//button[text()=" Consultar "]').click()
        time.sleep(1)
        print("✅ Consulta executada, baixando Excel...")
        driver.find_element(By.XPATH, '//button[@aria-label="Excel"]').click()
        print("✅ Botão Excel clicado, aguardando download...")

        sucesso = aguarda_arquivo_especifico(
            downloads_path, NOME_ARQUIVO_PRECOS, timeout=120
        )
        if not sucesso:
            driver.save_screenshot('precos_cultivos_timeout.png')
            raise FileNotFoundError(
                f"Download de '{NOME_ARQUIVO_PRECOS}' não foi concluído a tempo"
            )

        print("✅ Download de Preço Cultivos concluído")

    except Exception as e:
        print(f"❌ Erro em get_cultures_prices: {e}")
        driver.save_screenshot('precos_cultivos_erro.png')
        raise


# ==========================================================
# =================== FUNÇÕES DO PIB (CEPEA) ===============
# ==========================================================

def get_pib_values_direto(pasta_downloads, url_direta=None):
    """
    ✅ Baixa o Excel do PIB Agro do CEPEA diretamente via requests,
    contornando o Cloudflare Turnstile que bloqueia 100% das tentativas
    via Selenium. O arquivo tem URL estável e pública.
    """
    if url_direta is None:
        url_direta = (
            "https://www.cepea.esalq.usp.br/upload/kceditor/files/"
            "Pib_Cepea_Agro_Brasil.xlsx"
        )

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": (
            "text/html,application/xhtml+xml,application/xml;q=0.9,"
            "image/avif,image/webp,image/apng,*/*;q=0.8"
        ),
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept-Encoding": "gzip, deflate, br",
        "Referer": (
            "https://www.cepea.esalq.usp.br/br/"
            "pib-do-agronegocio-brasileiro.aspx"
        ),
        "Connection": "keep-alive",
    }

    print(f"📥 Tentando download direto do PIB CEPEA: {url_direta}")

    try:
        response = requests.get(
            url_direta, headers=headers, timeout=60, stream=True
        )
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if "spreadsheetml" in content_type or url_direta.endswith(".xlsx"):
            ext = "xlsx"
        elif "ms-excel" in content_type or url_direta.endswith(".xls"):
            ext = "xls"
        else:
            ext = url_direta.split(".")[-1].split("?")[0] or "xlsx"

        nome_arquivo = f"PIB.{ext}"
        destino = os.path.join(pasta_downloads, nome_arquivo)

        with open(destino, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        tamanho = os.path.getsize(destino)
        print(f"✅ PIB baixado com sucesso: {nome_arquivo} ({tamanho} bytes)")
        return nome_arquivo

    except requests.exceptions.HTTPError as e:
        print(f"❌ Erro HTTP ao baixar PIB: {e}")
    except requests.exceptions.ConnectionError as e:
        print(f"❌ Erro de conexão ao baixar PIB: {e}")
    except requests.exceptions.Timeout:
        print("❌ Timeout ao baixar PIB via requests")
    except Exception as e:
        print(f"❌ Erro inesperado ao baixar PIB: {e}")

    return None


def get_pib_url_dinamica(pagina_url):
    """
    ✅ Busca a URL do Excel do PIB na página do CEPEA via requests puro
    (sem Selenium), evitando o Cloudflare Turnstile. Usado como fallback
    caso a URL direta mude.
    """
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Referer": "https://www.google.com/",
        "Accept-Language": "pt-BR,pt;q=0.9",
    }

    try:
        resp = requests.get(pagina_url, headers=headers, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        for tag in soup.find_all("a", href=True):
            href = tag["href"]
            if re.search(r"\.(xlsx|xls)(\?.*)?$", href, re.IGNORECASE):
                if href.startswith("http"):
                    return href
                else:
                    from urllib.parse import urljoin
                    return urljoin(pagina_url, href)

    except Exception as e:
        print(f"⚠️ Não foi possível descobrir URL dinâmica do PIB: {e}")

    return None


# ==========================================================
# ===================== LINKS ==============================
# ==========================================================

link_graos        = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-graos"
link_cana         = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cana-de-acucar"
link_cafe         = "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/safra-de-cafe"
link_VBP          = "https://www.gov.br/agricultura/pt-br/assuntos/politica-agricola/valor-bruto-da-producao-agropecuaria-vbp"
link_dolar        = 'https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes'
link_credito_rural = 'https://www.bndes.gov.br/wps/portal/site/home/transparencia/consulta-operacoes-bndes/credito-rural-desempenho-operacional'
link_podas        = 'https://www.ibge.gov.br/explica/producao-agropecuaria/br'
link_precos_cultivos = 'https://consultaprecosdemercado.conab.gov.br/#/home'
link_pib_agro     = 'https://www.cepea.esalq.usp.br/br/pib-do-agronegocio-brasileiro.aspx'

URL_PIB_DIRETA = (
    "https://www.cepea.esalq.usp.br/upload/kceditor/files/"
    "Pib_Cepea_Agro_Brasil.xlsx"
)


# ==========================================================
# ===================== EXECUÇÃO ===========================
# ==========================================================

try:
    # Limpeza inicial da pasta de downloads temporários
    archs = os.listdir(downloads_path)
    if archs:
        for i in archs:
            complete_path = os.path.join(downloads_path, i)
            if os.path.isfile(complete_path):
                print(f"🗑️ Removendo arquivo antigo: {complete_path}")
                remove_seguro(complete_path)

    arquivos = os.listdir(downloads_path)
    tempo_inicial = datetime.now()

    # ── Safras ────────────────────────────────────────────────────────────────
    get_safras(driver, link_cafe, "cafe")
    verify_new_download("Cafe")
    move_to_dados_teste("Cafe.xlsx")

    get_safras(driver, link_graos, "graos")
    verify_new_download("Graos")
    move_to_dados_teste("Graos.xlsx")

    get_safras(driver, link_cana, "cana")
    verify_new_download("Cana")
    move_to_dados_teste("Cana.xlsx")

    # ── Consolida os 3 em um único Safras.xlsx ────────────────────────────────
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
                df_aba = pd.read_excel(
                    caminho, sheet_name=aba, engine="openpyxl", header=None
                )

                if df_aba.dropna(how="all").empty:
                    print(f"⏭️ Aba vazia ignorada: [{cultura}] {aba}")
                    continue

                prefixo   = cultura[:3].upper()
                nome_base = f"{prefixo} - {aba}"[:31]
                nome_final = nome_base
                contador  = 1
                while nome_final in nomes_usados:
                    sufixo     = f"_{contador}"
                    nome_final = nome_base[:31 - len(sufixo)] + sufixo
                    contador  += 1

                nomes_usados.append(nome_final)
                df_aba.to_excel(
                    writer, sheet_name=nome_final, index=False, header=False
                )
                print(f"✅ Aba adicionada: {nome_final}")

    print("✅ Safras.xlsx consolidado com sucesso!")

    # ── VBP ───────────────────────────────────────────────────────────────────
    get_VBP(driver, link_VBP)
    verify_new_download("VBP")
    move_to_dados_teste("VBP.xlsx")

    # ── Crédito Rural ─────────────────────────────────────────────────────────
    get_credito_rural(driver, link_credito_rural)
    verify_new_download("Credito Rural")
    move_to_dados_teste("Credito Rural.xlsx", "Credito_Rural.xlsx")

    # ── Preços de Cultivos ────────────────────────────────────────────────────
    get_cultures_prices(driver, link_precos_cultivos)
    move_to_dados_teste("Consulta-precos-mensal.xlsx", "Preço Cultivos.xlsx")

    # ── PIB Agro (CEPEA) — download direto via requests (sem Selenium) ────────
    # 🛑 O CEPEA usa Cloudflare Turnstile (clique humano obrigatório),
    # que bloqueia 100% das tentativas via Selenium. Solução: requests puro.
    print("📥 Iniciando download do PIB via requests (sem Selenium)...")
    nome_pib = get_pib_values_direto(downloads_path, url_direta=URL_PIB_DIRETA)

    if nome_pib is None:
        # Fallback: tenta descobrir a URL dinamicamente na página
        print("⚠️ URL direta falhou. Tentando descoberta dinâmica da URL do PIB...")
        url_descoberta = get_pib_url_dinamica(link_pib_agro)
        if url_descoberta:
            print(f"🔍 URL descoberta dinamicamente: {url_descoberta}")
            nome_pib = get_pib_values_direto(
                downloads_path, url_direta=url_descoberta
            )

    if nome_pib:
        move_to_dados_teste(nome_pib, "PIB.xlsx")
    else:
        print("❌ Não foi possível baixar o PIB por nenhuma estratégia")

    # ── Dólar ─────────────────────────────────────────────────────────────────
    get_dolar(driver, link_dolar)
    nome_dolar = verify_new_download("Dolar")
    move_to_dados_teste(nome_dolar, "Dolar.csv")

    # ── Podas (salva direto em dados_teste/ dentro da função) ─────────────────
    get_podas_tables(driver, link_podas)

except FileNotFoundError as e:
    print(f"Um dos arquivos não foi baixado: {e}")
    raise

except Exception as e:
    print('Ocorreu um erro na execução do script: ', e)

# finally:
#     tempo_final = datetime.now()
#     print("Tempo de Execução: ", tempo_final - tempo_inicial)
#     driver.quit()
