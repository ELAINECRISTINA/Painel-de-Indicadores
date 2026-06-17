

from PIL import Image
import os
import time
from datetime import datetime
from pprint import pprint
import random
import pandas as pd
from thefuzz import process
import json

from tools.llm import init_llm

PROMPT_DE_INSTRUCAO = (
"""Ignore todo o ruído visual, como linhas, pontos e distorções de fundo na imagem em anexo. 
Extraia e retorne apenas a sequência de caracteres alfanuméricos necessária para resolver o captcha. 
O resultado deve ser um texto sem espaços.
"""
)


MODEL = init_llm()
ITENS_NECESSARIOS = [
    'Maçãs', 'Abacates', 'Melões e outros melões', 'Grãos de cacau',
    'Café, verde', 'Figos', 'Uvas', 'Alho verde', 'Limões e limas',
    'Mangas, goiabas e mangostões', 'Azeitonas', 'Laranjas',
    'Pêssegos e nectarinas', 'Peras', 'Marmelos',
    'Tangerinas, tangerinas, clementinas', 'Tomates', 'Melancias', 'Banana',
    'Cana-de-Açúcar', 'Milho'
]
def analisar_imagem_com_gemini(caminho_da_imagem: str, model: str) -> str:
    try:
        # Carrega a imagem do caminho especificado usando a biblioteca Pillow (PIL)
        imagem = Image.open(caminho_da_imagem)

        # Inicializa o cliente do Gemini
        client = init_llm()
        
        # Verifica se o cliente foi inicializado corretamente
        if client is None:
            return "Erro: Falha ao inicializar o modelo LLM"
        
        # Nova API: usar client.models.generate_content
        print("LLM Trabalhando...")
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[PROMPT_DE_INSTRUCAO, imagem]
        )
        print("LLM Concluída")
        print("output: ", response.text)
        # Retorna o texto extraído da resposta do modelo.
        return response.text

    except FileNotFoundError:
        return f"Erro: O arquivo de imagem não foi encontrado em '{caminho_da_imagem}'."
    except Exception as e:
        # Captura outras possíveis exceções (ex: chave de API inválida, problema de conexão)
        return f"Ocorreu um erro inesperado: {e}"

def categorizar_com_llm(produto: str) -> str:
    """Função que faz a chamada à API para um único produto."""
    
    prompt = f"""
    Você é um especialista de produtos Agrícolas. Identifique em qual das categorias de {ITENS_NECESSARIOS} o produto "{produto}" se encaixa.

    Retorne APENAS o nome da categoria.
    Se não se encaixar em nenhuma, retorne "Não Categorizado".
    """
    try:
        response = MODEL.models.generate_content(
            model="gemini-2.5-flash", 
            contents=[prompt] 
        )
        categoria = response.text.strip()
        print(f'LLM: "{produto}" --> "{categoria}"')
        return categoria
    except Exception as e:
        print(f"Erro ao processar o produto {produto}: {e}")
        return "ERRO"

# 3. MELHORE A LÓGICA DE VERIFICAÇÃO INICIAL
def categorizar_produto_otimizado(produto: str) -> str:

    
    """Tenta categorizar usando lógica simples primeiro, depois usa o LLM como fallback."""
    melhor_match, score = process.extractOne(produto, ITENS_NECESSARIOS)

    # Se a pontuação de similaridade for alta (ex: > 85), usamos essa correspondência.
    if score > 85:
        return melhor_match

    # Se a pontuação for baixa, significa que o produto é muito diferente.
    # SÓ ENTÃO chamamos a IA.
    return categorizar_com_llm(produto)

def categorizar_em_lote_com_llm(produtos: list[str]) -> dict:
    """Envia uma lista de produtos para o LLM e pede um JSON de volta."""
    
    # Converte a lista de produtos em uma string formatada
    lista_produtos_str = "\n".join([f'- "{p}"' for p in produtos])

    prompt = f"""
    Você é um especialista em produtos agrícolas. Sua tarefa é categorizar a lista de produtos abaixo.
    Para cada produto, identifique em qual das categorias a seguir ele se encaixa: {ITENS_NECESSARIOS}.

    Produtos para categorizar:
    {lista_produtos_str}

    Responda com um único objeto JSON onde a chave é o nome do produto e o valor é a categoria correspondente.
    Se um produto não se encaixar em nenhuma categoria, use o valor "Não Categorizado".
    
    Exemplo de resposta:
    {{
      "Milho Doce": "Milho",
      "Café arábica": "Café, verde",
      "Alface": "Não Categorizado"
    }}
    """
    
    try:
        response = MODEL.models.generate_content(
            model="gemini-2.0-flash-lite",
            contents=[prompt]
        )
        # Limpa a resposta para garantir que seja um JSON válido
        resposta_limpa = response.text.strip().replace("```json", "").replace("```", "")
        print("--- Resposta do LLM em Lote ---")
        print(resposta_limpa)
        print("-----------------------------")
        return json.loads(resposta_limpa)
    except Exception as e:
        print(f"Erro ao processar o lote de produtos: {e}")
        # Retorna um dicionário de erro para cada produto no lote
        return {produto: "ERRO" for produto in produtos}


def retorna_labels_certos(produtos:list[dict]) -> list:
    produtos_formatados = '\n'.join([str(prod) for prod in produtos])
    PROMPT = f"""Você é um especialista de produtos. Seu dever é Identificar tudo o que é um produto VEGETAL e retornar o seu ID.
                lembre-se, apenas os vegetais, frutas e legumas, nada de produtos de origem de animais ou até mesmo óleos provenientes de vegetais.
                Você deve retornar em uma lista exatamente os ids, separados por ';' dos produtos que são vegetais. 
                Seus produtos são:
                {produtos_formatados}

                #FORMATO DA LISTA:
                [produtoxx;produtoxy;produtoyy;...]

                #VOCÊ NÃO DEVE RETORNAR NADA ALÉM DISSO, SEM 'Claro, aqui está sua lista', retorne APENAS A LISTA ESTRUTURADA

                """
    try:
        response = MODEL.models.generate_content(
            model="gemini-2.5-flash", 
            contents=[PROMPT] 
        )
        lista_estruturada = response.text
        
        lista = lista_estruturada.replace('[','').replace(']','').split(';')
        return lista
    except Exception as e:
        raise e
        return "ERRO"


def retorna_pib_atualizado(markdown) -> str:
    PROMPT = f"""
    Você é um especialista que lida com dados da agricultura diariamente. Seu papel é ler o Markdown e retornar o Valor do PIB AGRO no ano de {datetime.now().year}.
    
    Aqui estão seus dados: {markdown}

    Retorne APENAS o PIB AGRO que você encontrou. Traga ele, de maneira numérica, exemplo: $2,580,000,000,000

    """
    llm = init_llm()
    try:
        response = MODEL.models.generate_content(
            model='gemini-2.5-flash',
            contents=[PROMPT]
        )
        pib_val = response.text
    except Exception as e:
        raise e

    return pib_val