from google import genai
from google.genai import types

import requests
from bs4 import BeautifulSoup

def init_llm():
    try:
        GEMINI_API_KEY = "your_key"
        # Nova API: criar cliente diretamente
        client = genai.Client(api_key=GEMINI_API_KEY)
        return client
    except Exception as e:
        print(f"ERRO: Falha ao inicializar o cliente: {e}")
        return None  # Retorna None explicitamente em caso de erro

