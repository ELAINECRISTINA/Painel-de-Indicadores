import os
from google import genai
from google.genai import types

import requests
from bs4 import BeautifulSoup


def init_llm():
    try:
        # 🛑 FIX: a key estava hardcoded no código-fonte e foi detectada
        # como vazada pelo Google (bloqueio automático). Agora ela vem de
        # uma variável de ambiente, nunca do código versionado.
        GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

        if not GEMINI_API_KEY:
            print(
                "ERRO: Variável de ambiente GEMINI_API_KEY não encontrada. "
                "Defina-a antes de rodar o script (veja instruções abaixo)."
            )
            return None

        # Nova API: criar cliente diretamente
        client = genai.Client(api_key=GEMINI_API_KEY)
        return client
    except Exception as e:
        print(f"ERRO: Falha ao inicializar o cliente: {e}")
        return None  # Retorna None explicitamente em caso de erro
