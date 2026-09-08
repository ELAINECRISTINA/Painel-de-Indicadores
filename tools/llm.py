import os
import ssl
import httpx
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==============================================================
# ✅ MONKEY-PATCH: força verify=False e trust_env=True
# em todo httpx ANTES do SDK ser importado
# ==============================================================
_original_init = httpx.Client.__init__

def _patched_init(self, *args, **kwargs):
    kwargs["verify"]    = False  # ignora SSL corporativo
    kwargs["trust_env"] = True   # usa proxy do sistema automaticamente
    kwargs.setdefault("timeout", 60.0)
    _original_init(self, *args, **kwargs)

httpx.Client.__init__ = _patched_init

# Mesmo patch para cliente assíncrono (usado internamente pelo SDK)
_original_async_init = httpx.AsyncClient.__init__

def _patched_async_init(self, *args, **kwargs):
    kwargs["verify"]    = False
    kwargs["trust_env"] = True
    kwargs.setdefault("timeout", 60.0)
    _original_async_init(self, *args, **kwargs)

httpx.AsyncClient.__init__ = _patched_async_init

# ==============================================================
# SÓ AGORA importa o SDK — ele usará o httpx já patchado
# ==============================================================
from google import genai
from google.genai import types
import requests
from bs4 import BeautifulSoup


def init_llm():
    try:
        GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
        if not GEMINI_API_KEY:
            print("ERRO: GEMINI_API_KEY não encontrada.")
            return None

        # Cria normalmente — patch já garante verify=False + trust_env=True
        client = genai.Client(api_key=GEMINI_API_KEY)
        print("✅ Cliente Gemini inicializado")
        return client

    except Exception as e:
        print(f"ERRO: {e}")
        return None