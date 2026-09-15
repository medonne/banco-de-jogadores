import requests
from bs4 import BeautifulSoup

# Use o mesmo jogador que testamos
url = "https://www.ogol.com.br/jogador/tete/438954?epoca_id=143"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9"
}

session = requests.Session()
session.headers.update(headers)

# =====================================================
# ABRE PERFIL DO JOGADOR
# =====================================================

resposta = session.get(url, timeout=20)

print("HTTP JOGADOR:", resposta.status_code)

soup = BeautifulSoup(resposta.text, "html.parser")

# Procura o texto "Clube atual"
texto_clube = soup.find(
    string=lambda t: t and "Clube atual" in t
)

if not texto_clube:
    print("Não encontrei 'Clube atual'.")
    raise SystemExit

print("\n==============================")
print("BLOCO DO CLUBE ATUAL")
print("==============================")

elemento = texto_clube.parent

# Sobe alguns níveis para mostrar o bloco HTML
for nivel in range(5):

    if elemento is None:
        break

    print(f"\n--- NÍVEL {nivel} ---")
    print(elemento.prettify()[:4000])

    elemento = elemento.parent