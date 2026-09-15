import requests
from bs4 import BeautifulSoup

# Coloque aqui o link de um jogador que você sabe
# qual é o clube atual
url = "https://www.ogol.com.br/jogador/tete/438954?epoca_id=143"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9"
}

resposta = requests.get(
    url,
    headers=headers,
    timeout=20
)

print("HTTP:", resposta.status_code)

soup = BeautifulSoup(
    resposta.text,
    "html.parser"
)

texto = soup.get_text(
    "\n",
    strip=True
)

linhas = [
    linha.strip()
    for linha in texto.splitlines()
    if linha.strip()
]

# Mostra linhas que provavelmente têm
# clube, equipe, competição ou país
palavras = [
    "CLUBE",
    "EQUIPE",
    "COMPETIÇÃO",
    "CAMPEONATO",
    "PAÍS",
    "2026"
]

for i, linha in enumerate(linhas):

    if any(
        palavra in linha.upper()
        for palavra in palavras
    ):

        print("\n==============================")
        print(f"LINHA {i}")
        print("==============================")

        inicio = max(0, i - 3)
        fim = min(len(linhas), i + 8)

        for j in range(inicio, fim):
            print(f"{j}: {linhas[j]}")