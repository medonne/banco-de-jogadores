import requests
from bs4 import BeautifulSoup

url = "https://www.ogol.com.br/jogador/matheus-lagoa/465575?search=1"

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

for i, linha in enumerate(linhas):

    if (
        "NASCIMENTO" in linha.upper()
        or "NATURALIDADE" in linha.upper()
    ):

        print("\n==============================")
        print("TRECHO ENCONTRADO")
        print("==============================")

        inicio = max(0, i - 3)
        fim = min(len(linhas), i + 10)

        for j in range(inicio, fim):
            print(f"{j}: {linhas[j]}")