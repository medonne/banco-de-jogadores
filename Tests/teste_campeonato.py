import requests
from bs4 import BeautifulSoup


# =========================================================
# COLE O LINK DE UM JOGADOR AQUI
# =========================================================

url = "https://www.ogol.com.br/jogador/gabriel-oliveira/498858?epoca_id=155"


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


# =========================================================
# PROCURA POSSÍVEIS CAMPEONATOS
# =========================================================

palavras = [
    "SÉRIE A",
    "SÉRIE B",
    "SÉRIE C",
    "SÉRIE D",
    "BRASILEIRÃO",
    "GAÚCHO",
    "GAUCHÃO",
    "CATARINENSE",
    "PAULISTA",
    "PARANAENSE",
    "MINEIRO",
    "GOIANO",
    "PERNAMBUCANO",
    "CEARENSE",
    "BAIANO",
    "COPA DO BRASIL"
]


print("\n================================")
print("POSSÍVEIS CAMPEONATOS")
print("================================")


for i, linha in enumerate(linhas):

    linha_maiuscula = linha.upper()

    if any(
        palavra in linha_maiuscula
        for palavra in palavras
    ):

        print("\n------------------------------")

        inicio = max(
            0,
            i - 4
        )

        fim = min(
            len(linhas),
            i + 7
        )

        for j in range(
            inicio,
            fim
        ):

            print(
                f"{j}: {linhas[j]}"
            )