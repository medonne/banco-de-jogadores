import requests
from bs4 import BeautifulSoup


# =========================================================
# LINK DO MESMO JOGADOR
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


# =========================================================
# DESCOBRE CLUBE ATUAL
# =========================================================

clube_atual = None

for linha in soup.select(".card-data__row"):

    label = linha.select_one(".card-data__label")

    if not label:
        continue

    if label.get_text(" ", strip=True).lower() == "clube atual":

        clube = linha.select_one(
            ".micrologo_and_text .text"
        )

        if clube:
            clube_atual = clube.get_text(
                " ",
                strip=True
            )

        break


print("\n================================")
print("CLUBE ATUAL")
print("================================")

print(clube_atual)


# =========================================================
# TRANSFORMA PÁGINA EM LINHAS
# =========================================================

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
# MOSTRA TODAS AS APARIÇÕES DO CLUBE ATUAL
# =========================================================

print("\n================================")
print("ONDE O CLUBE ATUAL APARECE")
print("================================")


if clube_atual:

    for i, linha in enumerate(linhas):

        if clube_atual.lower() in linha.lower():

            print("\n--------------------------------")
            print(f"APARIÇÃO NA LINHA {i}")
            print("--------------------------------")

            inicio = max(
                0,
                i - 15
            )

            fim = min(
                len(linhas),
                i + 20
            )

            for j in range(
                inicio,
                fim
            ):

                print(
                    f"{j}: {linhas[j]}"
                )