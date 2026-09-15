import requests
from bs4 import BeautifulSoup


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


print("\n================================")
print("ELEMENTOS COM RS2")
print("================================")


encontrados = 0


for elemento in soup.find_all(string=True):

    texto = elemento.strip()

    if texto == "RS2":

        encontrados += 1

        print("\n--------------------------------")
        print(f"RS2 Nº {encontrados}")
        print("--------------------------------")

        pai = elemento.parent

        print("\nTAG:")
        print(pai.prettify())

        print("\nPAI:")
        if pai.parent:
            print(pai.parent.prettify()[:3000])


print("\n================================")
print(f"TOTAL DE RS2 ENCONTRADOS: {encontrados}")
print("================================")