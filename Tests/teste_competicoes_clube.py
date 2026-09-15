import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin


url_clube = "https://www.ogol.com.br/equipe/retro-fc/240756?epoca_id=155"


headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9"
}


resposta = requests.get(
    url_clube,
    headers=headers,
    timeout=20
)

print("HTTP:", resposta.status_code)


soup = BeautifulSoup(
    resposta.text,
    "html.parser"
)


print("\n================================")
print("TÍTULO")
print("================================")

if soup.title:
    print(
        soup.title.get_text(
            " ",
            strip=True
        )
    )


print("\n================================")
print("LINKS DE EDIÇÕES")
print("================================")


ja_mostrados = set()


for link in soup.find_all("a", href=True):

    href = link.get("href", "")

    if "edition.php?id_edicao=" not in href:
        continue

    texto = link.get_text(
        " ",
        strip=True
    )

    url = urljoin(
        url_clube,
        href
    )

    chave = (
        texto,
        url
    )

    if chave in ja_mostrados:
        continue

    ja_mostrados.add(chave)

    print("\n------------------------------")
    print("Texto:", texto)
    print("Link:", url)


print("\n================================")
print("TRECHOS COM COMPETIÇÕES")
print("================================")


texto_pagina = soup.get_text(
    "\n",
    strip=True
)

linhas = [
    linha.strip()
    for linha in texto_pagina.splitlines()
    if linha.strip()
]


palavras = [
    "SÉRIE",
    "BRASILEIRÃO",
    "PERNAMBUCANO",
    "COPA DO BRASIL",
    "NORDESTE",
    "ESTADUAL"
]


for i, linha in enumerate(linhas):

    linha_upper = linha.upper()

    if any(
        palavra in linha_upper
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