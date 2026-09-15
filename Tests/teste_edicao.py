import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin


url_jogador = "https://www.ogol.com.br/jogador/gabriel-oliveira/498858?epoca_id=155"

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


# =========================================================
# ABRE PERFIL DO JOGADOR
# =========================================================

resposta = session.get(
    url_jogador,
    timeout=20
)

print("HTTP JOGADOR:", resposta.status_code)

soup = BeautifulSoup(
    resposta.text,
    "html.parser"
)


# =========================================================
# PROCURA O PRIMEIRO LINK RS2
# =========================================================

link_edicao = None

for link in soup.find_all("a", href=True):

    texto = link.get_text(
        " ",
        strip=True
    )

    if texto == "RS2":

        link_edicao = urljoin(
            url_jogador,
            link["href"]
        )

        break


print("\nLink da edição:")
print(link_edicao)


if not link_edicao:
    print("Não encontrei a edição.")
    raise SystemExit


# =========================================================
# ABRE A PÁGINA DA COMPETIÇÃO
# =========================================================

resposta_edicao = session.get(
    link_edicao,
    timeout=20
)

print("\nHTTP EDIÇÃO:", resposta_edicao.status_code)


soup_edicao = BeautifulSoup(
    resposta_edicao.text,
    "html.parser"
)


# =========================================================
# MOSTRA TÍTULO DA PÁGINA
# =========================================================

print("\n================================")
print("TÍTULO")
print("================================")

if soup_edicao.title:
    print(
        soup_edicao.title.get_text(
            " ",
            strip=True
        )
    )


# =========================================================
# PROCURA CABEÇALHOS
# =========================================================

print("\n================================")
print("CABEÇALHOS")
print("================================")

for tag in soup_edicao.find_all(
    ["h1", "h2", "h3"]
):

    texto = tag.get_text(
        " ",
        strip=True
    )

    if texto:
        print(texto)


# =========================================================
# PRIMEIRAS LINHAS ÚTEIS
# =========================================================

texto = soup_edicao.get_text(
    "\n",
    strip=True
)

linhas = [
    linha.strip()
    for linha in texto.splitlines()
    if linha.strip()
]


print("\n================================")
print("PRIMEIRAS 150 LINHAS")
print("================================")

for i, linha in enumerate(
    linhas[:150]
):

    print(
        f"{i}: {linha}"
    )