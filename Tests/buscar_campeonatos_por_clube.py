import re
import time
import random
import requests

from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from pathlib import Path
from urllib.parse import urljoin


# =========================================================
# CONFIGURAÇÕES
# =========================================================

arquivo_entrada = "Revisar_Campeonatos.xlsx"
arquivo_saida = "Sugestoes_Campeonatos.xlsx"

TEMPO_MINIMO = 2
TEMPO_MAXIMO = 4


# =========================================================
# CAMINHOS
# =========================================================

pasta = Path(__file__).resolve().parent

caminho_entrada = pasta / arquivo_entrada
caminho_saida = pasta / arquivo_saida


# =========================================================
# ABRE PLANILHA
# =========================================================

wb = load_workbook(caminho_entrada)
ws = wb.active


cabecalhos = {}

for celula in ws[1]:

    if celula.value:

        cabecalhos[
            str(celula.value).strip()
        ] = celula.column


col_id = cabecalhos["ID_Jogador"]
col_nome = cabecalhos["Nome"]
col_clube = cabecalhos["Clube_Atual"]
col_pais = cabecalhos["Pais"]
col_campeonato = cabecalhos["Campeonato_Atual_Planilha"]
col_link = cabecalhos["Link_Ogol"]


# =========================================================
# CONEXÃO
# =========================================================

session = requests.Session()

session.headers.update({

    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0.0.0 Safari/537.36"
    ),

    "Accept-Language": "pt-BR,pt;q=0.9"
})


# =========================================================
# ENCONTRA LINK DO CLUBE EM 2026
# =========================================================

def buscar_link_clube(
    url_jogador,
    clube
):

    try:

        resposta = session.get(
            url_jogador,
            timeout=20
        )

    except requests.RequestException:

        return None


    if resposta.status_code != 200:

        return None


    soup = BeautifulSoup(
        resposta.text,
        "html.parser"
    )


    candidatos = []


    for link in soup.find_all(
        "a",
        href=True
    ):

        texto = link.get_text(
            " ",
            strip=True
        )

        href = link.get(
            "href",
            ""
        )


        if not texto:
            continue


        if clube.lower() not in texto.lower():
            continue


        if "/equipe/" not in href:
            continue


        url = urljoin(
            url_jogador,
            href
        )


        candidatos.append(url)


    if not candidatos:

        return None


    # =====================================================
    # epoca_id=155 = temporada 2026
    # =====================================================

    for url in candidatos:

        if "epoca_id=155" in url:

            return url


    # Se não encontrou com temporada explícita,
    # usa o link simples do clube

    for url in candidatos:

        if "epoca_id=" not in url:

            return url


    return candidatos[0]


# =========================================================
# LIMPA NOME
# =========================================================

def limpar_nome(nome):

    nome = re.sub(
        r"\s+20\d{2}\s*$",
        "",
        nome
    )

    nome_upper = nome.upper()


    # Séries nacionais brasileiras

    if (
        "BRASILEIRÃO" in nome_upper
        and "SÉRIE A" in nome_upper
    ):
        return "SÉRIE A"


    if (
        "BRASILEIRÃO" in nome_upper
        and "SÉRIE B" in nome_upper
    ):
        return "SÉRIE B"


    if (
        "BRASILEIRÃO" in nome_upper
        and "SÉRIE C" in nome_upper
    ):
        return "SÉRIE C"


    if (
        "BRASILEIRÃO" in nome_upper
        and "SÉRIE D" in nome_upper
    ):
        return "SÉRIE D"


    # Gaúcho acesso

    if (
        "GAÚCHO" in nome_upper
        and "ACESSO" in nome_upper
    ):

        return "GAÚCHO ACESSO"


    nome = re.sub(
        r"^Campeonato\s+",
        "",
        nome,
        flags=re.IGNORECASE
    )


    return nome.upper()


# =========================================================
# BUSCA COMPETIÇÕES DO CLUBE
# =========================================================

def buscar_competicoes_clube(url_clube):

    try:

        resposta = session.get(
            url_clube,
            timeout=20
        )

    except requests.RequestException:

        return []


    if resposta.status_code != 200:

        return []


    soup = BeautifulSoup(
        resposta.text,
        "html.parser"
    )


    competicoes = []


    for link in soup.find_all(
        "a",
        href=True
    ):

        href = link.get(
            "href",
            ""
        )


        if "edition.php?id_edicao=" not in href:
            continue


        texto = link.get_text(
            " ",
            strip=True
        )


        if not texto:
            continue


        # Só queremos temporada 2026
        if "2026" not in texto:
            continue


        url_edicao = urljoin(
            url_clube,
            href
        )


        item = (
            texto,
            url_edicao
        )


        if item not in competicoes:

            competicoes.append(item)


    return competicoes


# =========================================================
# ESCOLHE CAMPEONATO PRINCIPAL
# =========================================================

def escolher_campeonato(
    competicoes,
    pais
):

    if not competicoes:

        return None, None


    pais_upper = str(
        pais or ""
    ).upper()


    # =====================================================
    # BRASIL
    # prioridade absoluta:
    # Série A > B > C > D
    # =====================================================

    if pais_upper == "BRASIL":

        prioridades = [
            "BRASILEIRÃO SÉRIE A",
            "BRASILEIRÃO SÉRIE B",
            "BRASILEIRÃO SÉRIE C",
            "BRASILEIRÃO SÉRIE D"
        ]


        for prioridade in prioridades:

            for nome, url in competicoes:

                if prioridade in nome.upper():

                    return (
                        limpar_nome(nome),
                        url
                    )


    # =====================================================
    # SE NÃO ACHOU LIGA NACIONAL:
    # elimina Copas e amistosos
    # =====================================================

    ignorar = [
        "COPA",
        "SUPERCOPA",
        "RECOPA",
        "AMISTOSO",
        "AMISTOSOS",
        "TAÇA"
    ]


    candidatos = []


    for nome, url in competicoes:

        nome_upper = nome.upper()


        if any(
            palavra in nome_upper
            for palavra in ignorar
        ):

            continue


        candidatos.append(
            (nome, url)
        )


    # Se sobrou apenas uma competição de liga/estadual

    if candidatos:

        nome, url = candidatos[0]

        return (
            limpar_nome(nome),
            url
        )


    return None, None


# =========================================================
# PLANILHA DE RESULTADO
# =========================================================

wb_saida = Workbook()

ws_saida = wb_saida.active

ws_saida.title = "Sugestões"


ws_saida.append([
    "ID_Jogador",
    "Nome",
    "Clube_Atual",
    "Pais",
    "Campeonato_Antigo",
    "Campeonato_Sugerido",
    "Link_Clube_2026",
    "Link_Edicao",
    "Todas_Competicoes"
])


# =========================================================
# PROCESSAMENTO
# =========================================================

total = 0
encontrados = 0
nao_encontrados = 0


for linha in range(
    2,
    ws.max_row + 1
):

    total += 1


    id_jogador = ws.cell(
        linha,
        col_id
    ).value


    nome = ws.cell(
        linha,
        col_nome
    ).value


    clube = ws.cell(
        linha,
        col_clube
    ).value


    pais = ws.cell(
        linha,
        col_pais
    ).value


    campeonato_antigo = ws.cell(
        linha,
        col_campeonato
    ).value


    link_jogador = ws.cell(
        linha,
        col_link
    ).value


    print("\n================================")

    print(
        f"[{total}] {nome}"
    )

    print(
        f"Clube: {clube}"
    )


    # =====================================================
    # LINK DO CLUBE
    # =====================================================

    link_clube = buscar_link_clube(
        str(link_jogador),
        str(clube)
    )


    if not link_clube:

        print(
            "Link do clube não encontrado."
        )

        nao_encontrados += 1


        ws_saida.append([
            id_jogador,
            nome,
            clube,
            pais,
            campeonato_antigo,
            None,
            None,
            None,
            None
        ])


        wb_saida.save(
            caminho_saida
        )


        time.sleep(
            random.uniform(
                TEMPO_MINIMO,
                TEMPO_MAXIMO
            )
        )

        continue


    print(
        f"Link clube: {link_clube}"
    )


    # =====================================================
    # COMPETIÇÕES
    # =====================================================

    competicoes = buscar_competicoes_clube(
        link_clube
    )


    print("Competições:")


    for nome_comp, _ in competicoes:

        print(
            f"   - {nome_comp}"
        )


    campeonato_novo, link_edicao = escolher_campeonato(
        competicoes,
        pais
    )


    if campeonato_novo:

        encontrados += 1

        print(
            f">>> SUGESTÃO: {campeonato_novo}"
        )

    else:

        nao_encontrados += 1

        print(
            ">>> NÃO ENCONTRADO"
        )


    todas = " | ".join(
        nome
        for nome, _ in competicoes
    )


    ws_saida.append([
        id_jogador,
        nome,
        clube,
        pais,
        campeonato_antigo,
        campeonato_novo,
        link_clube,
        link_edicao,
        todas
    ])


    # Salva a cada jogador
    wb_saida.save(
        caminho_saida
    )


    time.sleep(
        random.uniform(
            TEMPO_MINIMO,
            TEMPO_MAXIMO
        )
    )


# =========================================================
# FINAL
# =========================================================

wb_saida.save(
    caminho_saida
)


print("\n================================")
print("PROCESSO FINALIZADO")
print("================================")

print(
    f"Jogadores analisados: {total}"
)

print(
    f"Campeonatos sugeridos: {encontrados}"
)

print(
    f"Não encontrados: {nao_encontrados}"
)

print("\nArquivo criado:")

print(
    caminho_saida
)