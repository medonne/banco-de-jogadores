# ⚽ Banco de Jogadores — Scouting de Futebol

Projeto desenvolvido para organizar e facilitar a consulta de informações de jogadores de futebol, reunindo dados de atletas em uma base estruturada para apoio à análise e ao scouting.

O projeto surgiu de uma necessidade real: manter uma base de jogadores organizada e atualizada, permitindo consultar informações como clube atual, campeonato, posição, idade, país e histórico de trabalho.

Além da organização dos dados, desenvolvi uma automação em **Python** para consultar informações públicas do OGol e atualizar dados que mudam ao longo do tempo.

---

## 🎯 Objetivo do projeto

Criar uma base centralizada de jogadores que possa ser atualizada periodicamente e posteriormente utilizada em análises no **SQL Server** e **Power BI**.

O fluxo do projeto é:

**OGol → Python → Excel → SQL Server → Power BI**

---

## 🗂️ Informações armazenadas

A base possui informações como:

- ID do jogador
- Nome
- Data de nascimento
- Posição
- Altura
- Clube atual
- Estado/país de origem
- Campeonato atual
- País do clube
- Pé dominante
- Histórico de trabalho
- Link do perfil no OGol
- Link da foto do jogador

---

## 🐍 Automação com Python

Foi desenvolvido um script em Python para atualizar automaticamente informações que podem mudar com o decorrer das temporadas:

- **Clube atual**
- **Campeonato**
- **País do clube**

O programa utiliza o link do perfil de cada jogador no OGol para consultar os dados.

Para identificar o campeonato atual, foram utilizadas duas estratégias:

1. Identificação da competição através dos jogos do jogador pelo clube atual.
2. Caso a primeira estratégia não encontre o campeonato, consulta das competições disputadas pelo clube na temporada.

Quando não é possível identificar a competição automaticamente, o registro é marcado como:

`NÃO ENCONTRADO`

Isso permite localizar esses casos posteriormente para revisão manual.

---

## 🛠️ Tecnologias utilizadas

- **Python**
- **Requests**
- **BeautifulSoup**
- **OpenPyXL**
- **Excel**
- **SQL Server**
- **Power BI**
- **Git e GitHub**
- **VS Code**

---

## 📁 Estrutura do projeto

```text
Banco de Jogadores/
│
├── Imagens/
├── SQL/
├── Tests/
│
├── atualizar_banco_jogadores.py
├── Banco_Jogadores.pbix
├── Banco_Jogadores.xlsx
└── README.md
```

> A base completa utilizada no projeto não é disponibilizada publicamente. Uma base de exemplo pode ser utilizada para demonstrar a estrutura dos dados.

---

## 🔄 Funcionamento

O arquivo principal utilizado localmente é:

`Banco_Jogadores.xlsx`

Ao executar:

```bash
python atualizar_banco_jogadores.py
```

o programa consulta os dados e gera:

`Banco_Jogadores_Atualizado.xlsx`

Dessa forma, a base original é preservada e o arquivo atualizado pode ser conferido antes da utilização.

---

## 🗄️ SQL Server

Os dados também podem ser armazenados no SQL Server, permitindo consultas e preparando a base para utilização em ferramentas de Business Intelligence.

A estrutura foi pensada para permitir consultas por características como:

- posição;
- idade;
- clube;
- campeonato;
- país;
- pé dominante;
- jogadores que já trabalharam com a comissão.

---

## 📊 Power BI

O projeto também prevê a utilização da base no Power BI para criação de um painel de scouting.

A proposta é permitir filtros e análises que facilitem a busca e comparação de jogadores de acordo com as necessidades da comissão técnica.

---

## 📚 Aprendizados

Este projeto está sendo desenvolvido como parte dos meus estudos em tecnologia e tem me permitido praticar:

- lógica de programação;
- Python;
- manipulação de arquivos Excel;
- coleta e tratamento de dados;
- web scraping;
- tratamento de exceções;
- automação de tarefas;
- organização e padronização de dados;
- SQL;
- integração entre diferentes ferramentas;
- versionamento com Git e GitHub.

---

## 🚧 Status do projeto

**Em desenvolvimento.**

A base e as automações continuam sendo aprimoradas conforme novas necessidades são identificadas.

Próximas etapas:

- revisar casos não identificados automaticamente;
- aprimorar a atualização das informações;
- finalizar a integração com SQL Server;
- desenvolver e aprimorar o dashboard no Power BI;
- adicionar novas funcionalidades à base.

---

## 👩‍💻 Autora

**Médonne Penteado**

Estudante de Análise e Desenvolvimento de Sistemas, desenvolvendo projetos práticos com Python, automação, SQL e Business Intelligence.

Este projeto foi criado para aplicar conhecimentos de programação e dados na solução de uma necessidade real relacionada ao scouting de futebol.