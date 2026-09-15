-- 1. Jogadores por campeonato
SELECT
    Campeonato,
    COUNT(*) AS Total_Jogadores
FROM dbo.Jogadores
GROUP BY Campeonato
ORDER BY Total_Jogadores DESC;

-- 2. Jogadores por posição
SELECT
    Posicao,
    COUNT(*) AS Total_Jogadores
FROM dbo.Jogadores
GROUP BY Posicao
ORDER BY Total_Jogadores DESC;

-- 3. Jogadores por estado de origem
SELECT
    Estado_Origem,
    COUNT(*) AS Total_Jogadores
FROM dbo.Jogadores
WHERE Estado_Origem IS NOT NULL
  AND Estado_Origem <> ''
GROUP BY Estado_Origem
ORDER BY Total_Jogadores DESC;


-- 4. Jogadores que já trabalharam com o Cristiano
SELECT
    Nome,
    Posicao,
    Clube_Atual,
    Clube_Comigo,
    Campeonato,
    Pe_Dominante,
    Altura_cm
FROM dbo.Jogadores
WHERE Trabalhou_Comigo = 'SIM'
ORDER BY Nome;


-- 5. Exemplo de filtro de recrutamento
SELECT
    Nome,
    Posicao,
    Altura_cm,
    Clube_Atual,
    Campeonato,
    Estado_Origem,
    Pe_Dominante
FROM dbo.Jogadores
WHERE Posicao = 'VOLANTE'
  AND Pe_Dominante = 'ESQUERDO'
ORDER BY Altura_cm DESC;


-- 6. Quantidade por clube atual
SELECT
    Clube_Atual,
    COUNT(*) AS Total_Jogadores
FROM dbo.Jogadores
GROUP BY Clube_Atual
ORDER BY Total_Jogadores DESC;

----------------------------------------------------------------------------------------

CREATE VIEW vw_Jogadores_Scout AS
SELECT
    ID_Jogador,
    Nome,
    Data_Nascimento,
    Posicao,
    Altura_cm,
    Clube_Atual,
    Estado_Origem,
    Campeonato,
    Pais,
    Pe_Dominante,
    Clube_Comigo,
    Trabalhou_Comigo,
    Link_Ogol,
    Link_Foto
FROM dbo.Jogadores;




