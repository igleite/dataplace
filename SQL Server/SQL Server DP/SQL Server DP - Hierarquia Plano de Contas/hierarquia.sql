DECLARE @ano INT, @cdEmpresa NCHAR(5), @classificacao NCHAR(30), @numcta NCHAR(7)
SET @ano = 2024
SET @classificacao = '3.0.00.00.00.0000             '

;WITH Hierarquia AS (
     SELECT PlanoContas.numcta
       , PlanoContas.Descricao
                         , PlanoContas.Nivel
       , PlanoContas.numctapai
                    FROM PlanoContas
                    WHERE 1 = 1
      AND PlanoContas.ano = @ano
      AND TRIM(PlanoContas.classificacao) = TRIM(@classificacao)
                    UNION ALL

                    SELECT p.numcta
       , p.Descricao
                         , p.Nivel
                         , p.numctapai
                    FROM PlanoContas p
                             INNER JOIN Hierarquia h ON p.numctapai = h.numcta
     WHERE 1 = 1
      AND p.ano = @ano
       )
SELECT H.numcta
     , REPLICATE('|---- ', H.Nivel - 1) + TRIM(H.Descricao) AS DescricaoHierarquica
FROM Hierarquia H
WHERE 1 = 1