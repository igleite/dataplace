/*==============================================================
    Author       : Igor Leite
    Create date  : 2025-08-08
    Description  : Retorna a hierarquia do plano de contas, 
                   formatando a descrição de forma indentada 
                   conforme o nível.
    
    Example Output:
    -----------------------------------------------------------
    numcta   DescricaoHierarquica
    -------  --------------------------------
    1        Conta Raiz
    2        |---- Conta Filha
    3        |---- |---- Conta Neta
    4        |---- Outra Conta Filha
==============================================================*/
DECLARE @ano INT, @cdEmpresa NCHAR(5), @classificacao NCHAR(30)--, @numcta NCHAR(7)
SET @ano = 2024
SET @cdEmpresa = 'TRE'
SET @classificacao = '3.0.00.00.00.0000             '

;WITH Hierarquia AS (
                    SELECT PlanoContas.numcta
                         , PlanoContas.Descricao
                         , PlanoContas.Nivel
                         , PlanoContas.numctapai
                    FROM PlanoContas
                    WHERE 1 = 1
                      AND PlanoContas.ano = @ano
                      AND PlanoContas.cdempresa = @cdEmpresa
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
                      AND p.cdempresa = @cdEmpresa
                    )
SELECT H.numcta
     , REPLICATE('|---- ', H.Nivel - 1) + TRIM(H.Descricao) AS DescricaoHierarquica
FROM Hierarquia H
WHERE 1 = 1


/*==============================================================
    Author       : Andre Mariano
    Create date  : 2025-08-08
    Description  : Função que retorna a hierarquia do plano 
                   de contas com descrição indentada conforme 
                   o nível de cada conta.
    
    Example Output:
    -----------------------------------------------------------
    numcta   DescricaoHierarquica
    -------  --------------------------------
    1        Conta Raiz
    2        |---- Conta Filha
    3        |---- |---- Conta Neta
    4        |---- Outra Conta Filha
==============================================================*/
IF EXISTS
    (SELECT name
     FROM sys.sysobjects
     WHERE name = 'SYM_PlanoContasBuscaHierarquia')
    DROP FUNCTION SYM_PlanoContasBuscaHierarquia;
GO

CREATE FUNCTION [dbo].[SYM_PlanoContasBuscaHierarquia](
    @CdEmpresa NVARCHAR(5),
    @Classificacao NVARCHAR(MAX),
    @ano INT
)
    RETURNS TABLE
        AS
        RETURN(
            -- Quebra a lista de classificações separadas por vírgula
            WITH SplitClassificacao AS (
                                        SELECT LTRIM(RTRIM(m.n.value('.', 'NVARCHAR(30)'))) AS valor
                                        FROM (SELECT CAST('<x>' + REPLACE(@Classificacao, ',', '</x><x>') + '</x>' AS XML) AS xmlData) AS t CROSS APPLY t.xmlData.nodes('/x') AS m(n)
                                        )
                 , Hierarquia AS (
                                SELECT p.numcta
                                     , p.Descricao
                                     , p.Nivel
                                     , p.numctapai
                                FROM PlanoContas p
                                WHERE 1 = 1
                                  AND p.ano = @ano
                                  AND p.cdempresa = TRIM(@CdEmpresa)
                                  AND TRIM(p.classificacao) IN (SELECT valor FROM SplitClassificacao)
    
                                UNION ALL
    
                                SELECT p.numcta
                                     , p.Descricao
                                     , p.Nivel
                                     , p.numctapai
                                FROM PlanoContas p
                                         INNER JOIN Hierarquia h ON p.numctapai = h.numcta
                                WHERE 1 = 1
                                  AND p.cdempresa = TRIM(@CdEmpresa)
                                  AND p.ano = @ano
                                )
            SELECT H.numcta,
                   REPLICATE('|---- ', H.Nivel - 1) + TRIM(H.Descricao) AS DescricaoHierarquica
            FROM Hierarquia H
        )
GO

SELECT *
FROM SYM_PlanoContasBuscaHierarquia('TRE  ', '1.0.00.000.0000               ', 2014)
