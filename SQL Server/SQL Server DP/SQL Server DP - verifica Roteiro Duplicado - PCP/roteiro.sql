DECLARE @cdProduto NVARCHAR(20)

SET @cdProduto = '761FSS04'

SELECT *
FROM ProdutoCelProdutiva PCP
WHERE PCP.CdProduto IN (@cdProduto)
  AND PCP.OperacaoNome IN (
    SELECT OperacaoNome
    FROM ProdutoCelProdutiva
    WHERE CdProduto IN (@cdProduto)
    GROUP BY OperacaoNome
    HAVING COUNT(*) > 1
)
